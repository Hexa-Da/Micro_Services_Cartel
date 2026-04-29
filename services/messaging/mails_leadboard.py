#!/usr/bin/env python3
"""Envoie à chaque participant un e-mail avec ses résultats individuels du classement (leadboard).

Lit leadboard.txt (format « CLASSEMENT DES PARIEURS ») et joint sur un CSV contenant au minimum
une colonne e-mail et une colonne bracelet (Code barres, etc.), comme mails_csv.py.

Enrichit avec le classement des délégations par précision (fichier type
classement_delegations_precision.txt) : rang et métriques de la délégation du parieur.

Si le CSV contient des UID NFC (colonne Code barres) et l’Excel Wilout (UID → Short tag)
est fourni, la clé de jointure avec le leadboard est le short tag — comme pour mails_csv.py.

Par défaut, joint à chaque message le leadboard et le classement délégations (fichiers texte).
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from email.utils import parseaddr
from pathlib import Path

# Réutilise chargement .env, envoi SMTP et détection colonnes depuis mails_csv
from mails_csv import (
    display_prenom,
    find_email_field,
    load_env_file,
    load_uid_to_short,
    looks_like_nfc_uid,
    norm_uid,
    pick_field,
    row_for_templates,
    send_one,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]


LEADBOARD_MAIL_SUBJECT = "Cartel Nancy 2026 — vos résultats aux paris"

LEADBOARD_MAIL_BODY = """Bonjour {prenom} {nom},

Voici le récapitulatif de vos paris pour le Cartel Nancy 2026 :

  Rang : {rang} / {nbr_total_de_parieur}
  Bons paris : {bons} / {paries}
  Taux de réussite : {pct}
  Délégation : {delegation}
  Classement de la délégation : rang {rang_delegation} / {nbr_total_delegations} — {pct_delegation} (votes {delegation_bons_total})

J'espère que l'application vous aura plu !

Merci pour votre participation !

Le Dev de l'App
"""


def norm_delegation(name: str) -> str:
    """Normalise le libellé délégation pour jointure avec le fichier de classement."""
    return " ".join((name or "").split())


# Délégations Cartel (suffixe de ligne leadboard après Nom / Prénom) — ordre géré par le matcher.
DEFAULT_DELEGATION_NAMES: frozenset[str] = frozenset(
    {
        "Mines Alès",
        "Mines Paris",
        "Mines Sainté",
        "Télécom Paris",
        "IMT A",
        "Mines Albi",
        "Eurecom",
        "ENSIC",
        "Mines Nancy",
        "TSP",
        "IMT NE",
        "Télécom Nancy",
        "Alumni",
        "Ponts",
        "ENSAIA",
    }
)


def known_delegations_for_leadboard_parsing(classement_path: Path) -> frozenset[str]:
    """Libellés pour découper la ligne leadboard : délégation = suffixe connu, le reste = nom + prénom."""
    names = set(DEFAULT_DELEGATION_NAMES)
    if classement_path.is_file():
        by, _ = parse_classement_delegations_precision(classement_path)
        names.update(by.keys())
    return frozenset(names)


def _token_eq(a: str, b: str) -> bool:
    return a.casefold() == b.casefold()


def split_middle_nom_prenom_delegation(
    middle: list[str],
    known_delegations: frozenset[str],
) -> tuple[str, str, str]:
    """Découpe [ … nom …, prénom, délégation ] en reconnaissant la délégation par suffixe."""
    if len(middle) < 3:
        raise ValueError("middle trop court")
    ranked = sorted(
        known_delegations,
        key=lambda d: (-len(d.split()), -len(d), d.casefold()),
    )
    for d in ranked:
        dtoks = d.split()
        n_tok = len(dtoks)
        if n_tok > len(middle):
            continue
        tail = middle[-n_tok:]
        if all(_token_eq(tail[i], dtoks[i]) for i in range(n_tok)):
            rest = middle[:-n_tok]
            if len(rest) < 1:
                continue
            if len(rest) == 1:
                return rest[0], "", d
            prenom = rest[-1]
            nom = " ".join(rest[:-1])
            return nom, prenom, d
    nom = middle[0]
    prenom = middle[1]
    delegation = " ".join(middle[2:])
    return nom, prenom, delegation


def parse_classement_delegations_precision(
    path: Path,
) -> tuple[dict[str, dict[str, str]], str]:
    """Parse classement_delegations_precision.txt ; clé = délégation normalisée.

    Chaque entrée contient : rang_delegation, pct_delegation, delegation_bons_total,
    delegation_sports (nombre de sports avec résultat pour cette délégation).
    Retourne aussi nbr_total_delegations (nombre de délégations classées).
    """
    if not path.is_file():
        return {}, "—"

    text = path.read_text(encoding="utf-8")
    out: dict[str, dict[str, str]] = {}

    for raw in text.splitlines():
        line = raw.strip()
        if (
            not line
            or line.startswith("-")
            or "CLASSEMENT DES" in line
            or "Métrique" in line
            or ("Rang" in line and "Délégation" in line)
        ):
            continue
        parts = line.split()
        if len(parts) < 5:
            continue
        try:
            rang = int(parts[0])
        except ValueError:
            continue
        if not parts[-3].endswith("%"):
            continue
        if "/" not in parts[-2]:
            continue
        if not parts[-1].isdigit():
            continue
        delegation = " ".join(parts[1:-3])
        if not delegation:
            continue
        key = norm_delegation(delegation)
        out[key] = {
            "rang_delegation": str(rang),
            "pct_delegation": parts[-3],
            "delegation_bons_total": parts[-2],
            "delegation_sports": parts[-1],
        }

    total = str(len(out)) if out else "—"
    return out, total


def apply_delegation_classement(
    ctx: dict[str, str],
    by_delegation: dict[str, dict[str, str]],
    nbr_total_delegations: str,
) -> bool:
    """Ajoute au contexte les champs rang_delegation, pct_delegation, etc.

    Retourne False si la délégation n’a pas été trouvée dans le classement (ou nom vide).
    """
    name = norm_delegation(ctx.get("delegation", ""))
    if name and name in by_delegation:
        ctx.update(by_delegation[name])
        ctx["nbr_total_delegations"] = nbr_total_delegations
        return True
    missing = "—"
    ctx["rang_delegation"] = missing
    ctx["pct_delegation"] = missing
    ctx["delegation_bons_total"] = missing
    ctx["delegation_sports"] = missing
    ctx["nbr_total_delegations"] = nbr_total_delegations or missing
    return False


def detect_delimiter(first_line: str) -> str:
    if first_line.count(";") > first_line.count(","):
        return ";"
    return ","


def parse_leadboard(
    path: Path,
    known_delegations: frozenset[str] | None = None,
) -> dict[str, dict[str, str]]:
    """Parse leadboard.txt ; clé = code bracelet (short tag ex. ND5X54, ou numérique legacy).

    Délégation : suffixe reconnu parmi known_delegations (noms composés « Le … » / prénoms
    multiples gérés). Si inconnu, repli sur l’ancien découpage (1er / 2e token / reste).
    """
    kd = known_delegations or DEFAULT_DELEGATION_NAMES
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    out: dict[str, dict[str, str]] = {}

    for raw in lines:
        line = raw.strip()
        if (
            not line
            or line.startswith("-")
            or "CLASSEMENT" in line
            or ("Rang" in line and "Bracelet" in line and "Nom" in line)
        ):
            continue
        parts = line.split()
        if len(parts) < 8:
            continue
        try:
            rang = int(parts[0])
        except ValueError:
            continue
        bracelet_raw = parts[1]
        if bracelet_raw.isdigit():
            key = str(int(bracelet_raw))
        else:
            key = bracelet_raw
        pct = parts[-1]
        if not pct.endswith("%"):
            continue
        try:
            paries = int(parts[-2])
            bons = int(parts[-3])
        except ValueError:
            continue
        middle = parts[2:-3]
        if len(middle) < 3:
            continue
        nom, prenom, delegation = split_middle_nom_prenom_delegation(middle, kd)
        out[key] = {
            "rang": str(rang),
            "bracelet": key,
            "nom": nom,
            "prenom": prenom,
            "delegation": delegation,
            "bons": str(bons),
            "paries": str(paries),
            "pct": pct,
        }
    total = str(len(out))
    for row in out.values():
        row["nbr_total_de_parieur"] = total
    return out


def display_nom(raw: str) -> str:
    """Export billetterie « Nom - Ville » (ex. Aguer - Albi) → nom seul pour la salutation."""
    raw = raw.strip()
    if " - " in raw:
        return raw.split(" - ", 1)[0].strip()
    return raw


def bracelet_key_from_row(
    row: dict[str, str],
    uid_to_short: dict[str, str] | None,
    lineno: int,
    to_email: str,
) -> str:
    """Clé pour jointure leadboard : Short tag Wilout si UID NFC + Excel, sinon code CSV."""
    code = pick_field(
        row,
        "braceletNumber",
        "Code bracelet",
        "Code barres",
        "Code_barres",
        "code_barres",
        "bracelet",
        "code",
        "Bracelet",
    )
    s = (code or "").strip()
    if not s:
        return ""
    if s.upper().startswith("OR-"):
        return s
    if uid_to_short:
        n = norm_uid(s)
        if n in uid_to_short:
            return uid_to_short[n]
        if looks_like_nfc_uid(s):
            print(
                f"Ligne {lineno}: UID {s!r} sans Short tag dans l'Excel ({to_email}).",
                file=sys.stderr,
            )
            return s
    try:
        return str(int(s.lstrip("0") or "0"))
    except ValueError:
        return s


def row_ctx_for_mail(
    row: dict[str, str], stats: dict[str, str]
) -> dict[str, str]:
    """Fusionne CSV + stats leadboard pour le template."""
    base = row_for_templates(row)
    base.update(stats)
    prenom_raw = pick_field(row, "Prénom", "Prenom", "prenom", "firstname", "First name")
    if prenom_raw:
        base["prenom"] = display_prenom(prenom_raw)
    nom_raw = pick_field(
        row, "Nom", "nom", "Nom de famille", "lastname", "Last name"
    )
    if nom_raw:
        base["nom"] = display_nom(nom_raw)
    return base


def build_body(template: str, row: dict[str, str]) -> str:
    try:
        return template.format(**row)
    except KeyError as e:
        sys.exit(
            f"Le modèle référence une clé absente : {e.args[0]!r}.\n"
            f"Clés disponibles : {', '.join(sorted(row.keys()))}"
        )


def main() -> None:
    load_env_file()
    parser = argparse.ArgumentParser(
        description="Envoie un e-mail par participant avec ses résultats (leadboard + CSV e-mails)."
    )
    parser.add_argument(
        "csv_path",
        type=Path,
        help="CSV (UTF-8) avec e-mail et colonne bracelet / Code barres.",
    )
    parser.add_argument(
        "leadboard_path",
        type=Path,
        nargs="?",
        default=PROJECT_ROOT / "data" / "output" / "leadboard.txt",
        help="Fichier leadboard (défaut: data/output/leadboard.txt).",
    )
    parser.add_argument(
        "--delegations-precision",
        type=Path,
        default=PROJECT_ROOT / "data" / "output" / "classement_delegations_precision.txt",
        help="Classement des délégations par précision (défaut: data/output/classement_delegations_precision.txt).",
    )
    parser.add_argument(
        "--no-delegation-classement",
        action="store_true",
        help="Ne pas utiliser le fichier de classement des délégations (placeholders —).",
    )
    parser.add_argument(
        "--subject",
        default=os.environ.get("LEADBOARD_MAIL_SUBJECT", LEADBOARD_MAIL_SUBJECT),
        help="Sujet du mail (placeholders {rang}, {nbr_total_de_parieur}, {prenom}, …).",
    )
    parser.add_argument(
        "--body-template",
        default=os.environ.get("LEADBOARD_MAIL_BODY", LEADBOARD_MAIL_BODY),
        help="Corps du message (placeholders voir modèle par défaut).",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Pause en secondes entre chaque envoi.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="N'envoie pas : affiche les envois prévus.",
    )
    parser.add_argument(
        "--no-attachments",
        action="store_true",
        help="Ne pas joindre leadboard.txt et classement_delegations_precision.txt.",
    )
    parser.add_argument(
        "--uid-xlsx",
        type=Path,
        default=PROJECT_ROOT
        / "data"
        / "input"
        / "WILOUT X CARTEL NANCY 2515_liste UID.xlsx",
        help="Excel Wilout UID → Short tag (défaut: data/input/WILOUT X CARTEL NANCY 2515_liste UID.xlsx).",
    )
    parser.add_argument(
        "--no-uid-short-map",
        action="store_true",
        help="Ne pas charger l'Excel : pas de conversion UID NFC → short tag.",
    )
    parser.add_argument(
        "--start-line",
        type=int,
        metavar="N",
        help="Ne traiter qu'à partir de la ligne N du CSV (même numéro que « Ligne N » ; ligne 1 = en-tête).",
    )
    args = parser.parse_args()

    if args.start_line is not None and args.start_line < 1:
        sys.exit("--start-line doit être >= 1 (1 = première ligne du fichier, en-tête CSV).")

    if not args.leadboard_path.is_file():
        sys.exit(f"Fichier leadboard introuvable : {args.leadboard_path}")

    uid_to_short: dict[str, str] | None = None
    if not args.no_uid_short_map:
        if not args.uid_xlsx.is_file():
            print(
                f"Avertissement : Excel Wilout introuvable ({args.uid_xlsx}), "
                "conversion UID → short tag désactivée.",
                file=sys.stderr,
            )
        else:
            try:
                uid_to_short = load_uid_to_short(args.uid_xlsx)
            except RuntimeError as e:
                print(f"{e} Conversion UID désactivée.", file=sys.stderr)
            except Exception as e:
                print(
                    f"Erreur lecture {args.uid_xlsx} : {e}. Conversion UID désactivée.",
                    file=sys.stderr,
                )

    known_deleg = known_delegations_for_leadboard_parsing(args.delegations_precision)
    stats_by_bracelet = parse_leadboard(args.leadboard_path, known_deleg)
    if not stats_by_bracelet:
        sys.exit(
            "Aucune ligne de classement parsée dans le leadboard. "
            "Vérifiez le format du fichier."
        )

    if args.no_delegation_classement:
        by_delegation: dict[str, dict[str, str]] = {}
        nbr_total_delegations = "—"
    else:
        by_delegation, nbr_total_delegations = parse_classement_delegations_precision(
            args.delegations_precision
        )
        if not args.delegations_precision.is_file():
            print(
                f"Avertissement : fichier délégations introuvable ({args.delegations_precision}), "
                "classement par délégation laissé vide (—).",
                file=sys.stderr,
            )
        elif not by_delegation:
            print(
                f"Avertissement : aucune délégation parsée dans {args.delegations_precision}.",
                file=sys.stderr,
            )

    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER", "")
    smtp_password = os.environ.get("SMTP_PASSWORD", "")
    smtp_ssl_flag = os.environ.get("SMTP_SSL", "").strip().lower() in ("1", "true", "yes")
    smtp_use_ssl = smtp_port == 465 or smtp_ssl_flag

    if not args.dry_run:
        if not smtp_user or not smtp_password:
            sys.exit(
                "Définissez SMTP_USER et SMTP_PASSWORD dans .env ou l'environnement.\n"
                "Ou utilisez --dry-run pour simuler."
            )

    if args.no_attachments:
        mail_attachments: list[tuple[str, bytes]] | None = None
    else:
        pj: list[tuple[str, bytes]] = []
        if args.leadboard_path.is_file():
            pj.append((args.leadboard_path.name, args.leadboard_path.read_bytes()))
        if args.delegations_precision.is_file():
            pj.append(
                (args.delegations_precision.name, args.delegations_precision.read_bytes())
            )
        mail_attachments = pj if pj else None

    with args.csv_path.open(newline="", encoding="utf-8-sig") as f:
        first_line = f.readline()
        if not first_line:
            sys.exit("Fichier CSV vide.")
        delimiter = detect_delimiter(first_line)
        f.seek(0)
        reader = csv.DictReader(f, delimiter=delimiter)
        email_key = find_email_field(reader.fieldnames)

        sent = 0
        failed = 0
        skipped_no_bracelet = 0
        skipped_no_stats = 0
        bracelets_seen: set[str] = set()
        warned_missing_delegations: set[str] = set()

        if args.start_line is not None:
            print(
                f"Reprise : lignes < {args.start_line} ignorées (même numérotation que « Ligne N »).",
                file=sys.stderr,
            )

        for lineno, row in enumerate(reader, start=2):
            if args.start_line is not None and lineno < args.start_line:
                continue
            raw_email = (row.get(email_key) or "").strip()
            if not raw_email:
                print(f"Ligne {lineno}: e-mail vide, ignorée.", file=sys.stderr)
                failed += 1
                continue
            _, addr = parseaddr(raw_email)
            to_email = addr or raw_email
            if "@" not in to_email:
                print(
                    f"Ligne {lineno}: adresse invalide {raw_email!r}, ignorée.",
                    file=sys.stderr,
                )
                failed += 1
                continue

            bkey = bracelet_key_from_row(row, uid_to_short, lineno, to_email)
            if not bkey:
                print(
                    f"Ligne {lineno}: bracelet / Code barres vide pour {to_email}, ignorée.",
                    file=sys.stderr,
                )
                skipped_no_bracelet += 1
                continue

            stats = stats_by_bracelet.get(bkey)
            if not stats:
                print(
                    f"Ligne {lineno}: bracelet {bkey!r} absent du leadboard pour {to_email}, ignorée.",
                    file=sys.stderr,
                )
                skipped_no_stats += 1
                continue

            bracelets_seen.add(bkey)
            ctx = row_ctx_for_mail(row, stats)
            found_deleg = apply_delegation_classement(
                ctx, by_delegation, nbr_total_delegations
            )
            dname = norm_delegation(ctx.get("delegation", ""))
            if (
                not found_deleg
                and dname
                and by_delegation
                and dname not in warned_missing_delegations
            ):
                warned_missing_delegations.add(dname)
                print(
                    f"Délégation {dname!r} absente de {args.delegations_precision}, "
                    "classement délégation laissé vide (—).",
                    file=sys.stderr,
                )
            subject = args.subject if "{" not in args.subject else build_body(args.subject, ctx)
            body = build_body(args.body_template, ctx)

            try:
                send_one(
                    smtp_host,
                    smtp_port,
                    smtp_user,
                    smtp_password,
                    to_email,
                    subject,
                    body,
                    args.dry_run,
                    smtp_use_ssl,
                    mail_attachments,
                )
                sent += 1
                print(f"OK {to_email} | rang {ctx['rang']}")
            except Exception as e:
                failed += 1
                print(f"Ligne {lineno} {to_email}: erreur {e}", file=sys.stderr)

            if args.delay > 0:
                time.sleep(args.delay)

    missing = set(stats_by_bracelet.keys()) - bracelets_seen
    if missing:

        def _missing_sort_key(x: str) -> tuple[int, str]:
            try:
                return (0, f"{int(x):020d}")
            except ValueError:
                return (1, x)

        sample = sorted(missing, key=_missing_sort_key)[:5]
        print(
            f"Avertissement : {len(missing)} bracelet(s) dans le leadboard sans ligne CSV correspondante "
            f"(ex. : {', '.join(sample)}{'…' if len(missing) > 5 else ''}).",
            file=sys.stderr,
        )

    print(
        f"Terminé : {sent} envoyé(s), {failed} erreur(s) ou ignoré(s) e-mail, "
        f"{skipped_no_bracelet} sans bracelet, {skipped_no_stats} sans entrée leadboard."
    )


if __name__ == "__main__":
    main()
