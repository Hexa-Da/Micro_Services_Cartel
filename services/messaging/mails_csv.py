#!/usr/bin/env python3
"""Envoie un mail par ligne depuis un CSV. Colonnes attendues : au minimum une colonne email.

Avec --preset bracelet, la colonne « Code barres » contient en pratique l’UID NFC ; le script
peut le remplacer par le Short tag Wilout (Excel UID → Short tag) pour le corps du mail.

Avec --preset guide, mail d’information (guide / dépistage Atoutbio aux Aiguillettes) : pas
d’obligation de code bracelet ; toutes les lignes avec un email valide sont envoyées.
L’Excel Wilout (UID → Short tag) n’est pas chargé pour ce preset.

Avec --preset relance, même base que --preset bracelet avec une phrase d’introduction
supplémentaire demandant de mettre à jour l’application si déjà installée.

Avec --preset memento, mail d'information envoyé à tous les participant·es

Utilisez --attach FICHIER (répétable) pour joindre un ou plusieurs fichiers à chaque envoi.
"""

from __future__ import annotations

import argparse
import csv
import os
import smtplib
import sys
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.policy import SMTP as SMTP_POLICY
from email.utils import formataddr, parseaddr

# En-têtes + adresses EAI (ex. Jérémy@…) : politique UTF-8, sinon as_bytes échoue en ASCII.
_SMTP_WIRE_POLICY = SMTP_POLICY.clone(utf8=True)
from pathlib import Path

try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore

def load_env_file() -> None:
    """Charge un fichier .env (même dossier que le script, puis répertoire courant).
    Ne remplace pas une variable déjà définie dans l'environnement."""
    for base in (Path(__file__).resolve().parent, Path.cwd()):
        path = base / ".env"
        if not path.is_file():
            continue
        with path.open(encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                if not key:
                    continue
                value = value.strip()
                if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                    value = value[1:-1]
                if key == "SMTP_PASSWORD":
                    value = value.replace(" ", "").replace("\t", "")
                if key not in os.environ:
                    os.environ[key] = value
        return


# Envoi code bracelet. Dans les exports billetterie, la colonne « Code barres »
# est le même identifiant que le code bracelet dans l’app (souvent rempli plus tard).
# Placeholders si code connu : prenom, nom, braceletNumber (alias CSV : Prénom, Nom, Code barres).

BRACELET_MAIL_SUBJECT = (
    "Cartel Nancy 2026 — votre code d’accès à l’application"
)

BRACELET_MAIL_BODY = """Bonjour {prenom} {nom},

Votre code d'accès personnel est : {braceletNumber}

Ce code correspond au code noté au dos du bracelet qui vous sera donné pour participer au Cartel. Il vous permet de vous connecter à l’application mobile du Cartel Nancy 2026, disponible sur l’App Store et sur le Google Play Store. Une fois connecté·e, vous aurez accès aux paris et au contenu réservé aux participant·es :

- Vous retrouvez sur la première page les matchs les plus intéressants selon les préférences que vous aurez mises dans les paramètres.
- Sur la page principale, des onglets, « événements » et « calendrier » liste tous les matchs et soirées, n'hésitez pas à les filtrer !
- Vous pouvez également zoomer sur la carte pour voir les arrêts de bus et les indications de l'organisation afin de naviguer dans Nancy.
- Les cartes des soirées sont aussi disponibles sur les marqueurs concernés afin de vous guider dans les événements culturels.
- Pour finir toutes les infos complémentaires seront dans la dernière page sous forme de FAQ ou de fichier.

Merci de ne pas partager votre code d'accès, il est à usage unique et personnel !

À très bientôt,
L’équipe Cartel Nancy 2026
"""

# Quand les codes ne sont pas encore dans le CSV (colonne Code barres vide).

BRACELET_MAIL_SUBJECT_PENDING = (
    "Cartel Nancy 2026 — votre code d’accès à l’application"
)

BRACELET_MAIL_BODY_PENDING = """Bonjour {prenom} {nom},

L’application mobile du Cartel Nancy 2026 est disponible sur l’App Store et sur le Google Play Store.

À très bientôt,
L’équipe Cartel Nancy 2026
"""

RELANCE_MAIL_SUBJECT = BRACELET_MAIL_SUBJECT

RELANCE_MAIL_BODY = """Bonjour {prenom} {nom},

Si vous avez déjà installé l'application Cartel Nancy 2026, pensez à la METTRE À JOUR (faite le vraiment c'est le dev qui vous parle, des bugs ont été corrigés) et n'oubliez pas de faire vos paris avant le debut du Cartel (Jeudi 15h00) !

Autrement, votre code d'accès personnel est : {braceletNumber}

Ce code correspond au code noté au dos du bracelet qui vous sera donné pour participer au Cartel. Il vous permet de vous connecter à l’application mobile du Cartel Nancy 2026, disponible sur l’App Store et sur le Google Play Store. Une fois connecté·e, vous aurez accès aux paris et au contenu réservé aux participant·es :

- Vous retrouvez sur la première page les matchs les plus intéressants selon les préférences que vous aurez mises dans les paramètres.
- Sur la page principale, des onglets, « événements » et « calendrier » liste tous les matchs et soirées, n'hésitez pas à les filtrer !
- Vous pouvez également zoomer sur la carte pour voir les arrêts de bus et les indications de l'organisation afin de naviguer dans Nancy.
- Les cartes des soirées sont aussi disponibles sur les marqueurs concernés afin de vous guider dans les événements culturels.
- Pour finir toutes les infos complémentaires seront dans la dernière page sous forme de FAQ ou de fichier.

Merci de ne pas partager votre code d'accès, il est à usage unique et personnel !

Si un de vos amis n'a pas reçu son code, vous pouvez me contacter en répondant à ce mail avec nom, prénom et délégation de votre ami.

Bonne route !
L’équipe Cartel Nancy 2026
"""

GUIDE_MAIL_SUBJECT = (
    "Cartel Nancy 2026 — guide et dépistage"
)

GUIDE_MAIL_BODY = """Bonjour {prenom} {nom},

Veuillez trouver en pièce jointe le guide du Cartel, contenant toutes les informations pratiques dont vous aurez besoin pour l'événement.

Nous souhaitions aussi vous informer de la présence d’un stand Atoutbio aux Aiguillettes, où vous pourrez vous faire dépister pour le VIH et les IST. C’est sans ordonnance, sans frais pour les personnes de moins de 26 ans, et c’est toute la journée vendredi et samedi !

À très bientôt,
L’équipe Cartel Nancy 2026
"""

MEMENTO_MAIL_SUBJECT = (
    "Cartel Nancy 2026 — mémento disponibile"
)

MEMENTO_MAIL_BODY = """Bonjour {prenom} {nom},

Nous avons le plaisir de vous informer que toutes les photos de l’événement sont désormais disponibles en ligne !

Vous pouvez y accéder via le lien suivant : https://memento.photo/j/W_APc945

Les photos sont organisées par dossiers, en fonction des différents événements et sports, afin de faciliter votre navigation.

Un système de reconnaissance faciale a également été mis en place : il vous permet de retrouver automatiquement, dans un même dossier, toutes les photos sur lesquelles vous apparaissez.

Si vous avez pris des photos et souhaitez les ajouter au site, n’hésitez pas à nous en faire la demande par mail à cartelnancy2026@gmail.com : nous vous expliquerons la marche à suivre une fois votre adresse ajoutée en tant que photographe.

Nous espérons que vous prendrez plaisir à revivre ces moments !

À très bientôt,  
L’équipe Cartel Nancy 2026
"""


EMAIL_COLUMN_CANDIDATES = (
    "email",
    "e-mail",
    "mail",
    "courriel",
    "adresse email",
    "adresse_email",
)


def detect_delimiter(first_line: str) -> str:
    if first_line.count(";") > first_line.count(","):
        return ";"
    return ","


def pick_field(row: dict[str, str], *header_names: str) -> str:
    """Première valeur non vide pour l’un des en-têtes possibles (sensible à la casse puis insensible)."""
    for name in header_names:
        if name in row:
            v = (row[name] or "").strip()
            if v:
                return v
    lower_to_key = {k.lower(): k for k in row}
    for name in header_names:
        lk = name.lower().strip()
        if lk in lower_to_key:
            v = (row[lower_to_key[lk]] or "").strip()
            if v:
                return v
    return ""


def display_prenom(raw: str) -> str:
    """Ex. export billetterie « Prénom - Ville » → retourne prénom seul"""
    raw = raw.strip()
    if " - " in raw:
        return raw.split(" - ", 1)[0].strip()
    return raw


def row_for_templates(row: dict[str, str]) -> dict[str, str]:
    """Ajoute des clés canoniques (prenom, nom, braceletNumber) pour les modèles.
    braceletNumber est alimenté par « Code barres » ou colonnes équivalentes (= code bracelet app)."""
    out = {k: (v or "").strip() for k, v in row.items()}
    prenom_raw = pick_field(row, "Prénom", "Prenom", "prenom", "firstname", "First name")
    nom_raw = pick_field(row, "Nom", "nom", "Nom de famille", "lastname", "Last name")
    code = pick_field(row, "braceletNumber", "Code bracelet", "Code barres", "Code_barres", "code_barres", "Code bracelet", "bracelet", "code")
    out["prenom"] = display_prenom(prenom_raw) if prenom_raw else out.get("prenom", "")
    out["nom"] = nom_raw or out.get("nom", "")
    out["braceletNumber"] = code or out.get("braceletNumber", "")
    return out


def norm_uid(s: str) -> str:
    """Aligne un UID lu dans le CSV avec les cellules Excel (majuscules, sans zéros de tête)."""
    return s.upper().strip().lstrip("0") or "0"


def load_uid_to_short(xlsx_path: Path) -> dict[str, str]:
    """Excel Wilout : colonnes UID, Short tag."""
    if openpyxl is None:
        raise RuntimeError("openpyxl requis : pip install openpyxl")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb.active
    out: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uid, short = row[0], row[1]
        if uid is None or str(uid).strip() == "":
            continue
        uid_s = str(uid).strip().upper()
        short_s = str(short).strip() if short else ""
        if not short_s:
            continue
        out[norm_uid(uid_s)] = short_s
    wb.close()
    return out


def looks_like_nfc_uid(code: str) -> bool:
    """Heuristique : chaîne longue en hex (UID NFC), pas un short tag déjà présent."""
    s = code.strip().upper()
    if len(s) < 12:
        return False
    return all(c in "0123456789ABCDEF" for c in s)


def resolve_bracelet_to_short_tag(
    row_ctx: dict[str, str],
    uid_to_short: dict[str, str] | None,
    lineno: int,
    to_email: str,
) -> None:
    """Remplace braceletNumber (UID issu du CSV) par le Short tag Wilout si trouvé dans l’Excel."""
    if not uid_to_short:
        return
    code = (row_ctx.get("braceletNumber") or "").strip()
    if not code:
        return
    if code.startswith("OR-"):
        return
    n = norm_uid(code)
    if n in uid_to_short:
        row_ctx["braceletNumber"] = uid_to_short[n]
        return
    if looks_like_nfc_uid(code):
        print(
            f"Ligne {lineno}: UID {code!r} sans Short tag dans l’Excel, code inchangé ({to_email}).",
            file=sys.stderr,
        )


def find_email_field(fieldnames: list[str] | None) -> str:
    if not fieldnames:
        sys.exit("Le CSV n'a pas d'en-tête de colonnes.")

    lower_map = {name.lower().strip(): name for name in fieldnames}

    for candidate in EMAIL_COLUMN_CANDIDATES:
        if candidate in lower_map:
            return lower_map[candidate]

    for name in fieldnames:
        n = name.lower().strip()
        if "email" in n or n in ("mail", "e-mail"):
            return name

    sys.exit(
        "Aucune colonne email trouvée. Ajoutez une colonne nommée par ex. « email ».\n"
        f"Colonnes détectées : {', '.join(fieldnames)}"
    )


def build_body(template: str, row: dict[str, str]) -> str:
    try:
        return template.format(**row)
    except KeyError as e:
        sys.exit(
            f"Le modèle de corps référence une colonne absente du CSV : {e.args[0]!r}.\n"
            f"Colonnes disponibles : {', '.join(row.keys())}"
        )


def send_one(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    to_email: str,
    subject: str,
    body: str,
    dry_run: bool,
    smtp_use_ssl: bool,
    attachments: list[tuple[str, bytes]] | None = None,
) -> None:
    attachments = attachments or []
    if not attachments:
        msg: MIMEText | MIMEMultipart = MIMEText(body, "plain", "utf-8")
    else:
        msg = MIMEMultipart()
        msg.attach(MIMEText(body, "plain", "utf-8"))
        for filename, raw in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(raw)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=filename)
            msg.attach(part)

    msg["Subject"] = subject
    display_name = os.environ.get("SMTP_FROM_NAME", "").strip()
    msg["From"] = (
        formataddr((display_name, smtp_user)) if display_name else smtp_user
    )
    msg["To"] = to_email

    if dry_run:
        pj = ""
        if attachments:
            pj = " | pj: " + ", ".join(name for name, _ in attachments)
        print(f"[dry-run] À: {to_email!r} | Sujet: {subject!r}{pj}")
        return

    # Octets pour le transport. Ne pas passer une str à sendmail (ré-encodage ascii).
    wire = msg.as_bytes(policy=_SMTP_WIRE_POLICY)

    if smtp_use_ssl:
        with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, [to_email], wire)
    else:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, [to_email], wire)


def main() -> None:
    load_env_file()
    parser = argparse.ArgumentParser(
        description="Lit un CSV et envoie un e-mail SMTP par ligne."
    )
    parser.add_argument(
        "csv_path",
        help="Fichier CSV (UTF-8), avec en-tête ; au moins une colonne email.",
    )
    parser.add_argument(
        "--preset",
        choices=("bracelet", "guide", "relance", "memento"),
        help="bracelet : code d’accès app (Prénom, Nom, Email ; Code barres). "
        "guide : info guide + Atoutbio / dépistage (toutes les lignes avec email). "
        "relance : comme bracelet + rappel de mise à jour de l’application. "
        "memento : informe les participants que le mémento est disponible. "
        "Ignore --subject et --body-template.",
    )
    parser.add_argument(
        "--bracelet-sans-code",
        action="store_true",
        help="Avec --preset bracelet : mail d’information quand la colonne Code barres "
        "est encore vide (pas d’envoi du code individuel).",
    )
    parser.add_argument(
        "--subject",
        default=os.environ.get("MAIL_SUBJECT", "Message"),
        help="Sujet (défaut: variable MAIL_SUBJECT ou « Message »).",
    )
    parser.add_argument(
        "--body-template",
        default=os.environ.get(
            "MAIL_BODY_TEMPLATE",
            "Bonjour {prenom} {nom},\n\nVotre message ici.\n",
        ),
        help="Corps du mail avec placeholders {nom_colonne} (défaut: variable MAIL_BODY_TEMPLATE ou texte générique).",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Pause en secondes entre chaque envoi (évite le rate limiting).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="N'envoie pas : affiche destinataires et sujets.",
    )
    parser.add_argument(
        "--only-email",
        action="append",
        metavar="ADRESSE",
        help="N'envoyer qu'à cette adresse (répétable). Comparaison sans tenir compte de la casse.",
    )
    parser.add_argument(
        "--start-line",
        type=int,
        metavar="N",
        help="Ne traiter qu'à partir de la ligne N du fichier (même numéro que « Ligne N » dans les messages du script ; la ligne 1 est l'en-tête CSV).",
    )
    parser.add_argument(
        "--uid-xlsx",
        type=Path,
        default=Path(__file__).resolve().parent
        / "WILOUT X CARTEL NANCY 2515_liste UID.xlsx",
        help="Excel Wilout (UID → Short tag), utilisé pour --preset bracelet et --preset relance. "
        "Ignoré avec --preset guide et --preset memento.",
    )
    parser.add_argument(
        "--no-uid-short-map",
        action="store_true",
        help="Ne pas charger l’Excel : le corps du mail utilise le code tel que dans le CSV. "
        "Inutile avec --preset guide et --preset memento (l’Excel n’est pas chargé).",
    )
    parser.add_argument(
        "--attach",
        type=Path,
        action="append",
        metavar="FICHIER",
        dest="attach",
        help="Pièce jointe (répétable). Les fichiers sont lus une fois au démarrage puis joints à chaque mail.",
    )
    args = parser.parse_args()

    if args.start_line is not None and args.start_line < 1:
        sys.exit("--start-line doit être >= 1 (1 = première ligne du fichier, en-tête CSV).")

    only_email_norm: set[str] | None = None
    if args.only_email:
        only_email_norm = {e.strip().lower() for e in args.only_email if e.strip()}

    if args.preset == "bracelet" and args.bracelet_sans_code:
        subject_src = BRACELET_MAIL_SUBJECT_PENDING
        body_template_src = BRACELET_MAIL_BODY_PENDING
    elif args.preset == "bracelet":
        subject_src = BRACELET_MAIL_SUBJECT
        body_template_src = BRACELET_MAIL_BODY
    elif args.preset == "guide":
        subject_src = GUIDE_MAIL_SUBJECT
        body_template_src = GUIDE_MAIL_BODY
    elif args.preset == "relance":
        subject_src = RELANCE_MAIL_SUBJECT
        body_template_src = RELANCE_MAIL_BODY
    elif args.preset == "memento":
        subject_src = MEMENTO_MAIL_SUBJECT
        body_template_src = MEMENTO_MAIL_BODY
    else:
        subject_src = args.subject
        body_template_src = args.body_template

    smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER", "")
    smtp_password = os.environ.get("SMTP_PASSWORD", "")
    smtp_ssl_flag = os.environ.get("SMTP_SSL", "").strip().lower() in ("1", "true", "yes")
    # Port 465 = TLS direct (SMTPS). Port 587 = STARTTLS. Ne pas utiliser un hôte *MX* pour l’envoi (ex. mx1.*).
    smtp_use_ssl = smtp_port == 465 or smtp_ssl_flag

    if not args.dry_run:
        if not smtp_user or not smtp_password:
            sys.exit(
                "Définissez SMTP_USER et SMTP_PASSWORD dans un fichier .env ou l’environnement.\n"
                "Ou lancez avec --dry-run pour tester sans envoi."
            )

    attach_payloads: list[tuple[str, bytes]] = []
    if args.attach:
        for p in args.attach:
            if not p.is_file():
                sys.exit(f"Pièce jointe introuvable ou non fichier : {p}")
            attach_payloads.append((p.name, p.read_bytes()))

    uid_to_short: dict[str, str] | None = None
    # Seuls les presets qui utilisent un code bracelet ont besoin de la map UID -> short tag.
    if not args.no_uid_short_map and args.preset in {"bracelet", "relance"}:
        if openpyxl is None:
            print(
                "openpyxl absent : impossible de charger l’Excel UID → Short tag. "
                "Installez openpyxl ou utilisez --no-uid-short-map.",
                file=sys.stderr,
            )
        elif not args.uid_xlsx.is_file():
            print(
                f"Excel Wilout introuvable ({args.uid_xlsx}) : les mails utiliseront le code CSV tel quel.",
                file=sys.stderr,
            )
        else:
            try:
                uid_to_short = load_uid_to_short(args.uid_xlsx)
            except Exception as e:
                print(
                    f"Erreur lecture {args.uid_xlsx} : {e}. Codes CSV non traduits.",
                    file=sys.stderr,
                )

    with open(args.csv_path, newline="", encoding="utf-8-sig") as f:
        first_line = f.readline()
        if not first_line:
            sys.exit("Fichier CSV vide.")
        delimiter = detect_delimiter(first_line)
        f.seek(0)
        reader = csv.DictReader(f, delimiter=delimiter)
        email_key = find_email_field(reader.fieldnames)

        sent = 0
        failed = 0

        if args.start_line is not None:
            print(
                f"Reprise : lignes < {args.start_line} ignorées (même numérotation que « Ligne N »).",
                file=sys.stderr,
            )

        for lineno, row in enumerate(reader, start=2):
            if args.start_line is not None and lineno < args.start_line:
                continue

            raw = (row.get(email_key) or "").strip()
            if not raw:
                print(f"Ligne {lineno}: email vide, ignorée.", file=sys.stderr)
                failed += 1
                continue

            _, addr = parseaddr(raw)
            to_email = addr or raw
            if "@" not in to_email:
                print(f"Ligne {lineno}: adresse invalide {raw!r}, ignorée.", file=sys.stderr)
                failed += 1
                continue

            if only_email_norm is not None and to_email.strip().lower() not in only_email_norm:
                continue

            row_ctx = row_for_templates(row)
            if uid_to_short:
                resolve_bracelet_to_short_tag(
                    row_ctx, uid_to_short, lineno, to_email
                )
            if (
                args.preset in ("bracelet", "relance")
                and not args.bracelet_sans_code
                and not row_ctx.get("braceletNumber", "").strip()
            ):
                print(
                    f"Ligne {lineno}: code bracelet / Code barres vide pour {to_email}, ignorée.",
                    file=sys.stderr,
                )
                failed += 1
                continue

            subject = (
                build_body(subject_src, row_ctx) if "{" in subject_src else subject_src
            )
            body = build_body(body_template_src, row_ctx)

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
                    attach_payloads or None,
                )
                sent += 1
                bn = (row_ctx.get("braceletNumber") or "").strip()
                if args.preset in ("bracelet", "relance") and bn:
                    print(f"OK {to_email} | bracelet : {bn}")
                else:
                    print(f"OK {to_email}")
            except Exception as e:
                failed += 1
                print(f"Ligne {lineno} {to_email}: erreur {e}", file=sys.stderr)

            if args.delay > 0:
                time.sleep(args.delay)

    print(f"Terminé : {sent} envoyé(s), {failed} ignoré(s) ou en échec.")


if __name__ == "__main__":
    main()
