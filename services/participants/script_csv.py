#!/usr/bin/env python3
"""
Convertit un Excel (.xlsx) ou un CSV de participants vers un JSON compatible avec le format
`participants_export_<délégation>.json` par défaut (structure: { "<braceletNumber>": { ... } }).

Champs générés:
- braceletNumber (string, colonne bracelet / code-barres ou fallback sur l'index)
- delegation
- nom
- prenom

Champs omis:
- activatedAt (géré côté app)
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Tuple


def _strip_accents(s: str) -> str:
    # Unicode NFKD + suppression des marques diacritiques.
    normalized = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def normalize_header(header: Any) -> str:
    if header is None:
        return ""
    s = str(header)
    s = _strip_accents(s)
    s = s.strip().lower()
    # Normalise un maximum de variations (espaces, underscores, tirets, #, etc.).
    s = re.sub(r"[\s_\-/#.]+", "", s)
    return s


def as_non_empty_string(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    return s or None


def _slug_for_output_filename(label: str) -> str:
    """Fragment sûr pour un nom de fichier (pas de séparateurs de chemin)."""
    s = _strip_accents(label.strip())
    s = re.sub(r"[^\w\-.]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("._")
    return s or "export"


def default_output_path(input_path: str, fixed_delegation: Optional[str]) -> str:
    """
    participants_export_<délégation ou stem>.json si --output n'est pas fourni.
    """
    if fixed_delegation:
        slug = _slug_for_output_filename(fixed_delegation)
    else:
        slug = _slug_for_output_filename(Path(input_path).stem)
    return f"participants_export_{slug}.json"


def strip_billetterie_albi_suffix(value: str) -> str:
    """Retire un suffixe type export billetterie « - Albi » en fin de prénom."""
    return re.sub(r"\s*-\s*Albi\s*$", "", value.strip(), flags=re.IGNORECASE).strip()


def parse_nom_et_delegation_billetterie(raw: str) -> Tuple[str, Optional[str]]:
    """
    Export billetterie : « NOM - Ville » (séparateur ` - `) → nom seul et délégation.
    Ex. « BANCEL - Albi » → (« BANCEL », « Albi »), « Abbas - Alès » → (« Abbas », « Alès »).
    """
    s = raw.strip()
    if not s:
        return "", None
    if " - " in s:
        left, right = s.rsplit(" - ", 1)
        left, right = left.strip(), right.strip()
        if left and right:
            return left, right
    return s, None


def build_header_index(header_values: Iterable[Any]) -> Dict[str, int]:
    """
    Construit une table: normalized_header -> column_index (1-based).
    Si plusieurs colonnes ont le même header normalisé, on garde la première.
    Les colonnes « Informations … » (export billetterie) sont ignorées.
    """

    out: Dict[str, int] = {}
    for idx, val in enumerate(header_values, start=1):
        norm = normalize_header(val)
        if not norm:
            continue
        if norm.startswith("information"):
            continue
        out.setdefault(norm, idx)
    return out


def resolve_column(
    *,
    header_index: Dict[str, int],
    desired_field: str,
    override_header_value: Optional[str],
    aliases: Iterable[str],
    strict: bool,
) -> Optional[int]:
    if override_header_value:
        override_norm = normalize_header(override_header_value)
        idx = header_index.get(override_norm)
        if idx is None and strict:
            raise ValueError(
                f"Colonne obligatoire introuvable pour '{desired_field}': "
                f"header='${override_header_value}' (normalisé='{override_norm}')"
            )
        return idx

    for alias in aliases:
        idx = header_index.get(normalize_header(alias))
        if idx is not None:
            return idx
    return None


def is_empty_row(values: Iterable[Any]) -> bool:
    for v in values:
        if v is None:
            continue
        if isinstance(v, str) and not v.strip():
            continue
        return False
    return True


def _load_csv_rows(path: str, delimiter: str) -> Tuple[Tuple[Any, ...], Iterator[Tuple[Any, ...]]]:
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        all_rows: List[List[str]] = list(reader)
    if not all_rows:
        raise ValueError(f"CSV vide: {path}")
    header = tuple(all_rows[0])

    def data_rows() -> Iterator[Tuple[Any, ...]]:
        for r in all_rows[1:]:
            yield tuple(r)

    return header, data_rows()


def _load_xlsx_rows(
    path: str, sheet: Optional[str]
) -> Tuple[Tuple[Any, ...], Iterator[Tuple[Any, ...]]]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet:
        if sheet not in wb.sheetnames:
            raise ValueError(
                f"Sheet introuvable: '{sheet}'. Sheets disponibles: {wb.sheetnames}"
            )
        ws = wb[sheet]
    else:
        ws = wb.active

    header_row = 1
    header_values = next(
        ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)
    )

    def data_rows() -> Iterator[Tuple[Any, ...]]:
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            yield row

    return header_values, data_rows()


def load_header_and_data_rows(
    path: str, sheet: Optional[str], csv_delimiter: str
) -> Tuple[Tuple[Any, ...], List[Tuple[Any, ...]]]:
    """
    Retourne (ligne d'en-tête, lignes de données matérialisées en liste).
    """
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        header, row_it = _load_csv_rows(path, csv_delimiter)
    elif suffix in (".xlsx", ".xlsm"):
        header, row_it = _load_xlsx_rows(path, sheet)
    else:
        raise ValueError(
            f"Format non supporté ({suffix}). Utilisez .csv, .xlsx ou .xlsm."
        )
    return header, list(row_it)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Excel / CSV -> participants_export (JSON)"
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Chemin du fichier (.xlsx, .xlsm ou .csv)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Nom de l'onglet Excel à lire (ignoré pour les CSV)",
    )
    parser.add_argument(
        "--csv-delimiter",
        default=";",
        help="Séparateur pour les fichiers CSV (défaut: ';', export Excel FR)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help=(
            "Chemin du JSON de sortie (défaut: participants_export_<délégation>.json, "
            "ex. Albi pour Input/Albi.csv ; sinon le nom du fichier sans extension)"
        ),
    )
    parser.add_argument(
        "--no-skip-empty-rows",
        dest="skip_empty_rows",
        action="store_false",
        default=True,
        help="Ne pas ignorer les lignes vides",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Echoue si une colonne/valeur requise est manquante",
    )
    parser.add_argument(
        "--on-duplicate",
        choices=["error", "overwrite"],
        default="error",
        help="Que faire en cas de doublon braceletNumber",
    )
    parser.add_argument(
        "--delegation-value",
        default=None,
        help=(
            "Délégation mise pour tous les participants (la colonne délégation du fichier est ignorée). "
            "Si non précisé et que le fichier s'appelle Albi.csv, la valeur par défaut est « Albi »."
        ),
    )

    # Overrides de colonnes par nom d'en-tête.
    parser.add_argument("--col-braceletNumber", default=None)
    parser.add_argument("--col-delegation", default=None)
    parser.add_argument("--col-nom", default=None)
    parser.add_argument("--col-prenom", default=None)

    return parser.parse_args(argv)


def _row_is_informations_section(
    row_values: list[Any], col_nom: Optional[int]
) -> bool:
    """Ignore les lignes « Informations » (titre de bloc, pas un participant)."""
    if not col_nom or col_nom - 1 >= len(row_values):
        return False
    v = as_non_empty_string(row_values[col_nom - 1])
    if not v:
        return False
    norm = normalize_header(v)
    return norm == "informations" or norm.startswith("informations")


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    header_values, data_rows = load_header_and_data_rows(
        args.input,
        sheet=args.sheet,
        csv_delimiter=args.csv_delimiter,
    )
    header_index = build_header_index(header_values)

    fixed_delegation: Optional[str] = args.delegation_value
    if fixed_delegation is None and Path(args.input).stem.lower() == "albi":
        fixed_delegation = "Albi"

    output_path = args.output or default_output_path(args.input, fixed_delegation)

    aliases = {
        "braceletNumber": [
            "braceletNumber",
            "bracelet number",
            "bracelet",
            "braceletNumber#",
            "numero de bracelet",
            "numéro de bracelet",
            "braceletid",
            "code barres",
            "code barre",
            "code-barres",
            "code-barre",
            "barcode",
        ],
        "delegation": [
            "delegation",
            "délégation",
            "delegation ",
            "délégation ",
        ],
        "nom": ["nom", "lastname", "nom de famille"],
        "prenom": ["prenom", "prénom", "firstname", "prénom "],
    }

    col_indices: Dict[str, Optional[int]] = {
        "braceletNumber": resolve_column(
            header_index=header_index,
            desired_field="braceletNumber",
            override_header_value=args.col_braceletNumber,
            aliases=aliases["braceletNumber"],
            strict=args.strict,
        ),
        "delegation": None
        if fixed_delegation is not None
        else resolve_column(
            header_index=header_index,
            desired_field="delegation",
            override_header_value=args.col_delegation,
            aliases=aliases["delegation"],
            strict=args.strict,
        ),
        "nom": resolve_column(
            header_index=header_index,
            desired_field="nom",
            override_header_value=args.col_nom,
            aliases=aliases["nom"],
            strict=args.strict,
        ),
        "prenom": resolve_column(
            header_index=header_index,
            desired_field="prenom",
            override_header_value=args.col_prenom,
            aliases=aliases["prenom"],
            strict=args.strict,
        ),
    }

    # Validation des colonnes requises (hors braceletNumber qui a un fallback).
    # La délégation peut venir du suffixe « - Ville » dans la colonne nom (export billetterie).
    required_fields = ["nom", "prenom"]
    missing_required = [f for f in required_fields if col_indices.get(f) is None]
    if missing_required:
        msg = "Colonnes requises introuvables dans le fichier: " + ", ".join(missing_required)
        if args.strict:
            raise ValueError(msg)
        print("WARNING:", msg, file=sys.stderr)

    participants: Dict[str, Dict[str, Any]] = {}
    used_rows = 0
    skipped_rows = 0

    # Iteration sur les lignes de données (après l'entête).
    # participant_idx = index dense sur les lignes réellement utilisées.
    participant_idx = 0
    for row in data_rows:
        row_values = list(row)

        if _row_is_informations_section(row_values, col_indices.get("nom")):
            skipped_rows += 1
            continue

        # Détection "ligne vide" basée sur les champs requis (s'ils existent).
        fields_for_empty_check = ["nom", "prenom"]
        check_values = []
        for f in fields_for_empty_check:
            idx = col_indices.get(f)
            check_values.append(row_values[idx - 1] if idx and idx - 1 < len(row_values) else None)

        if args.skip_empty_rows and is_empty_row(check_values):
            skipped_rows += 1
            continue

        used_rows += 1

        def get_cell(field: str) -> Any:
            idx = col_indices.get(field)
            if not idx:
                return None
            if idx - 1 >= len(row_values):
                return None
            return row_values[idx - 1]

        raw_bracelet = get_cell("braceletNumber")
        bracelet_number = as_non_empty_string(raw_bracelet)
        if bracelet_number is None:
            bracelet_number = str(participant_idx)

        raw_nom = as_non_empty_string(get_cell("nom")) or ""
        nom, deleg_from_nom = parse_nom_et_delegation_billetterie(raw_nom)

        if fixed_delegation is not None:
            delegation = fixed_delegation
        elif deleg_from_nom is not None:
            delegation = deleg_from_nom
        else:
            delegation = as_non_empty_string(get_cell("delegation")) or ""

        prenom = strip_billetterie_albi_suffix(
            as_non_empty_string(get_cell("prenom")) or ""
        )

        # Validation minimale (sauf si strict = on, on échoue).
        missing_values = []
        value_checks: list[tuple[str, str]] = [("nom", nom), ("prenom", prenom)]
        if fixed_delegation is None:
            value_checks.insert(0, ("delegation", delegation))
        for label, v in value_checks:
            if not v:
                missing_values.append(label)

        if missing_values:
            msg = f"Ligne ignorée (champs manquants: {', '.join(missing_values)})"
            if args.strict:
                raise ValueError(msg)
            print("WARNING:", msg, file=sys.stderr)

        participant = {
            "braceletNumber": str(bracelet_number),
            "delegation": delegation,
            "nom": nom,
            "prenom": prenom,
        }

        key = str(bracelet_number)
        if key in participants:
            if args.on_duplicate == "error":
                raise ValueError(f"Doublon braceletNumber: {key}")
            participants[key] = participant
        else:
            participants[key] = participant

        participant_idx += 1

    out = json.dumps(participants, indent=2, ensure_ascii=False)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(out)

    print(
        f"Export terminé. Lignes utilisées: {used_rows}, lignes ignorées: {skipped_rows}, participants: {len(participants)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

