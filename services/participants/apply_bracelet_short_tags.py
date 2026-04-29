#!/usr/bin/env python3
"""
Met à jour `braceletNumber` dans un JSON export participants à partir d'un Excel
Wilout (colonnes UID + Short tag). Les clés JSON sont les UID NFC (avec zéros de tête) ;
les UID Excel sont alignés en normalisant (majuscules, zéros initiaux retirés).

Les entrées dont la clé commence par OR- (badges orga) ne sont pas modifiées.
Les lignes Excel sans UID ou sans Short tag sont ignorées.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError as e:
    print("Installez openpyxl : pip install openpyxl", file=sys.stderr)
    raise SystemExit(1) from e


def norm_uid(s: str) -> str:
    return s.upper().strip().lstrip("0") or "0"


def is_org_key(key: str) -> bool:
    return key.startswith("OR-")


def load_uid_to_short(xlsx_path: Path) -> dict[str, str]:
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


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--xlsx",
        type=Path,
        default=Path("WILOUT X CARTEL NANCY 2515_liste UID.xlsx"),
        help="Fichier Excel (UID, Short tag)",
    )
    p.add_argument(
        "--json",
        type=Path,
        default=Path("participants_export_merged.json"),
        help="JSON participants à mettre à jour",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Affiche les compteurs sans écrire le fichier",
    )
    args = p.parse_args()

    if not args.xlsx.is_file():
        print(f"Fichier Excel introuvable : {args.xlsx}", file=sys.stderr)
        raise SystemExit(1)
    if not args.json.is_file():
        print(f"Fichier JSON introuvable : {args.json}", file=sys.stderr)
        raise SystemExit(1)

    uid_to_short = load_uid_to_short(args.xlsx)

    with args.json.open(encoding="utf-8") as f:
        data = json.load(f)

    updated = 0
    skipped_org = 0
    no_mapping = 0

    for key, rec in data.items():
        if is_org_key(key):
            skipped_org += 1
            continue
        n = norm_uid(key)
        if n not in uid_to_short:
            no_mapping += 1
            continue
        new_b = uid_to_short[n]
        if rec.get("braceletNumber") != new_b:
            updated += 1
        rec["braceletNumber"] = new_b

    print(f"Org (OR-*) non modifiés : {skipped_org}")
    print(f"braceletNumber mis à jour : {updated}")
    if no_mapping:
        print(f"Attention : entrées sans correspondance Excel : {no_mapping}", file=sys.stderr)

    if args.dry_run:
        print("Dry-run : aucune écriture.")
        return

    with args.json.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"Écrit : {args.json}")


if __name__ == "__main__":
    main()
