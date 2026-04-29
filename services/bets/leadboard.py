from __future__ import annotations

import json
from pathlib import Path

# Résultats officiels : premier `winner` rencontré par sport (même schéma que delegationBets)
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_RESULTS = PROJECT_ROOT / "data" / "resources" / "cummap-7afee-default-rtdb-delegationBets-export.json"
PARTICIPANTS_FILE = PROJECT_ROOT / "data" / "resources" / "cummap-7afee-default-rtdb-participants-export.json"
OUTPUT_TXT = PROJECT_ROOT / "data" / "output" / "leadboard.txt"
OUTPUT_DELEGATIONS_TXT = PROJECT_ROOT / "data" / "output" / "classement_delegations_precision.txt"

# Paris des fake_data (generate.py) → libellés des vainqueurs dans data.json
BET_TO_OFFICIAL_NAME = {
    "Albi": "Mines Albi",
    "Alès": "Mines Alès",
    "Nancy": "Mines Nancy",
    "Paris": "Mines Paris",
    "Sainté": "Mines Sainté",
    "Douai": "IMT NE",
    "ENSAIS": "ENSAIA",
    "IMT Atlantique": "IMT A",
}


def canonical_delegation(name: str) -> str:
    if not name:
        return name
    return BET_TO_OFFICIAL_NAME.get(name, name)


def load_data(filename: Path):
    with open(filename, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_winners(results_data: dict) -> dict:
    """
    Construit {sport: vainqueur} depuis l’export delegationBets :
    racine = délégation → sports → winner.
    On retient le premier winner trouvé pour chaque sport.
    """
    winners = {}
    for _delegation, sports in results_data.items():
        if not isinstance(sports, dict):
            continue
        for sport_name, info in sports.items():
            if not isinstance(info, dict):
                continue
            winner = info.get("winner")
            if winner and sport_name not in winners:
                winners[sport_name] = winner
    return winners


def compute_delegation_scores(results_data: dict) -> list:
    """
    Précision par délégation : (somme des votes placés sur le vainqueur officiel)
    / (somme des votes) sur tous les sports où un winner est renseigné.
    """
    all_sports = set()
    for sports in results_data.values():
        if isinstance(sports, dict):
            all_sports.update(sports.keys())
    all_sports = sorted(all_sports)

    rows = []
    for delegation, sports_data in results_data.items():
        if not isinstance(sports_data, dict):
            continue
        correct = 0
        total = 0
        n_sports = 0
        for sport in all_sports:
            r = sports_data.get(sport, {})
            if not isinstance(r, dict):
                continue
            w = r.get("winner")
            tv = r.get("totalVotes") or 0
            vd = r.get("votes") or {}
            if w and tv > 0:
                n_sports += 1
                correct += vd.get(w, 0)
                total += tv
        if total <= 0:
            continue
        rows.append(
            {
                "delegation": delegation,
                "correct": correct,
                "total": total,
                "n_sports": n_sports,
                "accuracy": correct / total * 100,
            }
        )

    rows.sort(
        key=lambda x: (x["accuracy"], x["total"]),
        reverse=True,
    )
    return rows


def format_delegation_table(rows: list) -> str:
    lines = [
        f"CLASSEMENT DES DÉLÉGATIONS PAR PRÉCISION DES PARIS ({len(rows)} ligne(s))",
        "Métrique : (votes sur le vainqueur) / (total des votes), tous sports avec résultat.",
        "-" * 88,
        f"{'Rang':<5} {'Délégation':<24} {'%Précision':>11} {'Bons/total votes':>20} {'Sports':>8}",
        "-" * 88,
    ]
    for rank, r in enumerate(rows, start=1):
        ct = f"{r['correct']}/{r['total']}"
        lines.append(
            f"{rank:<5} "
            f"{str(r['delegation']):<24} "
            f"{r['accuracy']:>10.2f}% "
            f"{ct:>20} "
            f"{r['n_sports']:>8}"
        )
    return "\n".join(lines) + "\n"


def compute_bettor_scores(participants: list, winners: dict):
    """
    participants : liste de dicts avec prenom, nom, delegation, braceletNumber, bets.
    Une ligne par participant (y compris sans paris : 0 parié, 0 %).
    """
    scores = []
    for p in participants:
        bets = p.get("bets") or {}

        correct = 0
        attempted = 0

        for sport_name, bet_team in bets.items():
            official_winner = winners.get(sport_name)
            if not official_winner:
                continue
            attempted += 1
            if canonical_delegation(bet_team) == official_winner:
                correct += 1

        if attempted == 0:
            accuracy = 0.0
        else:
            accuracy = correct / attempted * 100

        scores.append(
            {
                "prenom": p.get("prenom"),
                "nom": p.get("nom"),
                "delegation": p.get("delegation"),
                "braceletNumber": p.get("braceletNumber"),
                "correct": correct,
                "attempted": attempted,
                "accuracy": accuracy,
            }
        )

    scores.sort(
        key=lambda s: (s["correct"], s["accuracy"], s["attempted"]), reverse=True
    )
    return scores


def format_ranking_table(scores: list, top_n: int | None = None) -> str:
    """
    Tableau texte colonnes fixes.
    top_n None = aucune limite (tout le classement, pour le fichier txt).
    """
    if top_n is None:
        slice_scores = scores
    else:
        slice_scores = scores[:top_n]
    lines = [
        f"CLASSEMENT DES PARIEURS ({len(slice_scores)} ligne(s))",
        "-" * 90,
        f"{'Rang':<5} {'Bracelet':<14} {'Nom':<12} {'Prénom':<12} {'Délégation':<15} "
        f"{'Bons':<4} {'Pariés':<6} {'%Réussite':>10}",
        "-" * 90,
    ]
    for rank, s in enumerate(slice_scores, start=1):
        lines.append(
            f"{rank:<5} "
            f"{str(s.get('braceletNumber', '')):<14} "
            f"{str(s['nom']):<12} "
            f"{str(s['prenom']):<12} "
            f"{str(s['delegation']):<15} "
            f"{s['correct']:<4} "
            f"{s['attempted']:<6} "
            f"{s['accuracy']:>9.1f}%"
        )
    return "\n".join(lines) + "\n"


def print_ranking(scores, top_n=50):
    print(format_ranking_table(scores, top_n=top_n), end="")


def write_ranking_txt(path: Path, scores: list) -> None:
    """Écrit tout le classement sans troncature (pas de plafond de lignes)."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(format_ranking_table(scores, top_n=None))


def write_delegations_txt(path: Path, rows: list) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(format_delegation_table(rows))


def main():
    # 1) Vainqueur par sport (premier winner rencontré par discipline)
    results = load_data(DATA_RESULTS)
    winners = extract_winners(results)

    # 2) Paris par participant (export RTDB participants)
    raw = load_data(PARTICIPANTS_FILE)
    participants = [p for p in raw.values() if isinstance(p, dict)]

    # 3) Réussite : paris normalisés comparés aux vainqueurs officiels
    scores = compute_bettor_scores(participants, winners)
    print_ranking(scores, top_n=50)
    write_ranking_txt(OUTPUT_TXT, scores)
    print(f"(Export tableau : {OUTPUT_TXT})")

    # 4) Classement des délégations (précision agrégée des votes de groupe)
    deleg_rows = compute_delegation_scores(results)
    write_delegations_txt(OUTPUT_DELEGATIONS_TXT, deleg_rows)
    print(f"(Export délégations : {OUTPUT_DELEGATIONS_TXT})")
    if deleg_rows:
        top = deleg_rows[0]
        print(
            f"    1er : {top['delegation']} — {top['accuracy']:.2f}% "
            f"({top['correct']}/{top['total']} votes)"
        )


if __name__ == "__main__":
    main()
