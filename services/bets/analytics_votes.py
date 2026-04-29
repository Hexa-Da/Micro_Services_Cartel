import json
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pathlib import Path

# --- CONFIGURATION ---
PROJECT_ROOT = Path(__file__).resolve().parents[2]
filename = PROJECT_ROOT / "data" / "resources" / "cummap-7afee-default-rtdb-delegationBets-export.json"
output_folder = PROJECT_ROOT / "data" / "output" / "graphs_resultats_final"

# --- 1. CHARGEMENT ---
if not os.path.exists(filename):
    print(f"ERREUR : '{filename}' introuvable.")
    exit()

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

with open(filename, 'r', encoding='utf-8') as f:
    data = json.load(f)

delegations = sorted(data.keys())

all_sports = set()
for t in data.values():
    all_sports.update(t.keys())
all_sports = sorted(list(all_sports))

# Candidats possibles par sport = union des clés dans `votes` (noms réels du JSON)
candidates_by_sport = {sport: set() for sport in all_sports}
for sports_data in data.values():
    for sport in all_sports:
        votes_dict = sports_data.get(sport, {}).get('votes') or {}
        candidates_by_sport[sport].update(votes_dict.keys())
for sport in all_sports:
    candidates_by_sport[sport] = sorted(candidates_by_sport[sport])

# Premier vainqueur rencontré par sport (cohérent avec leadboard / RTDB)
winner_by_sport = {}
for team in delegations:
    for sport in all_sports:
        if sport in winner_by_sport:
            continue
        w = data[team].get(sport, {}).get("winner")
        if w:
            winner_by_sport[sport] = w


def title_with_winner(base: str, sport: str) -> str:
    w = winner_by_sport.get(sport)
    sub = f"Vainqueur : {w}" if w else "Vainqueur : non renseigné"
    return f"{base}\n{sub}"


precision_rows = []
vote_rows = []
grand_total_votes = 0

print(f"{'='*40}")
print(f"   STATISTIQUES DES VOTES PAR DÉLÉGATION")
print(f"{'='*40}\n")

# --- 2. TRAITEMENT ---
for team in delegations:
    sports_data = data[team]
    team_total_votes = 0
    print(f"🔵 DÉLÉGATION : {team.upper()}")

    for sport in all_sports:
        results = sports_data.get(sport, {})
        winner = results.get('winner')
        total_votes = results.get('totalVotes', 0)
        votes_dict = results.get('votes', {})
        
        if total_votes > 0:
            print(f"   - {sport:<25} : {total_votes} votes")
            team_total_votes += total_votes
            grand_total_votes += total_votes
        
        accuracy = 0
        if winner and total_votes > 0:
            votes_for_winner = votes_dict.get(winner, 0)
            accuracy = (votes_for_winner / total_votes) * 100
        
        if total_votes > 0 or winner:
             precision_rows.append({"Equipe": team, "Sport": sport, "Precision": accuracy})

        if total_votes > 0:
            for candidat in candidates_by_sport[sport]:
                vote_count = votes_dict.get(candidat, 0)
                percent = (vote_count / total_votes * 100)
                vote_rows.append({
                    "Sport": sport, "Votant": team, 
                    "Candidat": candidat, "Pourcentage": percent
                })

    if team_total_votes == 0: print("   (Aucun vote)")
    else: print(f"   => Total {team}: {team_total_votes} votes")
    print("-" * 20)

print(f"\n📢 TOTAL VOTES : {grand_total_votes}")
print("Génération des graphiques par sport...\n")

df_precision = pd.DataFrame(precision_rows)
df_votes = pd.DataFrame(vote_rows)

# --- 3. GÉNÉRATION DES GRAPHIQUES ---

for sport in all_sports:
    safe_sport_name = sport.replace(" ", "_").replace("é", "e").replace("è", "e")
    
    # === A. GRAPHIQUE DE PRÉCISION (Barplot) ===
    if not df_precision.empty:
        data_prec_sport = df_precision[df_precision["Sport"] == sport]
        
        if not data_prec_sport.empty:
            plt.figure(figsize=(10, 6))
            sns.set_theme(style="whitegrid")
            
            # CORRECTION ICI : Ajout de hue="Equipe" et legend=False
            ax = sns.barplot(
                data=data_prec_sport,
                x="Equipe", 
                y="Precision", 
                hue="Equipe",     # <--- Correction Warning
                legend=False,     # <--- Correction Warning
                order=[t for t in delegations if t in set(data_prec_sport["Equipe"].unique())],
                palette="viridis", 
                edgecolor="black"
            )
            
            # AJOUT ICI : Affichage des pourcentages sur les barres
            for container in ax.containers:
                ax.bar_label(container, fmt='%.0f%%', padding=3, fontweight='bold')
            
            plt.title(
                title_with_winner(f"Précision des pronostics : {sport}", sport),
                fontsize=13,
                fontweight="bold",
            )
            plt.ylabel("% de réussite")
            plt.xlabel("")
            plt.ylim(0, 115) # Un peu plus de marge en haut pour le texte
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            save_path = os.path.join(output_folder, f"precision_{safe_sport_name}.png")
            plt.savefig(save_path)
            plt.close()
            print(f"-> Créé : precision_{safe_sport_name}.png")

    # === B. GRAPHIQUE DES VOTES (Heatmap) ===
    if not df_votes.empty:
        data_vote_sport = df_votes[df_votes["Sport"] == sport]
        
        if not data_vote_sport.empty:
            matrix = data_vote_sport.pivot(index="Votant", columns="Candidat", values="Pourcentage")
            
            active_voters = [t for t in delegations if t in matrix.index and matrix.loc[t].sum() > 0]
            active_cols = [c for c in candidates_by_sport[sport] if c in matrix.columns and matrix[c].sum() > 0]
            matrix_filtered = matrix.loc[active_voters, active_cols]

            if not matrix_filtered.empty:
                w = min(22, 8 + len(active_cols) * 0.45)
                h = min(18, 5 + len(active_voters) * 0.38)
                plt.figure(figsize=(w, h))
                sns.heatmap(
                    matrix_filtered, annot=True, fmt=".0f", cmap="Blues", 
                    cbar=False, linewidths=.5, linecolor='lightgray', vmin=0, vmax=100
                )
                
                plt.title(
                    title_with_winner(f"Qui a voté pour qui : {sport}", sport),
                    fontsize=13,
                    fontweight="bold",
                    pad=18,
                )
                plt.xlabel("Candidat")
                plt.ylabel("Votant")
                plt.tight_layout()
                
                save_path = os.path.join(output_folder, f"vote_{safe_sport_name}.png")
                plt.savefig(save_path)
                plt.close()
                print(f"-> Créé : vote_{safe_sport_name}.png")

print(f"\nTerminé ! Tous les fichiers sont dans '{output_folder}'.")
