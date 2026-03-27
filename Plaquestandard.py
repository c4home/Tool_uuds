import tkinter as tk
from tkinter import messagebox

def calculate_layout():
    try:
        target = int(entry_target.get())
    except ValueError:
        messagebox.showerror("Erreur", "Veuillez entrer un nombre entier valide.")
        return

    # Dictionnaire pour garder la meilleure combinaison (le moins de pièces) pour chaque épaisseur
    best_combinations = {}
    
    # Limite pour éviter de calculer à l'infini (cible / 10 + marge de sécurité)
    max_pieces = (target // 10) + 3
    
    # Tester les combinaisons
    for i in range(max_pieces):  # 22mm
        if i * 22 > target + 15: break
        for j in range(max_pieces):  # 20mm
            if i * 22 + j * 20 > target + 15: break
            for k in range(max_pieces):  # 10mm
                total_thick = i * 22 + j * 20 + k * 10
                
                # On regarde tout ce qui est dans une tolérance de +/- 10 mm autour de la cible
                if abs(total_thick - target) <= 10:
                    pieces = i + j + k
                    if pieces == 0: continue
                    
                    # Si cette épaisseur n'est pas encore enregistrée, OU si on a trouvé un moyen
                    # de la faire avec MOINS de pièces, on la sauvegarde.
                    if total_thick not in best_combinations or pieces < best_combinations[total_thick]['total_pieces']:
                        best_combinations[total_thick] = {
                            '22mm': i, '20mm': j, '10mm': k, 
                            'total_pieces': pieces,
                            'diff': total_thick - target
                        }

    text_result.delete(1.0, tk.END)
    
    if not best_combinations:
        text_result.insert(tk.END, f"Aucune combinaison trouvée autour de {target} mm.\n")
        return

    # Trier les résultats : en priorité par la différence d'épaisseur (la plus proche en premier)
    # puis par le nombre de pièces
    sorted_results = sorted(best_combinations.items(), key=lambda x: (abs(x[1]['diff']), x[1]['total_pieces']))
    
    text_result.insert(tk.END, f"Meilleures options pour ~{target} mm :\n\n")
    
    # Afficher les 7 meilleures options
    for thick, res in sorted_results[:7]:
        parts = []
        if res['22mm'] > 0: parts.append(f"{res['22mm']} x 22")
        if res['20mm'] > 0: parts.append(f"{res['20mm']} x 20")
        if res['10mm'] > 0: parts.append(f"{res['10mm']} x 10")
        
        # Formatage de l'écart
        sign = "+" if res['diff'] > 0 else ""
        diff_str = f"(Écart: {sign}{res['diff']}mm)" if res['diff'] != 0 else "(Exact)"
        
        line = f"{thick}mm {diff_str:<15} -> " + " + ".join(parts) + f"  [{res['total_pieces']} pièces]"
        
        text_result.insert(tk.END, line + "\n")

# Configuration de l'interface graphique (UI)
root = tk.Tk()
root.title("Optimiseur d'Épaisseurs")
root.geometry("550x380")
root.config(padx=20, pady=20)

label_inst = tk.Label(root, text="Entrez l'épaisseur cible (mm) :", font=("Arial", 12))
label_inst.pack(pady=5)

entry_target = tk.Entry(root, font=("Arial", 14), justify="center", width=15)
entry_target.pack(pady=5)
entry_target.bind('<Return>', lambda event: calculate_layout())

btn_calc = tk.Button(root, text="Trouver le minimum de pièces", font=("Arial", 11, "bold"), 
                     command=calculate_layout, bg="#28a745", fg="white", cursor="hand2")
btn_calc.pack(pady=15)

text_result = tk.Text(root, font=("Consolas", 11), height=12, width=65, bg="#f4f4f4")
text_result.pack()

root.mainloop()
