import pandas as pd

# Charger le fichier Excel
df = pd.read_excel("dataAsso.xlsx")

# Liste des années à traiter
annees = list(range(2015, 2025))

# Colonnes personnelles à garder
colonnes_perso = [
    'Nom', 'Prenom', 'DATE DE NAISSANCE', 'Rue', 'Ville',
    'Latitude Ville', 'Longitude Ville', 'Etat', 'TELEPHONE'
]

# Heures possibles pour l’arrivée et le départ
heures_arrivee = ['10h', '12h', '14h', '16h']
heures_depart = ['12h', '14h', '16h', '18h']

# Colonnes dans le fichier Excel
colonnes_arrivee = [f'Réunion 21 décembre 2024 – arrivée {h}' for h in heures_arrivee]
colonnes_depart = [f'Réunion 21 décembre 2024 – depart {h}' for h in heures_depart]

# Autres infos de la réunion
autres_infos = [
    'Réunion 21 décembre 2024 – Repas proposé',
    'Réunion 21 décembre 2024 – Regime alimentaire',
    'Réunion 21 décembre 2024 – Remarques',
    'Réunion 21 décembre 2024 – Présent'
]

# Liste pour les lignes transformées
resultat = []

# Parcourir les membres
for index, ligne in df.iterrows():
    # Infos personnelles
    perso = {col: ligne.get(col) for col in colonnes_perso}
    
    # Chercher l’heure cochée (valeur = "oui") pour l’arrivée et le départ
    arrivee = next((h for h, col in zip(heures_arrivee, colonnes_arrivee) if ligne.get(col) == 'oui'), None)
    depart = next((h for h, col in zip(heures_depart, colonnes_depart) if ligne.get(col) == 'oui'), None)
    
    # Autres infos réunion
    infos_reunion = {
        'Heure_Arrivee': arrivee,
        'Heure_Depart': depart,
        'Repas': ligne.get(autres_infos[0]),
        'Regime_Alimentaire': ligne.get(autres_infos[1]),
        'Remarques': ligne.get(autres_infos[2]),
        'Presence': ligne.get(autres_infos[3])
    }
    
    # Parcourir les années
    for annee in annees:
        col_date = f'DATE ADHESION {annee}'
        col_montant = f'MONTANT {annee}'
        col_don = f'DON {annee}'
        col_paiement = f'MOYEN DE PAIEMENT {annee}'
        col_bureau = f'BUREAU {annee}'
        
        date_adh = ligne.get(col_date)
        montant = ligne.get(col_montant)
        don = ligne.get(col_don)
        paiement = ligne.get(col_paiement)
        bureau = ligne.get(col_bureau)
        
        if pd.notna(date_adh) or pd.notna(montant) or pd.notna(don):
            nouvelle_ligne = {
                **perso,
                'Annee': annee,
                'Date_Adhesion': date_adh,
                'Montant_Adhesion': montant,
                'Don': don,
                'Moyen_Paiement': paiement,
                'Fonction_Bureau': bureau,
                **infos_reunion
            }
            resultat.append(nouvelle_ligne)

# Créer un tableau final
df_final = pd.DataFrame(resultat)

# Sauvegarder dans un fichier Excel
df_final.to_excel("donnees_transformees.xlsx", index=False)

# Afficher les premières lignes
print(df_final.head())
