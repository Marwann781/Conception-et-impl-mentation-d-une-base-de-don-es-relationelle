import pandas as pd

# Charger le fichier Excel contenant les données nettoyées
df = pd.read_excel("donnees_transformees.xlsx")

# Créer un identifiant unique pour chaque combinaison (Nom, Prénom, Date de naissance)
# Le même membre peut apparaître plusieurs fois (s'il est adhérent plusieurs années)
df['id_membre'] = df.groupby(['Nom', 'Prenom', 'DATE DE NAISSANCE'], sort=False).ngroup() + 1

# Réorganiser les colonnes pour que 'id_membre' soit au début
colonnes = ['id_membre'] + [col for col in df.columns if col != 'id_membre']
df = df[colonnes]

# Sauvegarder le nouveau fichier avec l'ID membre
df.to_excel("donnees_avec_id_membre.xlsx", index=False)

# Afficher les 5 premières lignes pour vérifier
print(df.head())
