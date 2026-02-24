import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(page_title="Analyse des Équipements", layout="wide")

st.title("Tableau de bord de gestion des Équipements")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS ---
st.sidebar.header("1. Charger les fichiers Excel")

fichier_equipements = st.sidebar.file_uploader("Fichier des Équipements", type=["xlsx", "xls"])
fichier_models = st.sidebar.file_uploader("Fichier des Modèles", type=["xlsx", "xls"])
fichier_fournisseurs = st.sidebar.file_uploader("Fichier des Fournisseurs", type=["xlsx", "xls"])

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None:
    try:
        # Lecture du fichier Excel
        df_equipements = pd.read_excel(fichier_equipements)
        
        st.header("Analyse des Sites")
        
        # Vérification que la colonne 'site' (insensible à la casse) existe bien
        # On convertit les noms de colonnes en minuscules pour éviter les erreurs de frappe (Site vs site)
        colonnes_minuscules = {col: str(col).lower() for col in df_equipements.columns}
        df_equipements = df_equipements.rename(columns=colonnes_minuscules)

        if 'site' in df_equipements.columns:
            # 1. Afficher le nom des différents sites
            sites_uniques = df_equipements['site'].dropna().unique()
            st.subheader("Liste des différents sites")
            st.write(", ".join([str(site) for site in sites_uniques]))
            
            # 2. Compter le nombre d'équipements par site
            st.subheader("Nombre d'équipements par site")
            
            # On compte les occurrences de chaque site
            comptage_sites = df_equipements['site'].value_counts().reset_index()
            comptage_sites.columns = ['Site', "Nombre d'équipements"]
            
            # Affichage sous forme de tableau
            col1, col2 = st.columns(2)
            
            with col1:
                st.dataframe(comptage_sites, use_container_width=True)
                
            with col2:
                # Affichage sous forme de graphique pour rendre ça plus visuel (Bonus)
                st.bar_chart(comptage_sites.set_index('Site'))
                
        else:
            st.error("⚠️ La colonne 'site' est introuvable dans le fichier des équipements. "
                     "Veuillez vérifier le nom de vos colonnes.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors de la lecture du fichier : {e}")
        
else:
    st.info("👈 Veuillez charger le fichier des équipements dans le menu de gauche pour commencer l'analyse.")
    
# (Optionnel) Vous pouvez aussi vérifier si les autres fichiers sont chargés pour la suite du projet
if fichier_models is not None and fichier_fournisseurs is not None:
    st.success("Les fichiers Modèles et Fournisseurs sont chargés et prêts pour les prochaines analyses !")
