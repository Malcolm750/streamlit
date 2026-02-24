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
        
        # Nettoyage des noms de colonnes (minuscules + suppression des espaces)
        colonnes_nettoyees = {col: str(col).lower().strip() for col in df_equipements.columns}
        df_equipements = df_equipements.rename(columns=colonnes_nettoyees)

        if 'site principal' in df_equipements.columns:
            
            # --- NOUVEAU : NETTOYAGE DES DONNÉES DE LA COLONNE ---
            # 1. On supprime les lignes où le site n'est pas renseigné
            df_equipements = df_equipements.dropna(subset=['site principal'])
            
            # 2. On convertit en texte, on enlève les espaces parasites, et on met au format "Titre" (ex: "PARIS" -> "Paris")
            df_equipements['site principal'] = df_equipements['site principal'].astype(str).str.strip().str.title()
            # -----------------------------------------------------

            # 1. Afficher le nom des différents sites
            sites_uniques = df_equipements['site principal'].unique()
            st.subheader("Liste des différents sites")
            st.write(", ".join(sites_uniques))
            
            # 2. Compter le nombre d'équipements par site
            st.subheader("Nombre d'équipements par site")
            
            # On compte les occurrences de chaque site principal
            comptage_sites = df_equipements['site principal'].value_counts().reset_index()
            comptage_sites.columns = ['Site principal', "Nombre d'équipements"]
            
            # Affichage sous forme de tableau
            col1, col2 = st.columns(2)
            
            with col1:
                st.dataframe(comptage_sites, use_container_width=True)
                
            with col2:
                # Affichage sous forme de graphique
                st.bar_chart(comptage_sites.set_index('Site principal'))
                
        else:
            st.error("⚠️ La colonne 'site principal' est introuvable dans le fichier des équipements. "
                     f"Voici les colonnes trouvées : {', '.join(df_equipements.columns)}")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors de la lecture du fichier : {e}")
        
else:
    st.info("👈 Veuillez charger le fichier des équipements dans le menu de gauche pour commencer l'analyse.")
    
# (Optionnel) Vérification pour les autres fichiers
if fichier_models is not None and fichier_fournisseurs is not None:
    st.success("Les fichiers Modèles et Fournisseurs sont chargés et prêts pour les prochaines analyses !")
