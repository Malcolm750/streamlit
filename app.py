import streamlit as st
import pandas as pd
from io import BytesIO # Ajout très important pour générer le fichier Excel en mémoire

# Configuration de la page
st.set_page_config(page_title="Analyse des Équipements", layout="wide")

st.title("Tableau de bord de gestion des Équipements")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS ---
st.sidebar.header("1. Charger les fichiers Excel")

fichier_equipements = st.sidebar.file_uploader("Fichier des Équipements", type=["xlsx", "xls"])
fichier_models = st.sidebar.file_uploader("Fichier des Modèles", type=["xlsx", "xls"])
fichier_fournisseurs = st.sidebar.file_uploader("Fichier des Fournisseurs", type=["xlsx", "xls"])

# --- FONCTION POUR GÉNÉRER LE FICHIER EXCEL ---
@st.cache_data # Permet de ne pas recalculer si les données n'ont pas changé
def convertir_df_en_excel(df):
    output = BytesIO()
    # Utilisation de openpyxl pour écrire le fichier Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Equipements')
    processed_data = output.getvalue()
    return processed_data

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None:
    try:
        # Lecture du fichier Excel
        # On garde une copie de l'original au cas où on voudrait les vrais noms de colonnes à la fin
        df_original = pd.read_excel(fichier_equipements)
        df_equipements = df_original.copy()
        
        st.header("Analyse des Sites")
        
        # Nettoyage des noms de colonnes pour l'analyse
        colonnes_nettoyees = {col: str(col).lower().strip() for col in df_equipements.columns}
        df_equipements = df_equipements.rename(columns=colonnes_nettoyees)

        if 'site principal' in df_equipements.columns:
            
            # --- NETTOYAGE DES DONNÉES ---
            # On supprime les lignes où le site n'est pas renseigné
            df_equipements = df_equipements.dropna(subset=['site principal'])
            
            # On harmonise les noms de sites (Majuscule au début, suppression des espaces)
            df_equipements['site principal'] = df_equipements['site principal'].astype(str).str.strip().str.title()
            
            # Pour l'export, on met aussi à jour la colonne dans le dataframe original pour que l'export soit propre
            # (On cherche la colonne d'origine qui correspond à 'site principal' après nettoyage)
            col_originale_site = [col for col in df_original.columns if str(col).lower().strip() == 'site principal'][0]
            df_original[col_originale_site] = df_equipements['site principal']

            # 1. Afficher les statistiques globales
            sites_uniques = sorted(df_equipements['site principal'].unique()) # Ajout d'un tri alphabétique
            
            st.subheader("Nombre d'équipements par site")
            comptage_sites = df_equipements['site principal'].value_counts().reset_index()
            comptage_sites.columns = ['Site principal', "Nombre d'équipements"]
            
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(comptage_sites, use_container_width=True)
            with col2:
                st.bar_chart(comptage_sites.set_index('Site principal'))
                
            # --- NOUVEAU : SECTION DE TÉLÉCHARGEMENT ---
            st.divider() # Ligne de séparation visuelle
            st.header("📥 Extraire les données par site")
            
            # Menu déroulant pour choisir le site
            site_selectionne = st.selectbox("Sélectionnez le site à exporter :", sites_uniques)
            
            # Filtrer le dataframe original en fonction du site sélectionné
            df_filtre = df_original[df_original[col_originale_site] == site_selectionne]
            
            # Afficher un petit aperçu des données
            st.write(f"Aperçu des données pour **{site_selectionne}** ({len(df_filtre)} lignes) :")
            st.dataframe(df_filtre.head()) # .head() n'affiche que les 5 premières lignes pour ne pas surcharger l'écran
            
            # Générer le fichier Excel en mémoire
            excel_data = convertir_df_en_excel(df_filtre)
            
            # Bouton de téléchargement
            st.download_button(
                label=f"Télécharger le fichier Excel pour {site_selectionne}",
                data=excel_data,
                file_name=f"equipements_{site_selectionne.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            # -------------------------------------------

        else:
            st.error("⚠️ La colonne 'site principal' est introuvable. "
                     f"Voici les colonnes trouvées : {', '.join(df_equipements.columns)}")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
        
else:
    st.info("👈 Veuillez charger le fichier des équipements dans le menu de gauche.")
