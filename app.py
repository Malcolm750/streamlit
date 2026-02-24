import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(page_title="Tableau de bord - Équipements & Modèles", layout="wide")

st.title("📊 Tableau de Bord : Rapprochement Équipements / Modèles")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS ---
st.sidebar.header("📁 Chargement des fichiers")

fichier_equipements = st.sidebar.file_uploader("Fichier des Équipements", type=["xlsx", "xls"])
fichier_models = st.sidebar.file_uploader("Fichier des Modèles", type=["xlsx", "xls"])

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None and fichier_models is not None:
    try:
        # 1. Lecture des fichiers
        df_equipements = pd.read_excel(fichier_equipements)
        df_models = pd.read_excel(fichier_models)
        
        # Nettoyage des noms de colonnes (espaces invisibles)
        df_equipements.columns = df_equipements.columns.str.strip()
        df_models.columns = df_models.columns.str.strip()

        st.header("Analyse de la liaison des données")

        # Noms des colonnes cibles
        col_eq = 'Modèle (Référence)'
        col_mod = 'Code référence'
        col_nature = "Nature de l'objet"

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            # --- PRÉPARATION DES CLÉS DE LIAISON ---
            df_equipements['Code de liaison'] = df_equipements[col_eq].astype(str).str.split(',').str[0].str.strip()
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']

            # --- CALCUL DES ORPHELINS ---
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            modeles_sans_equipement = df_models[~df_models['Code de liaison'].isin(df_equipements['Code de liaison'])]
            
            nb_eq_sans_mod = len(equipements_sans_modele)
            nb_mod_sans_eq = len(modeles_sans_equipement)

            # --- AFFICHAGE DU RÉSUMÉ ---
            st.subheader("Résumé du rapprochement")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Équipements", len(df_equipements))
            col2.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)
            col3.metric("🟡 Modèles SANS équipement", nb_mod_sans_eq)
            
            st.divider()

            # --- ANALYSE PAR NATURE DE L'OBJET ---
            st.header("🔍 Analyse par Nature de l'objet")
            
            if col_nature in df_models.columns:
                
                # 1. JOINTURE
                df_fusion = pd.merge(
                    df_equipements, 
                    df_models[['Code de liaison', col_nature]], 
                    on='Code de liaison', 
                    how='left'
                )
                
                # 2. GESTION DES VIDES
                df_fusion[col_nature] = df_fusion[col_nature].fillna("Non défini / Sans modèle")
                
                # 3. NOUVEAU COMPTAGE MULTIPLE : Équipements ET Modèles distincts
                st.subheader("Répartition globale : Équipements vs Modèles distincts")
                
                # On groupe par la Nature de l'objet et on fait deux calculs :
                # - 'size' compte le nombre total de lignes (donc le nombre d'équipements)
                # - 'nunique' compte le nombre de valeurs uniques dans 'Code de liaison' (donc le nombre de modèles différents)
                analyse_nature = df_fusion.groupby(col_nature).agg(
                    nb_equipements=('Code de liaison', 'size'),
                    nb_modeles_uniques=('Code de liaison', 'nunique')
                ).reset_index()
                
                # On renomme les colonnes pour que ce soit propre à l'affichage
                analyse_nature.columns = ["Nature de l'objet", "Nombre d'équipements", "Nombre de modèles utilisés"]
                
                # On trie pour avoir les plus gros volumes d'équipements en haut du tableau
                analyse_nature = analyse_nature.sort_values(by="Nombre d'équipements", ascending=False)
                
                # 4. AFFICHAGE
                col_chart1, col_chart2 = st.columns([1, 2])
                
                with col_chart1:
                    st.dataframe(analyse_nature, use_container_width=True, hide_index=True)
                    
                with col_chart2:
                    # Streamlit va automatiquement créer un graphique avec 2 barres par Nature de l'objet !
                    st.bar_chart(analyse_nature.set_index("Nature de l'objet"))
                    
            else:
                st.error(f"⚠️ La colonne '{col_nature}' est introuvable dans le fichier des modèles.")

            st.divider()
            
            # --- AFFICHAGE DES DONNÉES ORPHELINES ---
            st.header("⚠️ Détail des anomalies (Orphelins)")
            col_gauche, col_droite = st.columns(2)
            
            with col_gauche:
                st.subheader(f"{nb_eq_sans_mod} Équipements sans modèle")
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                else:
                    st.success("Tous les équipements ont un modèle associé.")
                    
            with col_droite:
                st.subheader(f"{nb_mod_sans_eq} Modèles sans équipement")
                if nb_mod_sans_eq > 0:
                    st.dataframe(modeles_sans_equipement[[col_mod, 'Code de liaison']], use_container_width=True)
                else:
                    st.success("Tous les modèles sont utilisés.")

        else:
            st.error(f"⚠️ Les colonnes de liaison sont introuvables.\n"
                     f"- Attendue dans les Équipements : '{col_eq}'\n"
                     f"- Attendue dans les Modèles : '{col_mod}'")

    except Exception as e:
        st.error(f"Une erreur s'est produite lors de l'analyse : {e}")

else:
    st.info("👈 Veuillez charger le fichier des Équipements ET le fichier des Modèles dans le menu de gauche pour lancer l'analyse croisée.")
