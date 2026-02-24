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
        col_nature = "Nature de l'objet" # Nouvelle colonne ciblée

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

            # --- NOUVEAU : ANALYSE PAR NATURE DE L'OBJET ---
            st.header("🔍 Analyse par Nature de l'objet")
            
            # On vérifie si la colonne existe bien dans le fichier des modèles
            if col_nature in df_models.columns:
                
                # 1. JOINTURE : On ramène la Nature de l'objet dans le tableau des équipements
                # Le 'how="left"' signifie qu'on garde tous les équipements, même ceux qui n'ont pas de modèle
                df_fusion = pd.merge(
                    df_equipements, 
                    df_models[['Code de liaison', col_nature]], # On ne prend que le code et la nature
                    on='Code de liaison', 
                    how='left'
                )
                
                # 2. GESTION DES VIDES : Si un équipement n'a pas de modèle, sa nature sera vide. On remplace par un texte clair.
                df_fusion[col_nature] = df_fusion[col_nature].fillna("Non défini / Sans modèle")
                
                # 3. COMPTAGE : On compte les équipements par nature
                st.subheader("Répartition globale")
                comptage_nature = df_fusion[col_nature].value_counts().reset_index()
                comptage_nature.columns = ["Nature de l'objet", "Nombre d'équipements"]
                
                # 4. AFFICHAGE
                col_chart1, col_chart2 = st.columns([1, 2]) # Le graphique prendra 2 fois plus de place que le tableau
                
                with col_chart1:
                    st.dataframe(comptage_nature, use_container_width=True)
                    
                with col_chart2:
                    st.bar_chart(comptage_nature.set_index("Nature de l'objet"))
                    
            else:
                st.error(f"⚠️ La colonne '{col_nature}' est introuvable dans le fichier des modèles.")

            st.divider()
            
            # --- AFFICHAGE DES DONNÉES ORPHELINES (déplacé en bas) ---
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
