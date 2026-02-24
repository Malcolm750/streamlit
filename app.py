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
# On ne lance l'analyse que si les DEUX fichiers sont chargés
if fichier_equipements is not None and fichier_models is not None:
    try:
        # 1. Lecture des fichiers
        df_equipements = pd.read_excel(fichier_equipements)
        df_models = pd.read_excel(fichier_models)
        
        # Nettoyage rapide des noms de colonnes pour éviter les soucis d'espaces invisibles
        df_equipements.columns = df_equipements.columns.str.strip()
        df_models.columns = df_models.columns.str.strip()

        st.header("Analyse de la liaison des données")

        # Vérification de l'existence des colonnes nécessaires
        col_eq = 'Modèle (Référence)'
        col_mod = 'Code référence'

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            # --- PRÉPARATION DES CLÉS DE LIAISON ---
            
            # Pour les équipements : On extrait tout ce qui est avant la première virgule
            # .astype(str) : convertit tout en texte
            # .str.split(',').str[0] : coupe au niveau de la virgule et garde la 1ère partie
            # .str.strip() : enlève les espaces au début et à la fin
            df_equipements['Code de liaison'] = (
                df_equipements[col_eq]
                .astype(str)
                .str.split(',')
                .str[0]
                .str.strip()
            )
            
            # Pour les modèles : On s'assure juste que c'est du texte propre sans espaces
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            # On retire les valeurs vides (nan) qui pourraient fausser les calculs
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']


            # --- CALCUL DES ORPHELINS (Ceux qui n'ont pas de correspondance) ---
            
            # Équipements dont le 'Code de liaison' n'est PAS dans les 'Code de liaison' des modèles
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            
            # Modèles dont le 'Code de liaison' n'est PAS dans les 'Code de liaison' des équipements
            modeles_sans_equipement = df_models[~df_models['Code de liaison'].isin(df_equipements['Code de liaison'])]
            
            nb_eq_sans_mod = len(equipements_sans_modele)
            nb_mod_sans_eq = len(modeles_sans_equipement)


            # --- AFFICHAGE DES RÉSULTATS ---
            
            # Affichage des compteurs bien en évidence avec st.metric
            st.subheader("Résumé du rapprochement")
            col1, col2, col3 = st.columns(3)
            
            col1.metric("Total Équipements", len(df_equipements))
            col2.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)
            col3.metric("🟡 Modèles SANS équipement", nb_mod_sans_eq)
            
            st.divider()

            # Affichage des données orphelines pour que l'utilisateur puisse investiguer
            col_gauche, col_droite = st.columns(2)
            
            with col_gauche:
                st.subheader(f"Détail : {nb_eq_sans_mod} Équipements sans modèle")
                if nb_eq_sans_mod > 0:
                    # On affiche le dataframe, on peut sélectionner juste quelques colonnes pour la clarté
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                else:
                    st.success("Parfait ! Tous les équipements ont un modèle associé.")
                    
            with col_droite:
                st.subheader(f"Détail : {nb_mod_sans_eq} Modèles sans équipement")
                if nb_mod_sans_eq > 0:
                    st.dataframe(modeles_sans_equipement[[col_mod, 'Code de liaison']], use_container_width=True)
                else:
                    st.success("Parfait ! Tous les modèles sont utilisés par au moins un équipement.")

        else:
            st.error(f"⚠️ Les colonnes nécessaires sont introuvables.\n"
                     f"- Attendue dans les Équipements : '{col_eq}'\n"
                     f"- Attendue dans les Modèles : '{col_mod}'")

    except Exception as e:
        st.error(f"Une erreur s'est produite lors de l'analyse : {e}")

else:
    st.info("👈 Veuillez charger le fichier des Équipements ET le fichier des Modèles dans le menu de gauche pour lancer l'analyse croisée.")
