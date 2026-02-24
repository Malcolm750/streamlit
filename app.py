import streamlit as st
import pandas as pd
from io import BytesIO

# Configuration de la page (mode large)
st.set_page_config(page_title="Tableau de bord - Équipements", layout="wide")

st.title("📊 Tableau de Bord : Analyse du Parc d'Équipements")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS ---
st.sidebar.header("📁 Chargement des fichiers")
fichier_equipements = st.sidebar.file_uploader("Fichier des Équipements", type=["xlsx", "xls"])
fichier_models = st.sidebar.file_uploader("Fichier des Modèles", type=["xlsx", "xls"])

# --- FONCTION D'EXPORT EXCEL ---
@st.cache_data
def convertir_df_en_excel(df, sheet_name="Export"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None and fichier_models is not None:
    try:
        # Lecture et nettoyage basique
        df_equipements = pd.read_excel(fichier_equipements)
        df_models = pd.read_excel(fichier_models)
        
        df_equipements.columns = df_equipements.columns.str.strip()
        df_models.columns = df_models.columns.str.strip()

        # Noms des colonnes cibles
        col_eq = 'Modèle (Référence)'
        col_mod = 'Code référence'
        col_nature = "Nature de l'objet"

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            # 1. PRÉPARATION DES CLÉS DE LIAISON
            df_equipements['Code de liaison'] = df_equipements[col_eq].astype(str).str.split(',').str[0].str.strip()
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']

            # 2. JOINTURE DES DONNÉES (On ramène la nature de l'objet vers les équipements)
            df_fusion = pd.merge(
                df_equipements, 
                df_models[['Code de liaison', col_nature]], 
                on='Code de liaison', 
                how='left'
            )
            df_fusion[col_nature] = df_fusion[col_nature].fillna("Non défini / Sans modèle")

            # 3. CALCUL DES ANOMALIES
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            nb_eq_sans_mod = len(equipements_sans_modele)
            
            # --- CRÉATION DES ONGLETS POUR UNE MEILLEURE PRÉSENTATION ---
            tab_apercu, tab_qualite = st.tabs(["📈 Vue d'ensemble du parc", "⚠️ Qualité des données & Exports"])
            
            # ==========================================
            # ONGLET 1 : VUE D'ENSEMBLE (BUSINESS)
            # ==========================================
            with tab_apercu:
                st.header("Indicateurs clés")
                
                # Mise en avant des indicateurs pertinents pour les équipements
                col1, col2, col3 = st.columns(3)
                col1.metric("Total des Équipements", len(df_equipements))
                col2.metric("✅ Équipements rattachés à un modèle", len(df_equipements) - nb_eq_sans_mod)
                col3.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)
                
                st.divider()

                # Analyse par Nature de l'objet
                st.subheader("Répartition par Nature de l'objet")
                
                if col_nature in df_models.columns:
                    analyse_nature = df_fusion.groupby(col_nature).agg(
                        nb_equipements=('Code de liaison', 'size'),
                        nb_modeles_uniques=('Code de liaison', 'nunique')
                    ).reset_index()
                    
                    analyse_nature.columns = ["Nature de l'objet", "Nombre d'équipements", "Nombre de modèles distincts"]
                    analyse_nature = analyse_nature.sort_values(by="Nombre d'équipements", ascending=False)
                    
                    col_chart1, col_chart2 = st.columns([1, 2])
                    with col_chart1:
                        st.dataframe(analyse_nature, use_container_width=True, hide_index=True)
                    with col_chart2:
                        st.bar_chart(analyse_nature.set_index("Nature de l'objet"))
                else:
                    st.error(f"La colonne '{col_nature}' est introuvable.")

            # ==========================================
            # ONGLET 2 : QUALITÉ DES DONNÉES (CORRECTIONS)
            # ==========================================
            with tab_qualite:
                st.header("Actions correctives à mener")
                
                # Bloc 1 : Équipements sans modèle
                st.subheader(f"🔴 Équipements orphelins ({nb_eq_sans_mod})")
                st.write("Ces équipements n'ont pas de modèle reconnu dans la base. Le code extrait ne correspond à aucun 'Code référence'.")
                
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                    # Optionnel : bouton pour télécharger cette liste d'erreurs
                    excel_err = convertir_df_en_excel(equipements_sans_modele, "Equipements en erreur")
                    st.download_button("📥 Exporter les équipements sans modèle", excel_err, "equipements_sans_modele.xlsx")
                else:
                    st.success("Parfait ! Tous les équipements sont correctement liés à un modèle.")
                
                st.divider()

                # Bloc 2 : Le focus spécifique demandé sur les modèles techniques non utilisés
                st.subheader("🛠️ Nettoyage : Modèles 'Équipement Technique' inutilisés")
                st.write("Action spécifique : identifier les modèles de type 'Équipement Technique' qui ne sont affectés à aucun équipement actuel.")
                
                if col_nature in df_models.columns:
                    modeles_sans_equipement = df_models[~df_models['Code de liaison'].isin(df_equipements['Code de liaison'])]
                    modeles_tech_orphelins = modeles_sans_equipement[
                        modeles_sans_equipement[col_nature].astype(str).str.strip().str.title() == "Équipement Technique"
                    ]
                    
                    if len(modeles_tech_orphelins) > 0:
                        st.warning(f"**{len(modeles_tech_orphelins)}** modèles 'Équipement Technique' sont actuellement inutilisés.")
                        excel_tech = convertir_df_en_excel(modeles_tech_orphelins, "Modèles Tech Inutilisés")
                        st.download_button(
                            label="📥 Exporter cette liste pour nettoyage",
                            data=excel_tech,
                            file_name="modeles_equipement_technique_inutilises.xlsx"
                        )
                    else:
                        st.success("Aucun modèle 'Équipement Technique' n'est inutilisé.")

        else:
            st.error(f"⚠️ Les colonnes de liaison sont introuvables. Vérifiez '{col_eq}' et '{col_mod}'.")

    except Exception as e:
        st.error(f"Une erreur s'est produite lors de l'analyse : {e}")

else:
    # Message d'accueil propre quand rien n'est chargé
    st.info("👋 Bienvenue sur votre tableau de bord. Veuillez charger vos fichiers Excel dans le menu latéral pour commencer l'analyse.")
