import streamlit as st
import pandas as pd
from io import BytesIO

# Configuration de la page
st.set_page_config(page_title="Tableau de bord - Équipements", layout="wide")

st.title("📊 Tableau de Bord : Analyse du Parc d'Équipements")

# --- BARRE LATÉRALE : CHARGEMENT DES FICHIERS ---
st.sidebar.header("📁 Chargement des fichiers")
fichier_equipements = st.sidebar.file_uploader("Fichier des Équipements", type=["xlsx", "xls"])
fichier_models = st.sidebar.file_uploader("Fichier des Modèles", type=["xlsx", "xls"])

# --- FONCTIONS D'EXPORT EXCEL ---
@st.cache_data
def convertir_df_en_excel(df, sheet_name="Export"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

@st.cache_data
def generer_excel_doublons_colores(df):
    output = BytesIO()
    
    # 1. Créer une carte des couleurs : alternance de Vert clair et Blanc pour chaque groupe
    unique_keys = df['Cle_Comparaison'].unique()
    color_map = {key: 'background-color: #E2EFDA' if i % 2 == 0 else 'background-color: #FFFFFF' for i, key in enumerate(unique_keys)}
    
    # 2. Fonction qui applique la couleur à la ligne entière
    def highlight_rows(row):
        couleur = color_map.get(row['Cle_Comparaison'], '')
        return [couleur] * len(row)
        
    # 3. Application du style au DataFrame
    styled_df = df.style.apply(highlight_rows, axis=1)
    
    # 4. Écriture dans le fichier Excel avec ajustement de la largeur des colonnes
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        styled_df.to_excel(writer, index=False, sheet_name='Doublons_a_traiter')
        worksheet = writer.sheets['Doublons_a_traiter']
        
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter # Lettre de la colonne (A, B, C...)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            # Ajuste la largeur (max 50 pour éviter les colonnes infinies sur des longs textes)
            worksheet.column_dimensions[column].width = min(max_length + 2, 50)
            
    return output.getvalue()

# --- FONCTION DE NETTOYAGE EXTRÊME (POUR LES DOUBLONS) ---
def nettoyer_texte_doublon(series):
    return series.astype(str).str.lower().str.replace(r'[^a-z0-9]', '', regex=True)

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None and fichier_models is not None:
    try:
        # Lecture et nettoyage basique
        df_equipements = pd.read_excel(fichier_equipements)
        df_models = pd.read_excel(fichier_models)
        
        df_equipements.columns = df_equipements.columns.str.strip()
        df_models.columns = df_models.columns.str.strip()

        col_eq = 'Modèle (Référence)'
        col_mod = 'Code référence'
        col_nature = "Nature de l'objet"
        col_libelle = "Libellé Référence"
        col_fournisseur = "Fournisseur"

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            # --- 1. CLÉS DE LIAISON ---
            df_equipements['Code de liaison'] = df_equipements[col_eq].astype(str).str.split(',').str[0].str.strip()
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']

            # --- 2. COMPTAGE ÉQUIPEMENTS / MODÈLE ---
            comptage_equipements_par_modele = df_equipements.groupby('Code de liaison').size().reset_index(name="Nb d'équipements rattachés")
            df_models = pd.merge(df_models, comptage_equipements_par_modele, on='Code de liaison', how='left')
            df_models["Nb d'équipements rattachés"] = df_models["Nb d'équipements rattachés"].fillna(0).astype(int)

            # --- 3. RECHERCHE DES DOUBLONS COMPLEXES ---
            if col_libelle in df_models.columns and col_fournisseur in df_models.columns:
                
                df_models['Fournisseur_Extrait'] = df_models[col_fournisseur].astype(str).apply(
                    lambda x: x.split('-', 1)[-1] if '-' in x else x
                )
                
                df_models['Cle_Libelle'] = nettoyer_texte_doublon(df_models[col_libelle])
                df_models['Cle_Fournisseur'] = nettoyer_texte_doublon(df_models['Fournisseur_Extrait'])
                df_models['Cle_Comparaison'] = df_models['Cle_Libelle'] + "_" + df_models['Cle_Fournisseur']
                
                masque_doublons = df_models.duplicated(subset=['Cle_Comparaison'], keep=False)
                df_doublons = df_models[masque_doublons].copy()
                df_doublons = df_doublons.sort_values(by=['Cle_Comparaison', "Nb d'équipements rattachés"], ascending=[True, False])

            # --- 4. CALCUL DES ANOMALIES GLOBALES ---
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            nb_eq_sans_mod = len(equipements_sans_modele)
            
            # ==========================================
            # ONGLET 1 & 2 & 3
            # ==========================================
            tab_apercu, tab_qualite, tab_doublons = st.tabs([
                "📈 Vue d'ensemble du parc", 
                "⚠️ Qualité & Erreurs", 
                "🔄 Gestion des Doublons"
            ])
            
            with tab_apercu:
                # ... (Même code que précédemment pour l'onglet 1) ...
                st.header("Indicateurs clés")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total des Équipements", len(df_equipements))
                col2.metric("✅ Équipements rattachés à un modèle", len(df_equipements) - nb_eq_sans_mod)
                col3.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)
                st.info("Basculez sur les autres onglets pour nettoyer la base de données.")

            with tab_qualite:
                # ... (Même code que précédemment pour l'onglet 2) ...
                st.header("Actions correctives à mener")
                st.subheader(f"🔴 Équipements orphelins ({nb_eq_sans_mod})")
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                    st.download_button("📥 Exporter les équipements sans modèle", convertir_df_en_excel(equipements_sans_modele), "equipements_sans_modele.xlsx")

            # --- ONGLET 3 : GESTION DES DOUBLONS MISE À JOUR ---
            with tab_doublons:
                st.header("Analyse et Export des modèles en doublon")
                
                if col_libelle in df_models.columns and col_fournisseur in df_models.columns:
                    nb_doublons_total = len(df_doublons)
                    
                    if nb_doublons_total > 0:
                        st.write(f"**{nb_doublons_total} modèles** sont des doublons potentiels.")
                        
                        # PREPARATION DES COLONNES POUR L'EXPORT
                        # On retire les colonnes techniques temporaires qui n'intéressent pas l'utilisateur
                        colonnes_techniques = ['Code de liaison', 'Fournisseur_Extrait', 'Cle_Libelle', 'Cle_Fournisseur']
                        colonnes_finales = [col for col in df_doublons.columns if col not in colonnes_techniques]
                        
                        # ASTUCE : On déplace "Nb d'équipements rattachés" juste après "Code référence" pour que ça soit visible
                        if col_mod in colonnes_finales and "Nb d'équipements rattachés" in colonnes_finales:
                            colonnes_finales.remove("Nb d'équipements rattachés")
                            idx = colonnes_finales.index(col_mod) + 1
                            colonnes_finales.insert(idx, "Nb d'équipements rattachés")
                            
                        # Application du filtre des colonnes
                        df_export_doublons = df_doublons[colonnes_finales]
                        
                        # Affichage sur l'écran (limité aux infos importantes pour ne pas surcharger)
                        st.dataframe(df_export_doublons[[col_mod, "Nb d'équipements rattachés", col_libelle, col_fournisseur]], use_container_width=True, hide_index=True)
                        
                        # Génération du fichier Excel Colorié !
                        excel_doublons_colores = generer_excel_doublons_colores(df_export_doublons)
                        
                        st.download_button(
                            label="📥 Télécharger le fichier complet des Doublons (Avec couleurs et nb d'équipements)",
                            data=excel_doublons_colores,
                            file_name="modeles_doublons_a_corriger.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary" # Met le bouton en couleur bien visible
                        )
                        
                        st.success("💡 Dans le fichier Excel téléchargé, les groupes de doublons sont automatiquement colorés alternativement en vert et blanc pour faciliter votre tri.")
                        
                    else:
                        st.success("Félicitations ! Aucun doublon détecté.")
                else:
                    st.error("Impossible de vérifier les doublons. Il manque des colonnes.")

        else:
            st.error("⚠️ Les colonnes de liaison sont introuvables.")

    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
