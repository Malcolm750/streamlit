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
def generer_excel_doublons_multifeuilles(df_doublons):
    output = BytesIO()
    
    # 1. Somme des équipements par groupe de doublons
    group_sums = df_doublons.groupby('Cle_Comparaison')["Nb d'équipements rattachés"].sum()
    groupes_sans_equip = group_sums[group_sums == 0].index
    groupes_avec_equip = group_sums[group_sums > 0].index
    
    # 2. A - Les groupes 100% sans équipement
    df_total_zero = df_doublons[df_doublons['Cle_Comparaison'].isin(groupes_sans_equip)].copy()
    
    # 2. B - Les groupes mixtes (avec au moins 1 équipement au total)
    df_mixte = df_doublons[df_doublons['Cle_Comparaison'].isin(groupes_avec_equip)].copy()
    
    # Ceux à 0 dans les groupes mixtes vont dans "À supprimer"
    df_a_supprimer = df_mixte[df_mixte["Nb d'équipements rattachés"] == 0].copy()
    
    # Ceux > 0 sont les survivants potentiels
    df_survivants = df_mixte[df_mixte["Nb d'équipements rattachés"] > 0].copy()
    
    # --- LA NOUVELLE RÈGLE D'AUTO-RÉSOLUTION ---
    # On compte combien de survivants il reste par groupe
    comptage_survivants = df_survivants.groupby('Cle_Comparaison').size()
    # On ne garde en "À traiter" QUE s'il y a plus d'un survivant (vrai conflit)
    groupes_en_conflit = comptage_survivants[comptage_survivants > 1].index
    df_a_traiter = df_survivants[df_survivants['Cle_Comparaison'].isin(groupes_en_conflit)].copy()
    # ---------------------------------------------
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        def styliser_et_sauvegarder(df_subset, nom_feuille):
            if df_subset.empty:
                pd.DataFrame({'Message': [f'Aucun {nom_feuille}']}).to_excel(writer, index=False, sheet_name=nom_feuille)
                return
            
            unique_keys = df_subset['Cle_Comparaison'].unique()
            color_map = {key: 'background-color: #E2EFDA' if i % 2 == 0 else 'background-color: #FFFFFF' for i, key in enumerate(unique_keys)}
            
            df_to_export = df_subset.drop(columns=['Cle_Comparaison'])
            
            def highlight_rows(row):
                key = df_subset.loc[row.name, 'Cle_Comparaison']
                couleur = color_map.get(key, '')
                return [couleur] * len(row)
                
            styled_df = df_to_export.style.apply(highlight_rows, axis=1)
            styled_df.to_excel(writer, index=False, sheet_name=nom_feuille)
            
            worksheet = writer.sheets[nom_feuille]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column].width = min(max_length + 2, 50)

        styliser_et_sauvegarder(df_a_traiter, 'Doublons à traiter')
        styliser_et_sauvegarder(df_a_supprimer, 'Doublons à supprimer')
        styliser_et_sauvegarder(df_total_zero, 'Doublons sans équipement')

    return output.getvalue()

# --- FONCTION DE NETTOYAGE EXTRÊME ---
def nettoyer_texte_doublon(series):
    return series.astype(str).str.lower().str.replace(r'[^a-z0-9]', '', regex=True)

# --- ANALYSE PRINCIPALE ---
if fichier_equipements is not None and fichier_models is not None:
    try:
        df_equipements = pd.read_excel(fichier_equipements)
        df_models = pd.read_excel(fichier_models)
        
        df_equipements.columns = df_equipements.columns.str.strip()
        df_models.columns = df_models.columns.str.strip()

        col_eq = 'Modèle (Référence)'
        col_mod = 'Code référence'
        col_nature = "Nature de l'objet"
        col_libelle = "Libellé Référence"
        col_fournisseur = "Fournisseur"
        col_ref_const = "Référence constructeur" 

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            df_equipements['Code de liaison'] = df_equipements[col_eq].astype(str).str.split(',').str[0].str.strip()
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']

            comptage_equipements_par_modele = df_equipements.groupby('Code de liaison').size().reset_index(name="Nb d'équipements rattachés")
            df_models = pd.merge(df_models, comptage_equipements_par_modele, on='Code de liaison', how='left')
            df_models["Nb d'équipements rattachés"] = df_models["Nb d'équipements rattachés"].fillna(0).astype(int)

            if col_libelle in df_models.columns:
                
                def extraire_fournisseur(row):
                    fourn = str(row[col_fournisseur]) if col_fournisseur in row.index and pd.notna(row[col_fournisseur]) else ""
                    if fourn.strip().lower() in ["", "nan", "none"]:
                        if col_ref_const in row.index and pd.notna(row[col_ref_const]):
                            return str(row[col_ref_const])
                        return ""
                    else:
                        return fourn.split('-', 1)[-1] if '-' in fourn else fourn
                
                df_models['Fournisseur_Extrait'] = df_models.apply(extraire_fournisseur, axis=1)
                
                df_models['Cle_Libelle'] = nettoyer_texte_doublon(df_models[col_libelle])
                df_models['Cle_Fournisseur'] = nettoyer_texte_doublon(df_models['Fournisseur_Extrait'])
                df_models['Cle_Comparaison'] = df_models['Cle_Libelle'] + "_" + df_models['Cle_Fournisseur']
                
                masque_doublons = df_models.duplicated(subset=['Cle_Comparaison'], keep=False)
                df_doublons = df_models[masque_doublons].copy()
                df_doublons = df_doublons.sort_values(by=['Cle_Comparaison', "Nb d'équipements rattachés"], ascending=[True, False])

            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            nb_eq_sans_mod = len(equipements_sans_modele)
            
            tab_apercu, tab_qualite, tab_doublons = st.tabs([
                "📈 Vue d'ensemble", 
                "⚠️ Qualité & Erreurs", 
                "🔄 Gestion des Doublons"
            ])
            
            with tab_apercu:
                st.header("Indicateurs clés")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total des Équipements", len(df_equipements))
                col2.metric("✅ Équipements rattachés à un modèle", len(df_equipements) - nb_eq_sans_mod)
                col3.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)

            with tab_qualite:
                st.header("Actions correctives à mener")
                st.subheader(f"🔴 Équipements orphelins ({nb_eq_sans_mod})")
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                    st.download_button("📥 Exporter les équipements sans modèle", convertir_df_en_excel(equipements_sans_modele), "equipements_sans_modele.xlsx")

            with tab_doublons:
                st.header("Analyse et Export des modèles en doublon")
                
                if col_libelle in df_models.columns:
                    if len(df_doublons) > 0:
                        st.write("Le système a identifié des doublons et a appliqué l'auto-résolution (masquage des modèles maîtres).")
                        
                        colonnes_techniques_a_masquer = ['Code de liaison', 'Fournisseur_Extrait', 'Cle_Libelle', 'Cle_Fournisseur']
                        colonnes_finales = [col for col in df_doublons.columns if col not in colonnes_techniques_a_masquer]
                        
                        if col_mod in colonnes_finales and "Nb d'équipements rattachés" in colonnes_finales:
                            colonnes_finales.remove("Nb d'équipements rattachés")
                            idx = colonnes_finales.index(col_mod) + 1
                            colonnes_finales.insert(idx, "Nb d'équipements rattachés")
                            
                        df_export_doublons = df_doublons[colonnes_finales]
                        
                        excel_doublons_multifeuilles = generer_excel_doublons_multifeuilles(df_export_doublons)
                        
                        st.download_button(
                            label="📥 Télécharger le fichier des Doublons (Tri Intelligent)",
                            data=excel_doublons_multifeuilles,
                            file_name="modeles_doublons_a_corriger.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                        
                        st.success("💡 **Nouveauté :** Si un modèle est le seul de son groupe de doublons à posséder des équipements, il est considéré comme le 'Modèle Maître'. Ses copies à 0 seront placées dans 'Doublons à supprimer', et lui-même n'apparaitra pas dans le fichier pour éviter de polluer votre travail !")
                    else:
                        st.success("Félicitations ! Aucun doublon détecté.")
                else:
                    st.error(f"Il manque la colonne '{col_libelle}' pour l'analyse.")

        else:
            st.error("⚠️ Les colonnes de liaison sont introuvables.")

    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
