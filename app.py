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
    
    # 1. On identifie les groupes de doublons dont la SOMME des équipements rattachés est 0
    group_sums = df_doublons.groupby('Cle_Comparaison')["Nb d'équipements rattachés"].sum()
    groupes_sans_equip = group_sums[group_sums == 0].index
    
    # 2. On sépare en deux DataFrames
    df_sans_equip = df_doublons[df_doublons['Cle_Comparaison'].isin(groupes_sans_equip)].copy()
    df_avec_equip = df_doublons[~df_doublons['Cle_Comparaison'].isin(groupes_sans_equip)].copy()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Sous-fonction pour styliser et écrire une feuille
        def styliser_et_sauvegarder(df_subset, nom_feuille):
            if df_subset.empty:
                pd.DataFrame({'Message': [f'Aucun {nom_feuille}']}).to_excel(writer, index=False, sheet_name=nom_feuille)
                return
            
            # Création de la palette de couleurs par groupe
            unique_keys = df_subset['Cle_Comparaison'].unique()
            color_map = {key: 'background-color: #E2EFDA' if i % 2 == 0 else 'background-color: #FFFFFF' for i, key in enumerate(unique_keys)}
            
            # On supprime la clé de comparaison technique pour ne pas l'afficher dans l'Excel final
            df_to_export = df_subset.drop(columns=['Cle_Comparaison'])
            
            def highlight_rows(row):
                # On retrouve la clé via l'index pour appliquer la bonne couleur
                key = df_subset.loc[row.name, 'Cle_Comparaison']
                couleur = color_map.get(key, '')
                return [couleur] * len(row)
                
            # Application du style et sauvegarde
            styled_df = df_to_export.style.apply(highlight_rows, axis=1)
            styled_df.to_excel(writer, index=False, sheet_name=nom_feuille)
            
            # Ajustement automatique de la largeur des colonnes
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

        # 3. Création des deux feuilles dans le même classeur Excel
        styliser_et_sauvegarder(df_avec_equip, 'Doublons_a_traiter')
        styliser_et_sauvegarder(df_sans_equip, 'doublon sans équipement')

    return output.getvalue()

# --- FONCTION DE NETTOYAGE EXTRÊME (POUR LES DOUBLONS) ---
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
            if col_libelle in df_models.columns:
                
                # FONCTION POUR CASCADER FOURNISSEUR -> RÉFÉRENCE CONSTRUCTEUR
                def extraire_fournisseur(row):
                    fourn = str(row[col_fournisseur]) if col_fournisseur in row.index and pd.notna(row[col_fournisseur]) else ""
                    if fourn.strip().lower() in ["", "nan", "none"]:
                        if col_ref_const in row.index and pd.notna(row[col_ref_const]):
                            return str(row[col_ref_const])
                        return ""
                    else:
                        return fourn.split('-', 1)[-1] if '-' in fourn else fourn
                
                # Application de la règle d'extraction
                df_models['Fournisseur_Extrait'] = df_models.apply(extraire_fournisseur, axis=1)
                
                # Nettoyage et comparaison
                df_models['Cle_Libelle'] = nettoyer_texte_doublon(df_models[col_libelle])
                df_models['Cle_Fournisseur'] = nettoyer_texte_doublon(df_models['Fournisseur_Extrait'])
                df_models['Cle_Comparaison'] = df_models['Cle_Libelle'] + "_" + df_models['Cle_Fournisseur']
                
                masque_doublons = df_models.duplicated(subset=['Cle_Comparaison'], keep=False)
                df_doublons = df_models[masque_doublons].copy()
                df_doublons = df_doublons.sort_values(by=['Cle_Comparaison', "Nb d'équipements rattachés"], ascending=[True, False])

            # --- CALCUL DES ANOMALIES GLOBALES ---
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            nb_eq_sans_mod = len(equipements_sans_modele)
            
            # ==========================================
            # ONGLETS
            # ==========================================
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
                st.info("Basculez sur les autres onglets pour nettoyer la base de données.")

            with tab_qualite:
                st.header("Actions correctives à mener")
                st.subheader(f"🔴 Équipements orphelins ({nb_eq_sans_mod})")
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                    st.download_button("📥 Exporter les équipements sans modèle", convertir_df_en_excel(equipements_sans_modele), "equipements_sans_modele.xlsx")

            # --- ONGLET 3 : GESTION DES DOUBLONS ---
            with tab_doublons:
                st.header("Analyse et Export des modèles en doublon")
                
                if col_libelle in df_models.columns:
                    nb_doublons_total = len(df_doublons)
                    
                    if nb_doublons_total > 0:
                        st.write(f"**{nb_doublons_total} modèles** sont des doublons potentiels.")
                        
                        # PREPARATION DES COLONNES
                        colonnes_techniques_a_masquer = ['Code de liaison', 'Fournisseur_Extrait', 'Cle_Libelle', 'Cle_Fournisseur']
                        colonnes_finales = [col for col in df_doublons.columns if col not in colonnes_techniques_a_masquer]
                        
                        # Placement stratégique de la colonne
                        if col_mod in colonnes_finales and "Nb d'équipements rattachés" in colonnes_finales:
                            colonnes_finales.remove("Nb d'équipements rattachés")
                            idx = colonnes_finales.index(col_mod) + 1
                            colonnes_finales.insert(idx, "Nb d'équipements rattachés")
                            
                        # Application du filtre
                        df_export_doublons = df_doublons[colonnes_finales]
                        
                        # --- CORRECTION ICI ---
                        # On sélectionne dynamiquement les colonnes à afficher pour éviter l'erreur
                        colonnes_affichage = [col_mod, "Nb d'équipements rattachés", col_libelle]
                        if col_fournisseur in df_export_doublons.columns:
                            colonnes_affichage.append(col_fournisseur)
                        elif col_ref_const in df_export_doublons.columns:
                            colonnes_affichage.append(col_ref_const)
                        
                        # Affichage à l'écran
                        st.dataframe(df_export_doublons[colonnes_affichage], use_container_width=True, hide_index=True)
                        # ----------------------
                        
                        # Génération du fichier Excel MULTI-FEUILLES et COLORIÉ
                        excel_doublons_multifeuilles = generer_excel_doublons_multifeuilles(df_export_doublons)
                        
                        st.download_button(
                            label="📥 Télécharger le fichier des Doublons (Multi-feuilles)",
                            data=excel_doublons_multifeuilles,
                            file_name="modeles_doublons_a_corriger.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                        
                        st.success("💡 Le fichier Excel contient désormais deux onglets : un pour les doublons à arbitrer, et un onglet 'doublon sans équipement' pour les modèles 100% inutilisés.")
                    else:
                        st.success("Félicitations ! Aucun doublon détecté.")
                else:
                    st.error(f"Il manque la colonne '{col_libelle}' pour procéder à l'analyse des doublons.")

        else:
            st.error("⚠️ Les colonnes de liaison sont introuvables.")

    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
