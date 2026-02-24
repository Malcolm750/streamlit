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

# --- FONCTION DE NETTOYAGE EXTRÊME (POUR LES DOUBLONS) ---
def nettoyer_texte_doublon(series):
    # Convertit en texte, met en minuscules, et supprime tout ce qui n'est pas lettre ou chiffre
    return series.astype(str).str.lower().str.replace(r'[^a-z0-9]', '', regex=True)

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
        col_libelle = "Libellé Référence"
        col_fournisseur = "Fournisseur"

        if col_eq in df_equipements.columns and col_mod in df_models.columns:
            
            # --- 1. PRÉPARATION DES CLÉS DE LIAISON ---
            df_equipements['Code de liaison'] = df_equipements[col_eq].astype(str).str.split(',').str[0].str.strip()
            df_models['Code de liaison'] = df_models[col_mod].astype(str).str.strip()
            
            df_equipements = df_equipements[df_equipements['Code de liaison'] != 'nan']
            df_models = df_models[df_models['Code de liaison'] != 'nan']

            # --- 2. CALCUL DU NOMBRE D'ÉQUIPEMENTS PAR MODÈLE ---
            # On compte combien de fois chaque code de modèle apparaît dans le fichier équipements
            comptage_equipements_par_modele = df_equipements.groupby('Code de liaison').size().reset_index(name="Nb d'équipements rattachés")
            
            # On injecte ce nombre dans le fichier des modèles
            df_models = pd.merge(df_models, comptage_equipements_par_modele, on='Code de liaison', how='left')
            df_models["Nb d'équipements rattachés"] = df_models["Nb d'équipements rattachés"].fillna(0).astype(int)

            # --- 3. RECHERCHE DES DOUBLONS COMPLEXES ---
            if col_libelle in df_models.columns and col_fournisseur in df_models.columns:
                
                # A. Extraction du nom du fournisseur (tout ce qui est après le premier tiret)
                # Si pas de tiret, on garde tout.
                df_models['Fournisseur_Extrait'] = df_models[col_fournisseur].astype(str).apply(
                    lambda x: x.split('-', 1)[-1] if '-' in x else x
                )
                
                # B. Nettoyage extrême pour comparaison (minuscules, sans espaces, sans tirets)
                df_models['Cle_Libelle'] = nettoyer_texte_doublon(df_models[col_libelle])
                df_models['Cle_Fournisseur'] = nettoyer_texte_doublon(df_models['Fournisseur_Extrait'])
                
                # C. Création de la clé unique de comparaison
                df_models['Cle_Comparaison'] = df_models['Cle_Libelle'] + "_" + df_models['Cle_Fournisseur']
                
                # D. Identification des doublons (on garde toutes les lignes qui ont la même clé)
                masque_doublons = df_models.duplicated(subset=['Cle_Comparaison'], keep=False)
                df_doublons = df_models[masque_doublons].copy()
                
                # On trie pour que les doublons soient affichés les uns sous les autres
                df_doublons = df_doublons.sort_values(by=['Cle_Comparaison', "Nb d'équipements rattachés"], ascending=[True, False])

            # --- 4. JOINTURE CLASSIQUE (pour l'onglet 1) ---
            df_fusion = pd.merge(df_equipements, df_models[['Code de liaison', col_nature]], on='Code de liaison', how='left')
            df_fusion[col_nature] = df_fusion[col_nature].fillna("Non défini / Sans modèle")

            # --- CALCUL DES ANOMALIES ---
            equipements_sans_modele = df_equipements[~df_equipements['Code de liaison'].isin(df_models['Code de liaison'])]
            nb_eq_sans_mod = len(equipements_sans_modele)
            
            # ==========================================
            # CRÉATION DES ONGLETS
            # ==========================================
            tab_apercu, tab_qualite, tab_doublons = st.tabs([
                "📈 Vue d'ensemble du parc", 
                "⚠️ Qualité & Erreurs", 
                "🔄 Gestion des Doublons (Modèles)"
            ])
            
            # --- ONGLET 1 : VUE D'ENSEMBLE ---
            with tab_apercu:
                st.header("Indicateurs clés")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total des Équipements", len(df_equipements))
                col2.metric("✅ Équipements rattachés à un modèle", len(df_equipements) - nb_eq_sans_mod)
                col3.metric("🔴 Équipements SANS modèle", nb_eq_sans_mod)
                
                st.divider()
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

            # --- ONGLET 2 : QUALITÉ DES DONNÉES ---
            with tab_qualite:
                st.header("Actions correctives à mener")
                
                st.subheader(f"🔴 Équipements orphelins ({nb_eq_sans_mod})")
                if nb_eq_sans_mod > 0:
                    st.dataframe(equipements_sans_modele[[col_eq, 'Code de liaison']], use_container_width=True)
                    excel_err = convertir_df_en_excel(equipements_sans_modele, "Equipements en erreur")
                    st.download_button("📥 Exporter les équipements sans modèle", excel_err, "equipements_sans_modele.xlsx")
                else:
                    st.success("Parfait ! Tous les équipements sont correctement liés à un modèle.")
                
                st.divider()

                st.subheader("🛠️ Nettoyage : Modèles 'Équipement Technique' inutilisés")
                if col_nature in df_models.columns:
                    modeles_sans_equipement = df_models[~df_models['Code de liaison'].isin(df_equipements['Code de liaison'])]
                    modeles_tech_orphelins = modeles_sans_equipement[
                        modeles_sans_equipement[col_nature].astype(str).str.strip().str.title() == "Équipement Technique"
                    ]
                    
                    if len(modeles_tech_orphelins) > 0:
                        st.warning(f"**{len(modeles_tech_orphelins)}** modèles 'Équipement Technique' sont actuellement inutilisés.")
                        excel_tech = convertir_df_en_excel(modeles_tech_orphelins, "Modèles Tech Inutilisés")
                        st.download_button("📥 Exporter cette liste pour nettoyage", excel_tech, "modeles_equipement_technique_inutilises.xlsx")
                    else:
                        st.success("Aucun modèle 'Équipement Technique' n'est inutilisé.")

            # --- ONGLET 3 : GESTION DES DOUBLONS ---
            with tab_doublons:
                st.header("Analyse des modèles en doublon")
                
                if col_libelle in df_models.columns and col_fournisseur in df_models.columns:
                    nb_doublons_total = len(df_doublons)
                    nb_groupes_doublons = df_doublons['Cle_Comparaison'].nunique()
                    
                    st.write(f"Nous avons détecté **{nb_doublons_total} modèles** qui semblent être des doublons, répartis en **{nb_groupes_doublons} groupes** d'articles similaires.")
                    
                    if nb_doublons_total > 0:
                        
                        # 1. Analyse par Nature de l'objet
                        st.subheader("Répartition des doublons par Nature")
                        doublons_par_nature = df_doublons[col_nature].value_counts().reset_index()
                        doublons_par_nature.columns = ["Nature de l'objet", "Nombre de modèles impliqués"]
                        st.dataframe(doublons_par_nature, use_container_width=True, hide_index=True)
                        
                        st.divider()
                        
                        # 2. Affichage du détail pour nettoyage
                        st.subheader("Détail des doublons (pour arbitrage)")
                        st.info("💡 Astuce : Les modèles sont regroupés par ressemblance. Regardez la colonne **Nb d'équipements rattachés** pour savoir lequel conserver (idéalement celui qui n'a pas 0).")
                        
                        # On sélectionne les colonnes les plus pertinentes à afficher
                        colonnes_a_afficher = [
                            col_mod, 
                            col_libelle, 
                            col_fournisseur, 
                            col_nature, 
                            "Nb d'équipements rattachés",
                            "Cle_Comparaison" # On l'affiche pour qu'on comprenne pourquoi ils ont été groupés
                        ]
                        
                        st.dataframe(df_doublons[colonnes_a_afficher], use_container_width=True, hide_index=True)
                        
                        # Export Excel
                        excel_doublons = convertir_df_en_excel(df_doublons[colonnes_a_afficher], "Doublons Modèles")
                        st.download_button(
                            label="📥 Exporter la liste des doublons pour nettoyage",
                            data=excel_doublons,
                            file_name="modeles_doublons_a_corriger.xlsx"
                        )
                        
                    else:
                        st.success("Félicitations ! Aucun doublon détecté selon ces critères.")
                else:
                    st.error(f"Impossible de vérifier les doublons. Il manque les colonnes '{col_libelle}' ou '{col_fournisseur}'.")

        else:
            st.error(f"⚠️ Les colonnes de liaison sont introuvables. Vérifiez '{col_eq}' et '{col_mod}'.")

    except Exception as e:
        st.error(f"Une erreur s'est produite lors de l'analyse : {e}")

else:
    st.info("👋 Bienvenue sur votre tableau de bord. Veuillez charger vos fichiers Excel dans le menu latéral pour commencer l'analyse.")
