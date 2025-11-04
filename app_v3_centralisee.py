import streamlit as st
import pandas as pd
import io
import pickle
import os
from datetime import datetime

st.set_page_config(
    page_title="Gestion Factures - Système Centralisé",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-title {
        color: white;
        font-size: 2.5rem;
        font-weight: bold;
        margin: 0;
        text-align: center;
    }
    .main-subtitle {
        color: #f0f0f0;
        font-size: 1.2rem;
        text-align: center;
        margin-top: 0.5rem;
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Fichiers
FICHIER_CENTRAL = "Base_Centrale_Cocody.xlsx"
FICHIER_BT_PRINCIPAL = "FACTURAT_ELECTRICITE_BT.xlsx"
FICHIER_HT_PRINCIPAL = "FACTURAT_ELECTRICITE_HT.xlsx"

# Fichiers de sauvegarde
SAVE_FILE_CENTRAL = "data_centrale.pkl"

# Fonctions de chargement
def load_central():
    if os.path.exists(SAVE_FILE_CENTRAL):
        with open(SAVE_FILE_CENTRAL, 'rb') as f:
            return pickle.load(f)
    elif os.path.exists(FICHIER_CENTRAL):
        return pd.read_excel(FICHIER_CENTRAL)
    else:
        st.error(f"❌ Fichier central '{FICHIER_CENTRAL}' introuvable !")
        st.stop()

def save_central(df):
    with open(SAVE_FILE_CENTRAL, 'wb') as f:
        pickle.dump(df, f)

def load_fichier_principal(type_tension):
    fichier = FICHIER_BT_PRINCIPAL if type_tension == "BT" else FICHIER_HT_PRINCIPAL
    if os.path.exists(fichier):
        return pd.read_excel(fichier)
    else:
        st.error(f"❌ Fichier principal '{fichier}' introuvable !")
        return None

# Initialisation
if 'df_central' not in st.session_state:
    st.session_state.df_central = load_central()

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title">📊 Gestion Centralisée des Factures</h1>
    <p class="main-subtitle">Système avec historique mensuel - COCODY</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 📋 Navigation")
    
    page = st.radio(
        "Menu principal",
        ["📊 Base Centrale", "🔄 Import Factures", "📈 Statistiques", "⚙️ Génération Fichiers"]
    )
    
    st.markdown("---")
    st.markdown("### 📊 Informations")
    
    df_central = st.session_state.df_central
    st.metric("📝 Sites", len(df_central))
    
    # Compter les colonnes de type MONTANT_
    cols_montant = [col for col in df_central.columns if col.startswith('MONTANT_')]
    st.metric("📅 Mois enregistrés", len(cols_montant))

# CONTENU PRINCIPAL
if page == "📊 Base Centrale":
    st.markdown("## 📊 Base Centrale - Historique Complet")
    st.markdown("---")
    
    # Statistiques
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">🏢</h3>
            <h2>{len(df_central)}</h2>
            <p style="color: #666;">Sites totaux</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        uc = df_central['UC'].nunique() if 'UC' in df_central.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">🏘️</h3>
            <h2>{uc}</h2>
            <p style="color: #666;">Unités commerciales</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        cols_montant = [col for col in df_central.columns if col.startswith('MONTANT_')]
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">📅</h3>
            <h2>{len(cols_montant)}</h2>
            <p style="color: #666;">Périodes</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        bt_count = len(df_central[df_central['TENSION'] == 'BASSE']) if 'TENSION' in df_central.columns else 0
        ht_count = len(df_central[df_central['TENSION'] == 'HAUTE']) if 'TENSION' in df_central.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">⚡</h3>
            <h2>{bt_count} BT / {ht_count} HT</h2>
            <p style="color: #666;">Répartition</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Tableau
    st.markdown("### 📋 Données de la base centrale")
    
    edited_df = st.data_editor(
        df_central,
        use_container_width=True,
        num_rows="dynamic",
        height=500,
        key="editor_central"
    )
    
    # Actions
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("💾 Sauvegarder", type="primary", use_container_width=True):
            st.session_state.df_central = edited_df
            save_central(edited_df)
            st.success("✅ Base centrale sauvegardée !")
            st.rerun()
    
    with col2:
        if st.button("📥 Exporter Excel", use_container_width=True):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_central.to_excel(writer, index=False, sheet_name='Base_Centrale')
            output.seek(0)
            
            st.download_button(
                "📥 Télécharger",
                data=output,
                file_name=f"Base_Centrale_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_central"
            )
    
    with col3:
        if st.button("🔄 Actualiser", use_container_width=True):
            st.rerun()

elif page == "🔄 Import Factures":
    st.markdown("## 🔄 Import des Factures Mensuelles")
    st.markdown("---")
    
    # Sélection du type
    type_tension = st.radio(
        "Type de facture",
        ["🔌 Basse Tension (BT)", "⚡ Haute Tension (HT)"],
        horizontal=True
    )
    
    is_bt = "BT" in type_tension
    
    st.markdown(f"""
    <div style='background: linear-gradient(135deg, {'#667eea, #764ba2' if is_bt else '#f093fb, #f5576c'}); 
                color: white; 
                padding: 1.5rem; 
                border-radius: 10px;
                margin: 1rem 0;'>
        <h3 style='margin: 0;'>{'🔌 BASSE TENSION' if is_bt else '⚡ HAUTE TENSION'}</h3>
        <p style='margin: 0.5rem 0 0 0;'>Import mensuel des factures</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info(f"""
    📌 **Configuration automatique {'BT' if is_bt else 'HT'}** :
    - Clé fichier source : **{'reference contrat' if is_bt else 'refraccord'}**
    - Clé base centrale : **IDENTIFIANT**
    - Données à importer : **Montant, Consommation, Date**
    """)
    
    # Upload fichier
    fichier_facture = st.file_uploader(
        f"Sélectionnez le fichier de factures {'BT' if is_bt else 'HT'}",
        type=['xlsx', 'xls'],
        key=f"upload_facture_{'bt' if is_bt else 'ht'}"
    )
    
    if fichier_facture:
        try:
            df_facture = pd.read_excel(fichier_facture)
            
            st.success(f"✅ Fichier chargé : {len(df_facture)} ligne(s)")
            
            # Configuration des colonnes
            cle_facture = "reference contrat" if is_bt else "refraccord"
            montant_col = "Montant facture TTC" if is_bt else "montfact"
            conso_col = "conso" if is_bt else "conso"
            date_col = "caract"
            
            # Vérifications
            colonnes_manquantes = []
            for col in [cle_facture, montant_col, date_col]:
                if col not in df_facture.columns:
                    colonnes_manquantes.append(col)
            
            if colonnes_manquantes:
                st.error(f"❌ Colonnes manquantes : {', '.join(colonnes_manquantes)}")
                st.info(f"📋 Colonnes disponibles : {', '.join(df_facture.columns)}")
            else:
                # Récupérer la période (caract)
                periode = df_facture[date_col].dropna().unique()
                if len(periode) > 0:
                    periode = str(periode[0])
                    st.success(f"✅ Période détectée : **{periode}**")
                else:
                    periode = datetime.now().strftime("%m/%Y")
                    st.warning(f"⚠️ Période par défaut : {periode}")
                
                # Aperçu
                with st.expander("👁️ Aperçu du fichier de factures"):
                    cols_to_show = [cle_facture, montant_col, date_col]
                    if conso_col in df_facture.columns:
                        cols_to_show.insert(2, conso_col)
                    st.dataframe(df_facture[cols_to_show].head(10), use_container_width=True)
                
                st.markdown("---")
                
                # Bouton import
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button(f"🔄 LANCER L'IMPORT {'BT' if is_bt else 'HT'}", type="primary", use_container_width=True):
                        with st.spinner("⏳ Import en cours..."):
                            df_central = st.session_state.df_central
                            
                            # Créer les noms de colonnes pour cette période
                            periode_clean = periode.replace('/', '_')
                            col_montant = f"MONTANT_{periode_clean}"
                            col_conso = f"CONSO_{periode_clean}"
                            col_date = f"DATE_{periode_clean}"
                            
                            # Ajouter les colonnes si elles n'existent pas
                            if col_montant not in df_central.columns:
                                df_central[col_montant] = None
                            if col_conso not in df_central.columns:
                                df_central[col_conso] = None
                            if col_date not in df_central.columns:
                                df_central[col_date] = None
                            
                            # Créer un dictionnaire pour la correspondance
                            dict_montant = dict(zip(df_facture[cle_facture].astype(str), df_facture[montant_col]))
                            dict_conso = {}
                            if conso_col in df_facture.columns:
                                dict_conso = dict(zip(df_facture[cle_facture].astype(str), df_facture[conso_col]))
                            
                            # Mise à jour
                            nb_maj = 0
                            for idx, row in df_central.iterrows():
                                identifiant = str(row['IDENTIFIANT'])
                                if identifiant in dict_montant:
                                    df_central.at[idx, col_montant] = dict_montant[identifiant]
                                    if identifiant in dict_conso:
                                        df_central.at[idx, col_conso] = dict_conso[identifiant]
                                    df_central.at[idx, col_date] = periode
                                    nb_maj += 1
                            
                            # Sauvegarder
                            st.session_state.df_central = df_central
                            save_central(df_central)
                            
                            # Résultats
                            st.markdown("---")
                            st.success(f"🎉 Import terminé : {nb_maj} ligne(s) mises à jour !")
                            
                            col_r1, col_r2, col_r3 = st.columns(3)
                            with col_r1:
                                st.metric("✅ Importées", nb_maj)
                            with col_r2:
                                st.metric("⚠️ Non trouvées", len(df_central) - nb_maj)
                            with col_r3:
                                taux = (nb_maj / len(df_central) * 100) if len(df_central) > 0 else 0
                                st.metric("📈 Taux", f"{taux:.1f}%")
                            
                            st.balloons()
        
        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")
            st.exception(e)

elif page == "📈 Statistiques":
    st.markdown("## 📈 Statistiques et Évolution")
    st.markdown("---")
    
    df_central = st.session_state.df_central
    
    # Récupérer toutes les périodes
    cols_montant = [col for col in df_central.columns if col.startswith('MONTANT_')]
    
    if len(cols_montant) == 0:
        st.warning("⚠️ Aucune période enregistrée. Importez d'abord des factures.")
    else:
        st.success(f"✅ {len(cols_montant)} période(s) disponible(s)")
        
        # Sélection de périodes
        periodes = [col.replace('MONTANT_', '') for col in cols_montant]
        
        col1, col2 = st.columns(2)
        with col1:
            periode1 = st.selectbox("Période 1", periodes, index=0)
        with col2:
            periode2 = st.selectbox("Période 2", periodes, index=min(1, len(periodes)-1))
        
        # Calculs
        col_m1 = f"MONTANT_{periode1}"
        col_m2 = f"MONTANT_{periode2}"
        
        total1 = df_central[col_m1].sum() if col_m1 in df_central.columns else 0
        total2 = df_central[col_m2].sum() if col_m2 in df_central.columns else 0
        evolution = ((total2 - total1) / total1 * 100) if total1 > 0 else 0
        
        # Affichage
        st.markdown("### 💰 Comparaison des montants")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(f"💰 {periode1}", f"{total1:,.0f} FCFA")
        with col2:
            st.metric(f"💰 {periode2}", f"{total2:,.0f} FCFA")
        with col3:
            st.metric("📈 Évolution", f"{evolution:+.1f}%", delta=f"{total2-total1:,.0f} FCFA")
        
        # Tableau comparatif
        st.markdown("### 📊 Détails par site")
        
        df_compare = df_central[['UC', 'CORRESPONDANCE', col_m1, col_m2]].copy()
        df_compare['EVOLUTION'] = df_compare[col_m2] - df_compare[col_m1]
        df_compare['EVOLUTION_%'] = ((df_compare[col_m2] - df_compare[col_m1]) / df_compare[col_m1] * 100).fillna(0)
        
        st.dataframe(df_compare, use_container_width=True, height=400)

elif page == "⚙️ Génération Fichiers":
    st.markdown("## ⚙️ Génération des Fichiers Principaux")
    st.markdown("---")
    
    st.info("""
    📌 Cette fonction génère les fichiers BT et HT principaux à partir de la base centrale.
    Les fichiers générés peuvent être utilisés pour l'import dans votre système comptable.
    """)
    
    # À implémenter selon vos besoins
    st.warning("🚧 Fonctionnalité en cours de développement")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; padding: 2rem; color: #666;'>
    <p><strong>Système Centralisé</strong> - Version 3.0 - Historique Mensuel</p>
</div>
""", unsafe_allow_html=True)
