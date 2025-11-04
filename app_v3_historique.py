import streamlit as st
import pandas as pd
import io
import pickle
import os
from datetime import datetime

st.set_page_config(
    page_title="Gestion Factures - Historique par Ligne",
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
SAVE_FILE_CENTRAL = "data_centrale.pkl"

# Fonctions de chargement
def load_central():
    if os.path.exists(SAVE_FILE_CENTRAL):
        with open(SAVE_FILE_CENTRAL, 'rb') as f:
            return pickle.load(f)
    elif os.path.exists(FICHIER_CENTRAL):
        df = pd.read_excel(FICHIER_CENTRAL)
        # S'assurer que les colonnes nécessaires existent
        if 'MONTANT' not in df.columns:
            df['MONTANT'] = None
        if 'CONSO' not in df.columns:
            df['CONSO'] = None
        if 'DATE' not in df.columns:
            df['DATE'] = None
        return df
    else:
        st.error(f"❌ Fichier central '{FICHIER_CENTRAL}' introuvable !")
        st.stop()

def save_central(df):
    with open(SAVE_FILE_CENTRAL, 'wb') as f:
        pickle.dump(df, f)

# Initialisation
if 'df_central' not in st.session_state:
    st.session_state.df_central = load_central()

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title">📊 Gestion Centralisée des Factures</h1>
    <p class="main-subtitle" style="color: #f0f0f0; text-align: center; margin-top: 0.5rem;">
        Système avec historique mensuel par ligne - COCODY
    </p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 📋 Navigation")
    
    page = st.radio(
        "Menu principal",
        ["📊 Base Centrale", "🔄 Import Factures BT", "🔄 Import Factures HT", "📈 Statistiques"]
    )
    
    st.markdown("---")
    st.markdown("### 📊 Informations")
    
    df_central = st.session_state.df_central
    st.metric("📝 Lignes totales", len(df_central))
    
    # Compter les périodes uniques
    if 'DATE' in df_central.columns:
        periodes = df_central['DATE'].dropna().nunique()
        st.metric("📅 Périodes", periodes)

# CONTENU PRINCIPAL
if page == "📊 Base Centrale":
    st.markdown("## 📊 Base Centrale - Historique Complet")
    st.markdown("---")
    
    df_central = st.session_state.df_central
    
    # Statistiques
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">📝</h3>
            <h2>{len(df_central)}</h2>
            <p style="color: #666;">Lignes totales</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        sites_uniques = df_central['IDENTIFIANT'].nunique() if 'IDENTIFIANT' in df_central.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">🏢</h3>
            <h2>{sites_uniques}</h2>
            <p style="color: #666;">Sites uniques</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        periodes = df_central['DATE'].dropna().nunique() if 'DATE' in df_central.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">📅</h3>
            <h2>{periodes}</h2>
            <p style="color: #666;">Périodes</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        total = df_central['MONTANT'].sum() if 'MONTANT' in df_central.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea;">💰</h3>
            <h2>{total/1000:.0f}K</h2>
            <p style="color: #666;">Total FCFA</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Filtres
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        if 'UC' in df_central.columns:
            ucs = ['Tous'] + sorted(df_central['UC'].dropna().unique().tolist())
            uc_filter = st.selectbox("Filtrer par UC", ucs)
        else:
            uc_filter = 'Tous'
    
    with col_f2:
        if 'DATE' in df_central.columns:
            dates = ['Tous'] + sorted(df_central['DATE'].dropna().unique().tolist(), reverse=True)
            date_filter = st.selectbox("Filtrer par DATE", dates)
        else:
            date_filter = 'Tous'
    
    with col_f3:
        if 'TENSION' in df_central.columns:
            tensions = ['Tous'] + sorted(df_central['TENSION'].dropna().unique().tolist())
            tension_filter = st.selectbox("Filtrer par TENSION", tensions)
        else:
            tension_filter = 'Tous'
    
    # Appliquer les filtres
    df_filtered = df_central.copy()
    if uc_filter != 'Tous' and 'UC' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['UC'] == uc_filter]
    if date_filter != 'Tous' and 'DATE' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['DATE'] == date_filter]
    if tension_filter != 'Tous' and 'TENSION' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['TENSION'] == tension_filter]
    
    st.markdown(f"### 📋 Données filtrées ({len(df_filtered)} ligne(s))")
    
    # Tableau
    edited_df = st.data_editor(
        df_filtered,
        use_container_width=True,
        num_rows="dynamic",
        height=500,
        key="editor_central"
    )
    
    # Actions
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("💾 Sauvegarder", type="primary", use_container_width=True):
            # Mettre à jour seulement les lignes filtrées dans le df complet
            for idx in edited_df.index:
                if idx in st.session_state.df_central.index:
                    st.session_state.df_central.loc[idx] = edited_df.loc[idx]
            save_central(st.session_state.df_central)
            st.success("✅ Base centrale sauvegardée !")
            st.rerun()
    
    with col2:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name='Base_Centrale')
        output.seek(0)
        
        st.download_button(
            "📥 Exporter Excel",
            data=output,
            file_name=f"Base_Centrale_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_central"
        )
    
    with col3:
        if st.button("🔄 Actualiser", use_container_width=True):
            st.rerun()

elif page == "🔄 Import Factures BT":
    st.markdown("## 🔄 Import Factures - Basse Tension (BT)")
    st.markdown("---")
    
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea, #764ba2); 
                color: white; 
                padding: 1.5rem; 
                border-radius: 10px;
                margin: 1rem 0;'>
        <h3 style='margin: 0;'>🔌 BASSE TENSION</h3>
        <p style='margin: 0.5rem 0 0 0;'>Import mensuel - Ajout de nouvelles lignes</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    📌 **Configuration BT** :
    - Clé facture : **reference contrat**
    - Clé base centrale : **IDENTIFIANT**
    - Données : **Montant facture TTC**, **conso**, **caract** (période)
    
    💡 Pour chaque ligne trouvée, une **nouvelle ligne** sera ajoutée dans la base centrale avec les données du mois.
    """)
    
    # Upload fichier
    fichier_bt = st.file_uploader(
        "Sélectionnez le fichier de factures BT",
        type=['xlsx', 'xls'],
        key="upload_bt"
    )
    
    if fichier_bt:
        try:
            df_bt = pd.read_excel(fichier_bt)
            
            st.success(f"✅ Fichier chargé : {len(df_bt)} ligne(s)")
            
            # Configuration des colonnes
            cle_facture = "reference contrat"
            montant_col = "Montant facture TTC"
            conso_col = "conso"
            caract_col = "caract"
            
            # Vérifications
            colonnes_manquantes = []
            for col in [cle_facture, montant_col, caract_col]:
                if col not in df_bt.columns:
                    colonnes_manquantes.append(col)
            
            if colonnes_manquantes:
                st.error(f"❌ Colonnes manquantes : {', '.join(colonnes_manquantes)}")
                st.info(f"📋 Colonnes disponibles : {', '.join(df_bt.columns)}")
            else:
                # Récupérer la période
                periode_bt = df_bt[caract_col].dropna().unique()
                if len(periode_bt) > 0:
                    periode_bt = str(periode_bt[0])
                    st.success(f"✅ Période BT détectée : **{periode_bt}**")
                else:
                    periode_bt = ""
                    st.warning("⚠️ Aucune période détectée")
                
                # Aperçu
                with st.expander("👁️ Aperçu du fichier BT"):
                    cols_to_show = [cle_facture, montant_col, caract_col]
                    if conso_col in df_bt.columns:
                        cols_to_show.insert(2, conso_col)
                    st.dataframe(df_bt[cols_to_show].head(10), use_container_width=True)
                
                st.markdown("---")
                
                # Bouton import
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🔄 LANCER L'IMPORT BT", type="primary", use_container_width=True):
                        with st.spinner("⏳ Import BT en cours..."):
                            df_central = st.session_state.df_central.copy()
                            
                            # Créer un template de base centrale (garder les colonnes de structure)
                            colonnes_structure = ['UC', 'CODE AGCE', 'SITES', 'CORRESPONDANCE', 
                                                'IDENTIFIANT', 'REFERENCE', 'TENSION']
                            
                            nouvelles_lignes = []
                            nb_ajouts = 0
                            
                            # Pour chaque ligne de facture
                            for _, row_facture in df_bt.iterrows():
                                ref_contrat = str(row_facture[cle_facture])
                                
                                # Chercher dans la base centrale
                                ligne_centrale = df_central[df_central['IDENTIFIANT'].astype(str) == ref_contrat]
                                
                                if not ligne_centrale.empty:
                                    # Prendre la première occurrence
                                    ligne_base = ligne_centrale.iloc[0].copy()
                                    
                                    # Mettre à jour les valeurs
                                    ligne_base['MONTANT'] = row_facture[montant_col]
                                    ligne_base['DATE'] = periode_bt
                                    
                                    if conso_col in df_bt.columns:
                                        ligne_base['CONSO'] = row_facture.get(conso_col, None)
                                    
                                    nouvelles_lignes.append(ligne_base)
                                    nb_ajouts += 1
                            
                            if nouvelles_lignes:
                                # Créer un DataFrame avec les nouvelles lignes
                                df_nouvelles = pd.DataFrame(nouvelles_lignes)
                                
                                # Ajouter au DataFrame central
                                df_central = pd.concat([df_central, df_nouvelles], ignore_index=True)
                                
                                # Sauvegarder
                                st.session_state.df_central = df_central
                                save_central(df_central)
                                
                                # Résultats
                                st.markdown("---")
                                st.success(f"🎉 Import BT terminé : {nb_ajouts} ligne(s) ajoutée(s) !")
                                
                                col_r1, col_r2, col_r3 = st.columns(3)
                                with col_r1:
                                    st.metric("✅ Lignes ajoutées", nb_ajouts)
                                with col_r2:
                                    st.metric("📊 Total lignes", len(df_central))
                                with col_r3:
                                    st.metric("📅 Période", periode_bt)
                                
                                st.balloons()
                            else:
                                st.warning("⚠️ Aucune correspondance trouvée")
        
        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")
            st.exception(e)

elif page == "🔄 Import Factures HT":
    st.markdown("## 🔄 Import Factures - Haute Tension (HT)")
    st.markdown("---")
    
    st.markdown("""
    <div style='background: linear-gradient(135deg, #f093fb, #f5576c); 
                color: white; 
                padding: 1.5rem; 
                border-radius: 10px;
                margin: 1rem 0;'>
        <h3 style='margin: 0;'>⚡ HAUTE TENSION</h3>
        <p style='margin: 0.5rem 0 0 0;'>Import mensuel - Ajout de nouvelles lignes</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    📌 **Configuration HT** :
    - Clé facture : **refraccord**
    - Clé base centrale : **IDENTIFIANT**
    - Données : **montfact**, **conso**, **caract** (période)
    
    💡 Pour chaque ligne trouvée, une **nouvelle ligne** sera ajoutée dans la base centrale avec les données du mois.
    """)
    
    # Upload fichier
    fichier_ht = st.file_uploader(
        "Sélectionnez le fichier de factures HT",
        type=['xlsx', 'xls'],
        key="upload_ht"
    )
    
    if fichier_ht:
        try:
            df_ht = pd.read_excel(fichier_ht)
            
            st.success(f"✅ Fichier chargé : {len(df_ht)} ligne(s)")
            
            # Configuration des colonnes
            cle_facture = "refraccord"
            montant_col = "montfact"
            conso_col = "conso"
            caract_col = "caract"
            
            # Vérifications
            colonnes_manquantes = []
            for col in [cle_facture, montant_col, caract_col]:
                if col not in df_ht.columns:
                    colonnes_manquantes.append(col)
            
            if colonnes_manquantes:
                st.error(f"❌ Colonnes manquantes : {', '.join(colonnes_manquantes)}")
                st.info(f"📋 Colonnes disponibles : {', '.join(df_ht.columns)}")
            else:
                # Récupérer la période
                periode_ht = df_ht[caract_col].dropna().unique()
                if len(periode_ht) > 0:
                    periode_ht = str(periode_ht[0])
                    st.success(f"✅ Période HT détectée : **{periode_ht}**")
                else:
                    periode_ht = ""
                    st.warning("⚠️ Aucune période détectée")
                
                # Aperçu
                with st.expander("👁️ Aperçu du fichier HT"):
                    cols_to_show = [cle_facture, montant_col, caract_col]
                    if conso_col in df_ht.columns:
                        cols_to_show.insert(2, conso_col)
                    st.dataframe(df_ht[cols_to_show].head(10), use_container_width=True)
                
                st.markdown("---")
                
                # Bouton import
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🔄 LANCER L'IMPORT HT", type="primary", use_container_width=True):
                        with st.spinner("⏳ Import HT en cours..."):
                            df_central = st.session_state.df_central.copy()
                            
                            nouvelles_lignes = []
                            nb_ajouts = 0
                            
                            # Pour chaque ligne de facture
                            for _, row_facture in df_ht.iterrows():
                                refraccord = str(row_facture[cle_facture])
                                
                                # Chercher dans la base centrale
                                ligne_centrale = df_central[df_central['IDENTIFIANT'].astype(str) == refraccord]
                                
                                if not ligne_centrale.empty:
                                    # Prendre la première occurrence
                                    ligne_base = ligne_centrale.iloc[0].copy()
                                    
                                    # Mettre à jour les valeurs
                                    ligne_base['MONTANT'] = row_facture[montant_col]
                                    ligne_base['DATE'] = periode_ht
                                    
                                    if conso_col in df_ht.columns:
                                        ligne_base['CONSO'] = row_facture.get(conso_col, None)
                                    
                                    nouvelles_lignes.append(ligne_base)
                                    nb_ajouts += 1
                            
                            if nouvelles_lignes:
                                # Créer un DataFrame avec les nouvelles lignes
                                df_nouvelles = pd.DataFrame(nouvelles_lignes)
                                
                                # Ajouter au DataFrame central
                                df_central = pd.concat([df_central, df_nouvelles], ignore_index=True)
                                
                                # Sauvegarder
                                st.session_state.df_central = df_central
                                save_central(df_central)
                                
                                # Résultats
                                st.markdown("---")
                                st.success(f"🎉 Import HT terminé : {nb_ajouts} ligne(s) ajoutée(s) !")
                                
                                col_r1, col_r2, col_r3 = st.columns(3)
                                with col_r1:
                                    st.metric("✅ Lignes ajoutées", nb_ajouts)
                                with col_r2:
                                    st.metric("📊 Total lignes", len(df_central))
                                with col_r3:
                                    st.metric("📅 Période", periode_ht)
                                
                                st.balloons()
                            else:
                                st.warning("⚠️ Aucune correspondance trouvée")
        
        except Exception as e:
            st.error(f"❌ Erreur : {str(e)}")
            st.exception(e)

elif page == "📈 Statistiques":
    st.markdown("## 📈 Statistiques et Évolution")
    st.markdown("---")
    
    df_central = st.session_state.df_central
    
    if 'DATE' not in df_central.columns or df_central['DATE'].isna().all():
        st.warning("⚠️ Aucune période enregistrée. Importez d'abord des factures.")
    else:
        # Récupérer les périodes
        periodes = sorted(df_central['DATE'].dropna().unique().tolist(), reverse=True)
        
        if len(periodes) < 2:
            st.warning("⚠️ Au moins 2 périodes nécessaires pour les comparaisons.")
        else:
            st.success(f"✅ {len(periodes)} période(s) disponible(s)")
            
            # Sélection de périodes
            col1, col2 = st.columns(2)
            with col1:
                periode1 = st.selectbox("Période 1", periodes, index=0)
            with col2:
                periode2 = st.selectbox("Période 2", periodes, index=min(1, len(periodes)-1))
            
            # Filtrer les données
            df_p1 = df_central[df_central['DATE'] == periode1]
            df_p2 = df_central[df_central['DATE'] == periode2]
            
            # Calculs
            total1 = df_p1['MONTANT'].sum() if 'MONTANT' in df_p1.columns else 0
            total2 = df_p2['MONTANT'].sum() if 'MONTANT' in df_p2.columns else 0
            evolution = ((total2 - total1) / total1 * 100) if total1 > 0 else 0
            
            conso1 = df_p1['CONSO'].sum() if 'CONSO' in df_p1.columns else 0
            conso2 = df_p2['CONSO'].sum() if 'CONSO' in df_p2.columns else 0
            evolution_conso = ((conso2 - conso1) / conso1 * 100) if conso1 > 0 else 0
            
            # Affichage
            st.markdown("### 💰 Comparaison des montants")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(f"💰 {periode1}", f"{total1:,.0f} FCFA")
            with col2:
                st.metric(f"💰 {periode2}", f"{total2:,.0f} FCFA")
            with col3:
                st.metric("📈 Évolution", f"{evolution:+.1f}%", delta=f"{total2-total1:,.0f} FCFA")
            
            st.markdown("### ⚡ Comparaison des consommations")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(f"⚡ {periode1}", f"{conso1:,.0f} kWh")
            with col2:
                st.metric(f"⚡ {periode2}", f"{conso2:,.0f} kWh")
            with col3:
                st.metric("📈 Évolution", f"{evolution_conso:+.1f}%", delta=f"{conso2-conso1:,.0f} kWh")
            
            # Tableau comparatif par UC
            st.markdown("### 📊 Évolution par Unité Commerciale")
            
            if 'UC' in df_central.columns:
                # Grouper par UC
                group_p1 = df_p1.groupby('UC').agg({'MONTANT': 'sum', 'CONSO': 'sum'}).reset_index()
                group_p2 = df_p2.groupby('UC').agg({'MONTANT': 'sum', 'CONSO': 'sum'}).reset_index()
                
                # Fusionner
                df_compare = group_p1.merge(group_p2, on='UC', suffixes=(f'_{periode1}', f'_{periode2}'), how='outer').fillna(0)
                
                # Calculer évolution
                df_compare['EVOL_MONTANT'] = df_compare[f'MONTANT_{periode2}'] - df_compare[f'MONTANT_{periode1}']
                df_compare['EVOL_MONTANT_%'] = ((df_compare[f'MONTANT_{periode2}'] - df_compare[f'MONTANT_{periode1}']) / df_compare[f'MONTANT_{periode1}'] * 100).fillna(0)
                df_compare['EVOL_CONSO'] = df_compare[f'CONSO_{periode2}'] - df_compare[f'CONSO_{periode1}']
                df_compare['EVOL_CONSO_%'] = ((df_compare[f'CONSO_{periode2}'] - df_compare[f'CONSO_{periode1}']) / df_compare[f'CONSO_{periode1}'] * 100).fillna(0)
                
                st.dataframe(df_compare, use_container_width=True, height=400)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; padding: 2rem; color: #666;'>
    <p><strong>Système Centralisé</strong> - Version 3.1 - Historique par Ligne</p>
</div>
""", unsafe_allow_html=True)
