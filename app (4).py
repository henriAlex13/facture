import streamlit as st
import pandas as pd
import io
import pickle
import os
from datetime import datetime

st.set_page_config(
    page_title="Gestion Factures - SOCODECI & CIE",
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
    .badge-coming {
        background: #e0e0e0;
        color: #666;
        padding: 0.3rem 1rem;
        border-radius: 20px;
        font-style: italic;
        display: inline-block;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# Fichiers Excel principaux
FICHIER_BT = "FACTURAT_ELECTRICITE_BT.xlsx"
FICHIER_HT = "FACTURAT_ELECTRICITE_HT.xlsx"

# Fichiers de sauvegarde (backup automatique)
SAVE_FILE_BT = "data_factures_cie_bt.pkl"
SAVE_FILE_HT = "data_factures_cie_ht.pkl"

# Fonctions de chargement
def load_data_bt():
    # Essayer d'abord le pickle (sauvegarde)
    if os.path.exists(SAVE_FILE_BT):
        with open(SAVE_FILE_BT, 'rb') as f:
            return pickle.load(f)
    # Sinon charger depuis Excel
    elif os.path.exists(FICHIER_BT):
        return pd.read_excel(FICHIER_BT)
    else:
        st.error(f"""
        ❌ **Fichier BT introuvable !**
        
        Veuillez placer le fichier `{FICHIER_BT}` dans le même dossier que l'application.
        """)
        st.stop()

def load_data_ht():
    # Essayer d'abord le pickle (sauvegarde)
    if os.path.exists(SAVE_FILE_HT):
        with open(SAVE_FILE_HT, 'rb') as f:
            return pickle.load(f)
    # Sinon charger depuis Excel
    elif os.path.exists(FICHIER_HT):
        return pd.read_excel(FICHIER_HT)
    else:
        st.error(f"""
        ❌ **Fichier HT introuvable !**
        
        Veuillez placer le fichier `{FICHIER_HT}` dans le même dossier que l'application.
        """)
        st.stop()

def save_data(df, type_tension):
    file = SAVE_FILE_HT if type_tension == 'HT' else SAVE_FILE_BT
    with open(file, 'wb') as f:
        pickle.dump(df, f)

# Initialisation
if 'df_bt' not in st.session_state:
    st.session_state.df_bt = load_data_bt()
if 'df_ht' not in st.session_state:
    st.session_state.df_ht = load_data_ht()

# Header
st.markdown("""
<div class="main-header">
    <h1 class="main-title">⚡ Gestion des Factures Électriques</h1>
    <p class="main-subtitle">SOCODECI & CIE - Plateforme de gestion</p>
</div>
""", unsafe_allow_html=True)

# Sidebar - Sélection du type uniquement
with st.sidebar:
    st.markdown("### 📋 Navigation Principale")
    st.markdown("#### ⚡ CIE")
    
    type_tension = st.radio(
        "Type de tension",
        ["🔌 Basse Tension (BT)", "⚡ Haute Tension (HT)"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    # Affichage selon le type
    if "BT" in type_tension:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 1.5rem; 
                    border-radius: 10px;'>
            <h3 style='color: white; margin: 0; text-align: center;'>🔌 BASSE TENSION</h3>
            <p style='color: white; margin: 0.5rem 0 0 0; text-align: center;'>Module actif</p>
        </div>
        """, unsafe_allow_html=True)
        
        df_actuel = st.session_state.df_bt
        type_code = "BT"
        
        st.markdown("---")
        st.markdown("### 📊 Statistiques BT")
        
        try:
            montants = pd.to_numeric(df_actuel['MONTANT'], errors='coerce').fillna(0)
            total = montants.sum()
            moyenne = montants.mean()
        except:
            total = 0
            moyenne = 0
        
        st.metric("📝 Lignes", len(df_actuel))
        st.metric("💰 Total", f"{total:,.0f} FCFA")
        st.metric("📊 Moyenne", f"{moyenne:,.0f} FCFA")
    
    else:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); 
                    padding: 1.5rem; 
                    border-radius: 10px;'>
            <h3 style='color: white; margin: 0; text-align: center;'>⚡ HAUTE TENSION</h3>
            <p style='color: white; margin: 0.5rem 0 0 0; text-align: center;'>Module actif</p>
        </div>
        """, unsafe_allow_html=True)
        
        df_actuel = st.session_state.df_ht
        type_code = "HT"
        
        st.markdown("---")
        st.markdown("### 📊 Statistiques HT")
        
        try:
            montants = pd.to_numeric(df_actuel['MONTANT'], errors='coerce').fillna(0)
            total = montants.sum()
            moyenne = montants.mean()
        except:
            total = 0
            moyenne = 0
        
        st.metric("📝 Lignes", len(df_actuel))
        st.metric("💰 Total", f"{total:,.0f} FCFA")
        st.metric("📊 Moyenne", f"{moyenne:,.0f} FCFA")
    
    st.markdown("---")
    st.markdown("#### 💧 SOCODECI")
    st.markdown('<span class="badge-coming">À venir</span>', unsafe_allow_html=True)

# Contenu principal - ONGLETS HORIZONTAUX
if "BT" in type_tension:
    # ===== VUE BASSE TENSION =====
    tab1, tab2 = st.tabs(["📊 Tableau de bord", "🔄 Mise à jour montants"])
    
    with tab1:
        # TABLEAU DE BORD BT
        st.markdown("## 📊 Tableau de bord - Basse Tension (BT)")
        st.markdown("---")
        
        # Statistiques en cartes
        col1, col2, col3, col4 = st.columns(4)
        
        try:
            montants = pd.to_numeric(df_actuel['MONTANT'], errors='coerce').fillna(0)
            total = montants.sum()
            moyenne = montants.mean()
        except:
            total = 0
            moyenne = 0
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #667eea; margin: 0;">📝</h3>
                <h2 style="margin: 0.5rem 0;">{len(df_actuel)}</h2>
                <p style="color: #666; margin: 0;">Total lignes</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #667eea; margin: 0;">💰</h3>
                <h2 style="margin: 0.5rem 0;">{total:,.0f}</h2>
                <p style="color: #666; margin: 0;">Total FCFA</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #667eea; margin: 0;">📊</h3>
                <h2 style="margin: 0.5rem 0;">{moyenne:,.0f}</h2>
                <p style="color: #666; margin: 0;">Moyenne FCFA</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            comptes = df_actuel['COMPTE DE CHARGES'].nunique()
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #667eea; margin: 0;">🔢</h3>
                <h2 style="margin: 0.5rem 0;">{comptes}</h2>
                <p style="color: #666; margin: 0;">Comptes uniques</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Tableau
        st.markdown("### 📋 Données CIE BT")
        st.markdown("*Ajoutez, modifiez ou supprimez des lignes directement*")
        
        edited_df = st.data_editor(
            df_actuel,
            use_container_width=True,
            num_rows="dynamic",
            height=400,
            key="editor_bt"
        )
        
        # Actions
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("💾 Sauvegarder", type="primary", use_container_width=True, key="save_bt"):
                st.session_state.df_bt = edited_df
                save_data(edited_df, "BT")
                st.success("✅ Sauvegardé!")
                st.rerun()
        
        with col2:
            if st.button("📥 Télécharger Excel BT", use_container_width=True, key="dl_bt"):
                st.info("Génération Excel BT...")
        
        with col3:
            if st.button("🔄 Actualiser", use_container_width=True, key="refresh_bt"):
                st.rerun()
    
    with tab2:
        # MISE À JOUR BT
        st.markdown("## 🔄 Mise à jour des montants - Basse Tension (BT)")
        st.markdown("---")
        
        st.info("""
        📌 **Configuration automatique BT** :
        - ✅ Colonne clé : **reference contrat**
        - ✅ Colonne montant : **Montant facture TTC**
        
        ℹ️ *Note: Le fichier source contient "Montant" et "Montant facture TTC", nous utilisons "Montant facture TTC"*
        """)
        
        fichier_source = st.file_uploader(
            "Sélectionnez le fichier source CIE BT",
            type=['xlsx', 'xls'],
            key="upload_bt"
        )
        
        if fichier_source:
            try:
                df_source = pd.read_excel(fichier_source)
                
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #667eea, #764ba2); 
                            color: white; 
                            padding: 1rem; 
                            border-radius: 8px; 
                            margin: 1rem 0;'>
                    <strong>✅ Fichier chargé avec succès</strong><br>
                    📊 {len(df_source)} ligne(s) · 📋 {len(df_source.columns)} colonne(s)
                </div>
                """, unsafe_allow_html=True)
                
                # Configuration des colonnes pour BT
                cle_source = "reference contrat"
                montant_source = "Montant facture TTC"  # On prend "Montant facture TTC" et pas "Montant"
                cle_principal = "COMPTE DE CHARGES"
                montant_principal = "MONTANT"
                
                if cle_source not in df_source.columns:
                    st.error(f"❌ Colonne '{cle_source}' introuvable dans le fichier source !")
                    st.info(f"📋 Colonnes disponibles : {', '.join(df_source.columns)}")
                elif montant_source not in df_source.columns:
                    st.error(f"❌ Colonne '{montant_source}' introuvable dans le fichier source !")
                    st.info(f"📋 Colonnes disponibles : {', '.join(df_source.columns)}")
                else:
                    # Aperçu
                    with st.expander("👁️ Aperçu du fichier source"):
                        colonnes_a_afficher = [cle_source, montant_source]
                        if "Montant" in df_source.columns:
                            colonnes_a_afficher.insert(1, "Montant")
                        st.dataframe(df_source[colonnes_a_afficher].head(10), use_container_width=True)
                        st.caption("💡 Nous utilisons 'Montant facture TTC' pour la mise à jour")
                    
                    st.markdown("---")
                    
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        if st.button("🔄 LANCER LA MISE À JOUR BT", type="primary", use_container_width=True, key="maj_bt"):
                            with st.spinner("⏳ Mise à jour BT en cours..."):
                                dict_montants = dict(zip(
                                    df_source[cle_source].astype(str), 
                                    df_source[montant_source]
                                ))
                                
                                nb_maj = 0
                                for idx, row in df_actuel.iterrows():
                                    cle = str(row[cle_principal])
                                    if cle in dict_montants:
                                        df_actuel.at[idx, montant_principal] = dict_montants[cle]
                                        nb_maj += 1
                                
                                st.session_state.df_bt = df_actuel
                                save_data(df_actuel, "BT")
                                
                                # Résultats
                                st.markdown("---")
                                st.markdown(f"""
                                <div style='background: linear-gradient(135deg, #667eea, #764ba2); 
                                            color: white; 
                                            padding: 2rem; 
                                            border-radius: 10px; 
                                            text-align: center;'>
                                    <h2 style='margin: 0;'>✅ MISE À JOUR BT TERMINÉE</h2>
                                    <p style='margin: 0.5rem 0 0 0;'>{datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                col_r1, col_r2, col_r3 = st.columns(3)
                                with col_r1:
                                    st.metric("✅ Mises à jour", nb_maj)
                                with col_r2:
                                    st.metric("⚠️ Non trouvées", len(df_actuel) - nb_maj)
                                with col_r3:
                                    taux = (nb_maj / len(df_actuel) * 100) if len(df_actuel) > 0 else 0
                                    st.metric("📈 Taux", f"{taux:.1f}%")
                                
                                st.success(f"🎉 {nb_maj} ligne(s) mise(s) à jour pour BT avec 'Montant facture TTC' !")
                                st.balloons()
            
            except Exception as e:
                st.error(f"❌ Erreur : {str(e)}")
                st.exception(e)

else:
    # ===== VUE HAUTE TENSION =====
    tab1, tab2 = st.tabs(["📊 Tableau de bord", "🔄 Mise à jour montants"])
    
    with tab1:
        # TABLEAU DE BORD HT
        st.markdown("## 📊 Tableau de bord - Haute Tension (HT)")
        st.markdown("---")
        
        # Statistiques en cartes
        col1, col2, col3, col4 = st.columns(4)
        
        try:
            montants = pd.to_numeric(df_actuel['MONTANT'], errors='coerce').fillna(0)
            total = montants.sum()
            moyenne = montants.mean()
        except:
            total = 0
            moyenne = 0
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #f5576c; margin: 0;">📝</h3>
                <h2 style="margin: 0.5rem 0;">{len(df_actuel)}</h2>
                <p style="color: #666; margin: 0;">Total lignes</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #f5576c; margin: 0;">💰</h3>
                <h2 style="margin: 0.5rem 0;">{total:,.0f}</h2>
                <p style="color: #666; margin: 0;">Total FCFA</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #f5576c; margin: 0;">📊</h3>
                <h2 style="margin: 0.5rem 0;">{moyenne:,.0f}</h2>
                <p style="color: #666; margin: 0;">Moyenne FCFA</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            comptes = df_actuel['COMPTE DE CHARGES'].nunique()
            st.markdown(f"""
            <div class="metric-card">
                <h3 style="color: #f5576c; margin: 0;">🔢</h3>
                <h2 style="margin: 0.5rem 0;">{comptes}</h2>
                <p style="color: #666; margin: 0;">Comptes uniques</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Tableau
        st.markdown("### 📋 Données CIE HT")
        st.markdown("*Ajoutez, modifiez ou supprimez des lignes directement*")
        
        edited_df = st.data_editor(
            df_actuel,
            use_container_width=True,
            num_rows="dynamic",
            height=400,
            key="editor_ht"
        )
        
        # Actions
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("💾 Sauvegarder", type="primary", use_container_width=True, key="save_ht"):
                st.session_state.df_ht = edited_df
                save_data(edited_df, "HT")
                st.success("✅ Sauvegardé!")
                st.rerun()
        
        with col2:
            if st.button("📥 Télécharger Excel HT", use_container_width=True, key="dl_ht"):
                st.info("Génération Excel HT...")
        
        with col3:
            if st.button("🔄 Actualiser", use_container_width=True, key="refresh_ht"):
                st.rerun()
    
    with tab2:
        # MISE À JOUR HT
        st.markdown("## 🔄 Mise à jour des montants - Haute Tension (HT)")
        st.markdown("---")
        
        st.info("""
        📌 **Configuration automatique HT** :
        - ✅ Colonne clé : **refraccord**
        - ✅ Colonne montant : **montfact**
        
        ℹ️ *Note: Le fichier source contient "Montant" et "montfact", nous utilisons "montfact"*
        """)
        
        fichier_source = st.file_uploader(
            "Sélectionnez le fichier source CIE HT",
            type=['xlsx', 'xls'],
            key="upload_ht"
        )
        
        if fichier_source:
            try:
                df_source = pd.read_excel(fichier_source)
                
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #f093fb, #f5576c); 
                            color: white; 
                            padding: 1rem; 
                            border-radius: 8px; 
                            margin: 1rem 0;'>
                    <strong>✅ Fichier chargé avec succès</strong><br>
                    📊 {len(df_source)} ligne(s) · 📋 {len(df_source.columns)} colonne(s)
                </div>
                """, unsafe_allow_html=True)
                
                # Configuration des colonnes pour HT
                cle_source = "refraccord"
                montant_source = "montfact"  # On prend "montfact" et pas "Montant"
                cle_principal = "COMPTE DE CHARGES"
                montant_principal = "MONTANT"
                
                if cle_source not in df_source.columns:
                    st.error(f"❌ Colonne '{cle_source}' introuvable dans le fichier source !")
                    st.info(f"📋 Colonnes disponibles : {', '.join(df_source.columns)}")
                elif montant_source not in df_source.columns:
                    st.error(f"❌ Colonne '{montant_source}' introuvable dans le fichier source !")
                    st.info(f"📋 Colonnes disponibles : {', '.join(df_source.columns)}")
                else:
                    # Aperçu
                    with st.expander("👁️ Aperçu du fichier source"):
                        colonnes_a_afficher = [cle_source, montant_source]
                        if "Montant" in df_source.columns:
                            colonnes_a_afficher.insert(1, "Montant")
                        st.dataframe(df_source[colonnes_a_afficher].head(10), use_container_width=True)
                        st.caption("💡 Nous utilisons 'montfact' pour la mise à jour")
                    
                    st.markdown("---")
                    
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        if st.button("🔄 LANCER LA MISE À JOUR HT", type="primary", use_container_width=True, key="maj_ht"):
                            with st.spinner("⏳ Mise à jour HT en cours..."):
                                dict_montants = dict(zip(
                                    df_source[cle_source].astype(str), 
                                    df_source[montant_source]
                                ))
                                
                                nb_maj = 0
                                for idx, row in df_actuel.iterrows():
                                    cle = str(row[cle_principal])
                                    if cle in dict_montants:
                                        df_actuel.at[idx, montant_principal] = dict_montants[cle]
                                        nb_maj += 1
                                
                                st.session_state.df_ht = df_actuel
                                save_data(df_actuel, "HT")
                                
                                # Résultats
                                st.markdown("---")
                                st.markdown(f"""
                                <div style='background: linear-gradient(135deg, #f093fb, #f5576c); 
                                            color: white; 
                                            padding: 2rem; 
                                            border-radius: 10px; 
                                            text-align: center;'>
                                    <h2 style='margin: 0;'>✅ MISE À JOUR HT TERMINÉE</h2>
                                    <p style='margin: 0.5rem 0 0 0;'>{datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                col_r1, col_r2, col_r3 = st.columns(3)
                                with col_r1:
                                    st.metric("✅ Mises à jour", nb_maj)
                                with col_r2:
                                    st.metric("⚠️ Non trouvées", len(df_actuel) - nb_maj)
                                with col_r3:
                                    taux = (nb_maj / len(df_actuel) * 100) if len(df_actuel) > 0 else 0
                                    st.metric("📈 Taux", f"{taux:.1f}%")
                                
                                st.success(f"🎉 {nb_maj} ligne(s) mise(s) à jour pour HT avec 'montfact' !")
                                st.balloons()
            
            except Exception as e:
                st.error(f"❌ Erreur : {str(e)}")
                st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; padding: 2rem; color: #666;'>
    <p><strong>SOCODECI & CIE</strong> - Version 2.0 - BT & HT Actifs</p>
</div>
""", unsafe_allow_html=True)
