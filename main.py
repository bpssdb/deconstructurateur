"""
Déconstructurateur Excel - Application principale
Point d'entrée de l'application Streamlit
"""

import streamlit as st
from datetime import datetime
import json

# Import des modules locaux
from utils.excel_utils import load_workbook, get_sheet_names
from utils.color_detector import detect_all_colors
from utils.zone_detector import detect_zones_with_palette, detect_zones_with_flexible_palette
from utils.visualization import create_excel_visualization, create_color_preview_html, create_zone_detail_view, create_dataframe_view
from utils.export import export_to_json
import plotly.express as px
import pandas as pd
from collections import defaultdict

# Configuration de la page Streamlit
st.set_page_config(
    page_title="📊 Déconstructurateur Excel",
    page_icon="📊",
    layout="wide"
)

# CSS personnalisé
st.markdown("""
<style>
    .stDataFrame {
        font-size: 12px;
    }
    .color-preview {
        display: inline-block;
        width: 30px;
        height: 30px;
        border: 2px solid #333;
        border-radius: 4px;
        margin-right: 10px;
        vertical-align: middle;
    }
    div[data-testid="stHorizontalBlock"] {
        align-items: stretch;
    }
</style>
""", unsafe_allow_html=True)

def init_session_state():
    """Initialise les variables de session"""
    if 'zones' not in st.session_state:
        st.session_state.zones = []
    if 'current_sheet' not in st.session_state:
        st.session_state.current_sheet = None
    if 'workbook' not in st.session_state:
        st.session_state.workbook = None
    if 'selected_zone' not in st.session_state:
        st.session_state.selected_zone = None
    if 'color_palette' not in st.session_state:
        st.session_state.color_palette = None
    if 'detected_colors' not in st.session_state:
        st.session_state.detected_colors = []
    if 'color_cells' not in st.session_state:
        st.session_state.color_cells = {}

def main():
    """Fonction principale de l'application"""
    init_session_state()
    
    st.title("📊 Déconstructurateur Excel Python")
    st.markdown("Détection automatique des couleurs et configuration de la palette")
    
    # Upload du fichier
    uploaded_file = st.file_uploader(
        "📂 Charger un fichier Excel (.xlsx, .xls)", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file:
        try:
            # Charger le workbook
            st.session_state.workbook = load_workbook(uploaded_file)
            sheet_names = get_sheet_names(st.session_state.workbook)
            
            # Sélection de la feuille
            selected_sheet = st.selectbox("📄 Sélectionner une feuille", sheet_names)
            
            # Étape 1: Détection des couleurs
            st.header("🎨 Étape 1: Détection des couleurs")
            
            col1, col2 = st.columns([1, 3])
            
            with col1:
                if st.button("🔍 Analyser les couleurs", type="primary"):
                    with st.spinner("Analyse des couleurs en cours..."):
                        colors, color_cells = detect_all_colors(
                            st.session_state.workbook, 
                            selected_sheet
                        )
                        st.session_state.detected_colors = colors
                        st.session_state.color_cells = color_cells
                        st.session_state.current_sheet = selected_sheet
                        
                        if len(colors) > 0:
                            st.success(f"✅ {len(colors)} couleurs détectées!")
                        else:
                            st.warning("⚠️ Aucune couleur détectée. Vérifiez que votre fichier contient des cellules colorées.")
                            
                            # Afficher des informations de debug
                            with st.expander("🔧 Informations de débogage"):
                                st.write("Essayez de:")
                                st.write("- Vérifier que les cellules ont bien une couleur de fond (pas juste du texte coloré)")
                                st.write("- Resauvegarder le fichier dans Excel")
                                st.write("- Utiliser un format .xlsx plutôt que .xls")
                                st.write("- S'assurer que les couleurs ne sont pas blanches (#FFFFFF)")
                                
                                # Afficher un échantillon de cellules pour debug
                                ws = st.session_state.workbook[selected_sheet]
                                st.write(f"Dimensions de la feuille: {ws.max_row} lignes x {ws.max_column} colonnes")
            
            # Afficher les couleurs détectées
            if st.session_state.detected_colors:
                display_detected_colors()
                
                # Étape 2: Configuration de la palette
                configure_color_palette(selected_sheet)
            
            # Affichage des résultats
            if st.session_state.zones and st.session_state.color_palette:
                display_results(selected_sheet)
                
        except Exception as e:
            st.error(f"❌ Erreur lors du chargement du fichier: {str(e)}")
            st.info("Assurez-vous que le fichier n'est pas corrompu et qu'il s'agit bien d'un fichier Excel.")
    
    # Instructions
    display_instructions()

def display_detected_colors():
    """Affiche les couleurs détectées avec une visualisation améliorée"""
    st.subheader("Couleurs trouvées dans la feuille:")
    
    if not st.session_state.detected_colors:
        st.warning("Aucune couleur détectée dans la feuille.")
        return
    
    # Créer deux colonnes pour l'affichage
    col1, col2 = st.columns([2, 3])
    
    with col1:
        # Tableau des couleurs
        st.markdown("### 🎨 Palette détectée")
        html_table = create_color_preview_html(st.session_state.detected_colors)
        st.markdown(html_table, unsafe_allow_html=True)
    
    with col2:
        # Visualisation de la distribution des couleurs
        st.markdown("### 📊 Distribution des couleurs")
        
        # Créer un graphique en barres des couleurs
        import plotly.express as px
        
        if st.session_state.detected_colors:
            color_data = []
            color_map = {}
            
            for color in st.session_state.detected_colors[:10]:  # Limiter aux 10 premières
                color_name = f"{color['name']} (#{color['hex']})"
                color_data.append({
                    'Couleur': color_name,
                    'Occurrences': color['count']
                })
                # Créer un mapping pour les couleurs réelles
                color_map[color_name] = f"#{color['hex']}"
            
            if color_data:
                df_colors = pd.DataFrame(color_data)
                
                fig = px.bar(
                    df_colors, 
                    x='Couleur', 
                    y='Occurrences',
                    title="Nombre de cellules par couleur"
                )
                
                # Appliquer les couleurs réelles aux barres
                colors_list = [color_map.get(name, '#888888') for name in df_colors['Couleur']]
                fig.update_traces(marker_color=colors_list)
                
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Aucune donnée de couleur à afficher")
        else:
            st.info("Aucune couleur détectée pour créer le graphique")

def configure_color_palette(selected_sheet):
    """Configure la palette de couleurs pour la détection"""
    st.header("🎯 Étape 2: Configuration de la palette")
    st.info("Identifiez les 3 types de couleurs dans votre fichier Excel")
    
    # Préparer les options de couleurs
    color_options = {
        f"{c['name']} (#{c['hex']})": c['hex'] 
        for c in st.session_state.detected_colors
    }
    
    # Configuration pour exactement 3 couleurs
    st.markdown("### 📦 1. Couleur des zones de données")
    zone_color = st.selectbox(
        "Cellules à labelliser (données à compléter par le LLM)",
        options=list(color_options.keys()),
        help="Sélectionnez la couleur des cellules qui contiennent les données à traiter"
    )
    
    st.markdown("### 🏷️ 2. Couleurs des labels (en-têtes)")
    col1, col2 = st.columns(2)
    
    with col1:
        label_h_color = st.selectbox(
            "Labels horizontaux (en-têtes de colonnes)",
            options=list(color_options.keys()),
            help="Couleur des cellules qui servent d'en-têtes en haut des colonnes"
        )
    
    with col2:
        label_v_color = st.selectbox(
            "Labels verticaux (en-têtes de lignes)",
            options=list(color_options.keys()),
            help="Couleur des cellules qui servent d'en-têtes à gauche des lignes"
        )
    
    # Afficher un aperçu de la logique
    with st.expander("💡 Comment ça marche ?"):
        st.markdown("""
        **Structure attendue du fichier Excel :**
        
        1. **Zones de données** : Les cellules colorées qui contiennent les valeurs à traiter
        2. **Labels horizontaux** : Les en-têtes situés AU-DESSUS des zones (peuvent être fusionnés sur plusieurs colonnes)
        3. **Labels verticaux** : Les en-têtes situés À GAUCHE des zones (peuvent être fusionnés sur plusieurs lignes)
        
        **L'application va :**
        - Détecter toutes les zones contiguës de la couleur "données"
        - Pour chaque zone, chercher les labels immédiatement adjacents (au-dessus et à gauche)
        - Gérer automatiquement les cellules fusionnées
        - Créer un mapping structuré pour le LLM
        """)
    
    # Bouton de validation
    if st.button("✅ Valider et détecter les zones", type="primary"):
        # Vérifier que les 3 couleurs sont différentes
        selected_colors = [color_options[zone_color], color_options[label_h_color], color_options[label_v_color]]
        
        if len(set(selected_colors)) != 3:
            st.error("❌ Veuillez sélectionner 3 couleurs différentes !")
        else:
            # Créer la palette dans le format attendu
            label_colors = {
                'horizontal': {
                    'color': color_options[label_h_color],
                    'name': f"Labels horizontaux ({label_h_color.split(' (')[0]})"
                },
                'vertical': {
                    'color': color_options[label_v_color],
                    'name': f"Labels verticaux ({label_v_color.split(' (')[0]})"
                }
            }
            
            validate_and_detect_zones_flexible(
                selected_sheet, 
                color_options[zone_color],
                zone_color.split(' (')[0],
                label_colors
            )
    
    # Afficher la palette sélectionnée
    if st.session_state.color_palette:
        display_selected_palette()

def validate_and_detect_zones_flexible(selected_sheet, zone_color, zone_name, label_colors):
    """Valide la palette et lance la détection des zones avec support multi-labels"""
    st.session_state.color_palette = {
        'zone_color': zone_color,
        'zone_name': zone_name,
        'label_colors': label_colors  # Dictionnaire des couleurs de labels
    }
    
    # Détecter les zones
    with st.spinner("Détection des zones en cours..."):
        from utils.zone_detector import detect_zones_with_flexible_palette
        zones, all_labels = detect_zones_with_flexible_palette(
            st.session_state.workbook,
            selected_sheet,
            st.session_state.color_palette,
            st.session_state.color_cells
        )
        st.session_state.zones = zones
        st.session_state.all_labels = all_labels
        st.success(f"✅ {len(zones)} zones détectées!")

def display_selected_palette():
    """Affiche la palette de couleurs sélectionnée"""
    st.subheader("Palette sélectionnée:")
    
    # Zone de données
    st.markdown(f"""
    <div style="display: flex; align-items: center; margin: 10px 0;">
        <div class="color-preview" style="background-color: #{st.session_state.color_palette['zone_color']}; margin-right: 10px;"></div>
        <strong>Zones de données:</strong> {st.session_state.color_palette['zone_name']}
    </div>
    """, unsafe_allow_html=True)
    
    # Labels
    if st.session_state.color_palette.get('label_colors'):
        for label_type, label_info in st.session_state.color_palette['label_colors'].items():
            st.markdown(f"""
            <div style="display: flex; align-items: center; margin: 10px 0;">
                <div class="color-preview" style="background-color: #{label_info['color']}; margin-right: 10px;"></div>
                <strong>{label_type}:</strong> {label_info['name']}
            </div>
            """, unsafe_allow_html=True)

def display_results(selected_sheet):
    """Affiche les résultats de la détection avec contrôles améliorés"""
    st.header("📊 Visualisation et édition")
    
    # Barre d'outils
    tool_col1, tool_col2, tool_col3, tool_col4 = st.columns([1, 1, 1, 1])
    
    with tool_col1:
        if st.button("🔄 Rafraîchir la vue"):
            st.rerun()
    
    with tool_col2:
        if st.button("🔀 Fusionner zones proches"):
            from utils.zone_detector import merge_zones
            st.session_state.zones = merge_zones(st.session_state.zones, max_gap=1)
            st.success("Zones fusionnées!")
            st.rerun()
    
    with tool_col3:
        if st.button("➕ Nouvelle zone manuelle"):
            st.session_state.show_manual_zone = True
    
    with tool_col4:
        if st.button("📥 Télécharger JSON"):
            json_data = export_to_json(
                st.session_state.zones,
                st.session_state.current_sheet,
                st.session_state.color_palette
            )
            st.download_button(
                label="💾 Télécharger",
                data=json_data,
                file_name=f"zones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )
    
    # Affichage principal avec tabs
    tab1, tab2, tab3 = st.tabs(["📋 Vue d'ensemble", "🔍 Vue détaillée", "📊 Statistiques"])
    
    with tab1:
        display_overview_tab(selected_sheet)
    
    with tab2:
        display_detailed_tab(selected_sheet)
    
    with tab3:
        display_statistics_tab()
    
    # Modal pour l'ajout manuel de zone
    if hasattr(st.session_state, 'show_manual_zone') and st.session_state.show_manual_zone:
        display_manual_zone_modal()

def display_overview_tab(selected_sheet):
    """Affiche l'onglet vue d'ensemble"""
    # Sous-tabs pour différentes vues
    view_tab1, view_tab2 = st.tabs(["🗺️ Vue schématique", "📋 Vue tableau"])
    
    with view_tab1:
        col1, col2 = st.columns([3, 1])
        
        with col1:
            # Vue Plotly interactive principale
            fig = create_excel_visualization(
                st.session_state.workbook,
                selected_sheet,
                st.session_state.zones,
                st.session_state.selected_zone,
                st.session_state.color_palette
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Légende interactive
            st.markdown("### 🎯 Légende")
            leg_cols = st.columns(min(len(st.session_state.color_palette.get('label_colors', {})) + 1, 4))
            
            # Zone de données
            with leg_cols[0]:
                st.markdown(f"""
                <div style="display: flex; align-items: center;">
                    <div style="width: 20px; height: 20px; background-color: #{st.session_state.color_palette['zone_color']}; border: 1px solid black; margin-right: 10px;"></div>
                    <span>Zones de données</span>
                </div>
                """, unsafe_allow_html=True)
            
            # Labels
            if 'label_colors' in st.session_state.color_palette:
                for i, (label_type, label_info) in enumerate(st.session_state.color_palette['label_colors'].items(), 1):
                    if i < len(leg_cols):
                        with leg_cols[i]:
                            st.markdown(f"""
                            <div style="display: flex; align-items: center;">
                                <div style="width: 20px; height: 20px; background-color: #{label_info['color']}; border: 1px solid black; margin-right: 10px;"></div>
                                <span>{label_info['name']}</span>
                            </div>
                            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("### 🎮 Contrôles rapides")
            
            # Liste des zones avec actions rapides
            for zone in st.session_state.zones:
                with st.container():
                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        if st.button(f"Zone {zone['id']} ({zone['cell_count']} cellules)", 
                                    key=f"select_zone_{zone['id']}",
                                    use_container_width=True):
                            st.session_state.selected_zone = zone['id']
                            st.rerun()
                    with col_b:
                        if st.button("❌", key=f"delete_zone_{zone['id']}", help="Supprimer"):
                            st.session_state.zones = [z for z in st.session_state.zones if z['id'] != zone['id']]
                            st.rerun()
    
    with view_tab2:
        st.markdown("### 📊 Vue tableau avec contenu des cellules")
        
        # Options d'affichage
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            show_colors = st.checkbox("Afficher les couleurs", value=True)
        with col2:
            max_rows = st.number_input("Nombre de lignes max", min_value=10, max_value=200, value=50)
        
        # Créer la vue tableau
        from utils.visualization import create_dataframe_view
        try:
            df_styled = create_dataframe_view(
                st.session_state.workbook,
                selected_sheet,
                st.session_state.zones if show_colors else None,
                st.session_state.color_palette if show_colors else None,
                max_rows=max_rows
            )
            
            if show_colors and hasattr(df_styled, 'to_html'):
                # Afficher avec style
                st.markdown(df_styled.to_html(), unsafe_allow_html=True)
            else:
                # Afficher sans style
                st.dataframe(df_styled, use_container_width=True, height=600)
                
        except Exception as e:
            st.error(f"Erreur lors de la création de la vue tableau: {str(e)}")
            st.info("Essayez de réduire le nombre de lignes à afficher.")

def display_detailed_tab(selected_sheet):
    """Affiche l'onglet vue détaillée avec zoom, tableau et analyse comparative"""
    if not st.session_state.selected_zone:
        st.info("👆 Sélectionnez une zone dans l'onglet 'Vue d'ensemble' pour voir les détails")
        return
    
    zone = next((z for z in st.session_state.zones if z['id'] == st.session_state.selected_zone), None)
    if not zone:
        return
    
    # En-tête avec navigation
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"### Zone {zone['id']}")
    with col2:
        if st.button("⬅️ Zone précédente", disabled=zone['id'] == 1):
            st.session_state.selected_zone = max(1, zone['id'] - 1)
            st.rerun()
    with col3:
        if st.button("Zone suivante ➡️", disabled=zone['id'] == len(st.session_state.zones)):
            st.session_state.selected_zone = min(len(st.session_state.zones), zone['id'] + 1)
            st.rerun()
    
    # NOUVEAU: Tabs pour différentes vues de la zone
    detail_view_tab1, detail_view_tab2, detail_view_tab3 = st.tabs([
        "🗺️ Vue schématique", 
        "📋 Vue tableau", 
        "🔍 Analyse comparative"
    ])
    
    with detail_view_tab1:
        # Vue schématique existante (Plotly)
        st.markdown("#### 🔍 Vue zoomée de la zone")
        from utils.visualization import create_zone_detail_view
        zoom_fig = create_zone_detail_view(
            st.session_state.workbook,
            selected_sheet,
            zone,
            st.session_state.color_palette
        )
        st.plotly_chart(zoom_fig, use_container_width=True)
        
        # Note sur les problèmes d'affichage
        st.info("💡 **Problème d'affichage des zones ?** Essayez la vue tableau qui fonctionne parfaitement.")
    
    with detail_view_tab2:
        # NOUVELLE: Vue tableau détaillée
        st.markdown("#### 📋 Vue tableau de la zone")
        
        # Options d'affichage
        col1, col2, col3 = st.columns(3)
        with col1:
            show_markers = st.checkbox("Afficher les marqueurs visuels", value=True)
        with col2:
            table_style = st.selectbox("Style du tableau", ["Standard", "Avec marqueurs"], index=1)
        with col3:
            show_legend = st.checkbox("Afficher la légende", value=True)
        
        try:
            if table_style == "Avec marqueurs":
                from utils.visualization import create_zone_detail_table_view_enhanced
                styled_table = create_zone_detail_table_view_enhanced(
                    st.session_state.workbook,
                    selected_sheet,
                    zone,
                    st.session_state.color_palette
                )
            else:
                from utils.visualization import create_zone_detail_table_view
                styled_table = create_zone_detail_table_view(
                    st.session_state.workbook,
                    selected_sheet,
                    zone,
                    st.session_state.color_palette
                )
            
            # Afficher le tableau
            if hasattr(styled_table, 'to_html'):
                st.markdown(styled_table.to_html(escape=False), unsafe_allow_html=True)
            else:
                st.dataframe(styled_table, use_container_width=True)
            
            # Légende
            if show_legend:
                st.markdown("#### 🎨 Légende")
                leg_col1, leg_col2 = st.columns(2)
                
                with leg_col1:
                    zone_color = st.session_state.color_palette['zone_color']
                    st.markdown(f"""
                    <div style="display: flex; align-items: center; margin: 5px 0;">
                        <div style="width: 20px; height: 20px; background-color: #{zone_color}; opacity: 0.3; border: 3px solid #{zone_color}; margin-right: 10px;"></div>
                        <span>🔵 Cellules de zone</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                with leg_col2:
                    if 'label_colors' in st.session_state.color_palette:
                        for label_type, label_info in st.session_state.color_palette['label_colors'].items():
                            marker = "🏷️" if label_type == 'horizontal' else "📍"
                            color = label_info['color']
                            st.markdown(f"""
                            <div style="display: flex; align-items: center; margin: 5px 0;">
                                <div style="width: 20px; height: 20px; background-color: #{color}; opacity: 0.5; border: 2px solid #{color}; margin-right: 10px;"></div>
                                <span>{marker} {label_info['name']}</span>
                            </div>
                            """, unsafe_allow_html=True)
        
        except Exception as e:
            st.error(f"Erreur lors de la création de la vue tableau: {str(e)}")
            st.info("Essayez de réduire la taille de la zone ou vérifiez vos données.")
    
    with detail_view_tab3:
        # NOUVELLE: Analyse comparative
        st.markdown("#### 🔍 Analyse comparative détaillée")
        
        st.info("Cette analyse compare les données détectées avec la réalité du fichier Excel.")
        
        try:
            from utils.visualization import display_zone_comparison_table
            zone_df, label_df = display_zone_comparison_table(
                st.session_state.workbook,
                selected_sheet,
                zone,
                st.session_state.color_palette
            )
            
            if not zone_df.empty:
                st.markdown("##### 🔵 Analyse des cellules de zone")
                st.dataframe(zone_df, use_container_width=True)
                
                # Statistiques
                matches = zone_df['Correspondance'].str.count('✅').sum()
                total = len(zone_df)
                st.metric("Correspondances couleurs", f"{matches}/{total}", f"{matches/total*100:.1f}%" if total > 0 else "0%")
            
            if not label_df.empty:
                st.markdown("##### 🏷️ Analyse des labels")
                st.dataframe(label_df, use_container_width=True)
                
                # Statistiques
                matches = label_df['Correspondance'].str.count('✅').sum()
                total = len(label_df)
                st.metric("Correspondances labels", f"{matches}/{total}", f"{matches/total*100:.1f}%" if total > 0 else "0%")
            
            # Diagnostic
            st.markdown("##### 🎯 Diagnostic")
            
            if zone_df.empty and label_df.empty:
                st.warning("Aucune donnée à analyser pour cette zone")
            elif zone_df['Correspondance'].str.count('❌').sum() > 0:
                st.error("❌ Problèmes de correspondance couleurs détectés dans les cellules de zone")
                st.markdown("**Recommandations:**")
                st.write("- Vérifiez que les couleurs sélectionnées dans la palette correspondent exactement aux couleurs Excel")
                st.write("- Essayez de recalculer la détection des couleurs")
                st.write("- Vérifiez que les cellules ne sont pas fusionnées de manière inattendue")
            else:
                st.success("✅ Toutes les correspondances sont correctes")
        
        except Exception as e:
            st.error(f"Erreur lors de l'analyse comparative: {str(e)}")
    
    # EXISTANT: Tableau récapitulatif des labels (conservé pour compatibilité)
    st.markdown("#### 📊 Tableau des labels identifiés")
    if zone.get('labels'):
        from utils.excel_utils import num_to_excel_col
        # Créer un DataFrame pour les labels
        labels_data = []
        for label in zone['labels']:
            # Déterminer le nom du type de label
            label_type_name = label['type']
            if 'label_colors' in st.session_state.color_palette:
                for lt, linfo in st.session_state.color_palette['label_colors'].items():
                    if label['type'] == lt:
                        label_type_name = linfo['name']
                        break
            
            labels_data.append({
                'Position': f"{num_to_excel_col(label['col'])}{label['row']}",
                'Valeur': label['value'],
                'Type': label_type_name,
                'Direction': 'Colonne' if label['position'] == 'top' else 'Ligne',
                'Distance': label.get('distance', 1),
                'Appliqué à': len(label.get('for_cells', []))
            })
        
        labels_df = pd.DataFrame(labels_data)
        st.dataframe(labels_df, use_container_width=True)
        
        # Statistiques des labels
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total labels", len(labels_data))
        with col2:
            h_count = sum(1 for l in labels_data if l['Direction'] == 'Colonne')
            st.metric("Labels colonnes", h_count)
        with col3:
            v_count = sum(1 for l in labels_data if l['Direction'] == 'Ligne')
            st.metric("Labels lignes", v_count)
    else:
        st.warning("Aucun label identifié pour cette zone")
    
    # EXISTANT: Informations de la zone (conservé)
    st.markdown("#### 📍 Informations de la zone")
    info_col1, info_col2 = st.columns(2)
    
    with info_col1:
        from utils.excel_utils import num_to_excel_col
        st.write(f"**Lignes:** {zone['bounds']['min_row']} à {zone['bounds']['max_row']}")
        st.write(f"**Colonnes:** {num_to_excel_col(zone['bounds']['min_col'])} à {num_to_excel_col(zone['bounds']['max_col'])}")
        st.write(f"**Nombre de cellules:** {zone['cell_count']}")
    
    with info_col2:
        # Afficher un échantillon des valeurs de la zone
        st.write("**Échantillon de valeurs:**")
        sample_values = []
        for cell in zone['cells'][:5]:
            if cell.get('value'):
                sample_values.append(f"• {cell['value']}")
        if sample_values:
            st.write("\n".join(sample_values))
        else:
            st.write("(cellules vides)")
    
    # EXISTANT: Actions sur la zone (conservé)
    st.markdown("#### 🛠️ Actions")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        with st.expander("➕ Ajouter un label"):
            new_label_value = st.text_input("Valeur du label", key=f"new_label_{zone['id']}")
            
            # Types de labels disponibles
            label_types = list(st.session_state.color_palette.get('label_colors', {}).keys())
            if label_types:
                new_label_type = st.selectbox("Type", label_types, key=f"new_label_type_{zone['id']}")
            else:
                st.warning("Aucun type de label défini dans la palette")
                new_label_type = None
                
            new_label_pos = st.selectbox("Position", ["top", "left", "bottom", "right"], key=f"new_label_pos_{zone['id']}")
            
            if st.button("Ajouter", key=f"add_label_{zone['id']}") and new_label_type:
                if new_label_value:
                    # Calculer la position du nouveau label
                    if new_label_pos == "top":
                        row = zone['bounds']['min_row'] - 1
                        col = zone['bounds']['min_col']
                    elif new_label_pos == "left":
                        row = zone['bounds']['min_row']
                        col = zone['bounds']['min_col'] - 1
                    elif new_label_pos == "bottom":
                        row = zone['bounds']['max_row'] + 1
                        col = zone['bounds']['min_col']
                    else:  # right
                        row = zone['bounds']['min_row']
                        col = zone['bounds']['max_col'] + 1
                    
                    new_label = {
                        'row': row,
                        'col': col,
                        'value': new_label_value,
                        'type': new_label_type,
                        'position': new_label_pos,
                        'color': st.session_state.color_palette['label_colors'][new_label_type]['color']
                    }
                    
                    if 'labels' not in zone:
                        zone['labels'] = []
                    zone['labels'].append(new_label)
                    st.success("Label ajouté!")
                    st.rerun()
    
    with col2:
        with st.expander("📏 Modifier les limites"):
            new_min_row = st.number_input("Ligne min", value=zone['bounds']['min_row'], min_value=1, key=f"min_row_{zone['id']}")
            new_max_row = st.number_input("Ligne max", value=zone['bounds']['max_row'], min_value=1, key=f"max_row_{zone['id']}")
            new_min_col = st.number_input("Colonne min", value=zone['bounds']['min_col'], min_value=1, key=f"min_col_{zone['id']}")
            new_max_col = st.number_input("Colonne max", value=zone['bounds']['max_col'], min_value=1, key=f"max_col_{zone['id']}")
            
            if st.button("Appliquer", key=f"apply_bounds_{zone['id']}"):
                zone['bounds']['min_row'] = int(new_min_row)
                zone['bounds']['max_row'] = int(new_max_row)
                zone['bounds']['min_col'] = int(new_min_col)
                zone['bounds']['max_col'] = int(new_max_col)
                # Recalculer le nombre de cellules
                zone['cell_count'] = (new_max_row - new_min_row + 1) * (new_max_col - new_min_col + 1)
                st.success("Limites modifiées!")
                st.rerun()
    
    with col3:
        with st.expander("🔧 Autres actions"):
            if st.button("📋 Dupliquer la zone", key=f"duplicate_{zone['id']}"):
                new_zone = zone.copy()
                new_zone['id'] = max(z['id'] for z in st.session_state.zones) + 1
                st.session_state.zones.append(new_zone)
                st.success(f"Zone dupliquée (ID: {new_zone['id']})")
                st.rerun()
            
            if st.button("🗑️ Supprimer la zone", key=f"delete_detailed_{zone['id']}", type="secondary"):
                st.session_state.zones = [z for z in st.session_state.zones if z['id'] != zone['id']]
                st.session_state.selected_zone = None
                st.rerun()

def display_statistics_tab():
    """Affiche l'onglet statistiques"""
    if not st.session_state.zones:
        st.info("Aucune zone détectée pour afficher les statistiques")
        return
    
    # Statistiques globales
    col1, col2, col3, col4 = st.columns(4)
    
    total_zones = len(st.session_state.zones)
    total_cells = sum(z['cell_count'] for z in st.session_state.zones)
    total_labels = sum(len(z.get('labels', [])) for z in st.session_state.zones)
    avg_cells_per_zone = total_cells / total_zones if total_zones > 0 else 0
    
    col1.metric("📦 Zones", total_zones)
    col2.metric("📋 Cellules totales", total_cells)
    col3.metric("🏷️ Labels totaux", total_labels)
    col4.metric("📊 Moy. cellules/zone", f"{avg_cells_per_zone:.1f}")
    
    # Graphiques
    st.markdown("### 📊 Analyse détaillée")
    
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        # Distribution des tailles de zones
        zone_sizes = [z['cell_count'] for z in st.session_state.zones]
        df_sizes = pd.DataFrame({
            'Zone': [f"Zone {z['id']}" for z in st.session_state.zones],
            'Taille': zone_sizes
        })
        
        fig1 = px.bar(df_sizes, x='Zone', y='Taille', 
                      title="Taille des zones (nombre de cellules)")
        st.plotly_chart(fig1, use_container_width=True)
    
    with chart_col2:
        # Distribution des types de labels
        label_counts = defaultdict(int)
        label_names = {}
        
        # Compter les labels selon le format de palette
        if 'label_colors' in st.session_state.color_palette:
            # Format flexible
            for zone in st.session_state.zones:
                for label in zone.get('labels', []):
                    label_counts[label['type']] += 1
            
            # Noms des labels
            for label_type, label_info in st.session_state.color_palette['label_colors'].items():
                label_names[label_type] = label_info['name']
        else:
            # Ancien format (rétrocompatibilité)
            for zone in st.session_state.zones:
                for label in zone.get('labels', []):
                    label_counts[label['type']] += 1
            
            label_names = {
                'label1': st.session_state.color_palette.get('label1_name', 'Label 1'),
                'label2': st.session_state.color_palette.get('label2_name', 'Label 2')
            }
        
        # Créer le DataFrame pour le graphique
        if label_counts:
            df_labels = pd.DataFrame({
                'Type': [label_names.get(lt, lt) for lt in label_counts.keys()],
                'Nombre': list(label_counts.values())
            })
            
            fig2 = px.pie(df_labels, values='Nombre', names='Type',
                          title="Répartition des types de labels")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Aucun label détecté dans les zones")
    
    # Tableau récapitulatif
    st.markdown("### 📋 Tableau récapitulatif")
    from utils.visualization import create_zone_summary_dataframe
    summary_df = create_zone_summary_dataframe(st.session_state.zones)
    st.dataframe(summary_df, use_container_width=True)

def display_manual_zone_modal():
    """Affiche le modal pour créer une zone manuellement"""
    with st.container():
        st.markdown("### ➕ Créer une zone manuellement")
        
        col1, col2 = st.columns(2)
        with col1:
            man_min_row = st.number_input("Ligne début", min_value=1, value=1, key="manual_min_row")
            man_max_row = st.number_input("Ligne fin", min_value=1, value=1, key="manual_max_row")
        with col2:
            man_min_col = st.text_input("Colonne début (ex: A)", value="A", key="manual_min_col")
            man_max_col = st.text_input("Colonne fin (ex: B)", value="B", key="manual_max_col")
        
        col3, col4 = st.columns(2)
        with col3:
            if st.button("✅ Créer", type="primary"):
                try:
                    from utils.excel_utils import excel_col_to_num
                    min_col_num = excel_col_to_num(man_min_col)
                    max_col_num = excel_col_to_num(man_max_col)
                    
                    new_zone = {
                        'id': max([z['id'] for z in st.session_state.zones], default=0) + 1,
                        'cells': [],
                        'bounds': {
                            'min_row': int(man_min_row),
                            'max_row': int(man_max_row),
                            'min_col': min_col_num,
                            'max_col': max_col_num
                        },
                        'cell_count': (int(man_max_row) - int(man_min_row) + 1) * (max_col_num - min_col_num + 1),
                        'labels': []
                    }
                    
                    st.session_state.zones.append(new_zone)
                    st.session_state.show_manual_zone = False
                    st.success(f"Zone {new_zone['id']} créée!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur: {str(e)}")
        
        with col4:
            if st.button("❌ Annuler"):
                st.session_state.show_manual_zone = False
                st.rerun()

def display_instructions():
    """Affiche les instructions d'utilisation"""
    with st.expander("ℹ️ Guide d'utilisation"):
        st.markdown("""
        ## 🚀 Comment utiliser l'application
        
        ### 1. Analyse des couleurs
        - **Chargez votre fichier Excel** (.xlsx ou .xls)
        - **Sélectionnez la feuille** à analyser
        - **Cliquez sur "Analyser les couleurs"** pour détecter toutes les couleurs
        
        ### 2. Configuration de la palette
        - **Zones de données** : Cellules à remplir par le LLM
        - **Labels type 1** : Première couleur de labels
        - **Labels type 2** : Deuxième couleur de labels
        - **Validez la palette** pour lancer la détection
        
        ### 3. Visualisation et édition
        - Les zones sont entourées et numérotées
        - Sélectionnez une zone pour voir ses détails
        - Supprimez les zones incorrectes
        
        ### 4. Export
        - **Téléchargez le JSON** avec toutes les informations
        - Format compatible avec votre chatbot
        """)

if __name__ == "__main__":
    main()
