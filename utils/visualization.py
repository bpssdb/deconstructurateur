"""
Module de visualisation pour l'affichage des feuilles Excel et des zones
"""

import plotly.graph_objects as go
import pandas as pd
from typing import List, Dict, Optional
from .excel_utils import num_to_excel_col, get_cell_color
from .color_detector import hex_to_rgb

def create_color_detection_preview(workbook, sheet_name: str, color_cells: Dict) -> go.Figure:
    """
    Crée un aperçu de la feuille avec toutes les couleurs détectées
    """
    ws = workbook[sheet_name]
    
    # Limiter la taille pour la performance
    max_row = min(ws.max_row, 50)
    max_col = min(ws.max_column, 20)
    
    # Créer la figure
    fig = go.Figure()
    
    # Créer une grille de couleurs
    for hex_color, cells in color_cells.items():
        if not cells:
            continue
            
        # Extraire les coordonnées
        rows = [c['row'] for c in cells if c['row'] <= max_row and c['col'] <= max_col]
        cols = [c['col'] for c in cells if c['row'] <= max_row and c['col'] <= max_col]
        
        if rows and cols:
            r, g, b = hex_to_rgb(hex_color)
            
            # Ajouter les cellules colorées comme scatter
            fig.add_trace(go.Scatter(
                x=[num_to_excel_col(c) for c in cols],
                y=rows,
                mode='markers',
                marker=dict(
                    size=20,
                    color=f'rgb({r},{g},{b})',
                    line=dict(width=1, color='black')
                ),
                name=f"#{hex_color}",
                text=[f"Valeur: {cells[i]['value']}" for i in range(len(rows))],
                hovertemplate='%{x}%{y}<br>%{text}<extra></extra>'
            ))
    
    # Mise en forme
    fig.update_layout(
        title="Aperçu des cellules colorées",
        xaxis=dict(
            title="Colonnes",
            tickmode='array',
            tickvals=[num_to_excel_col(i) for i in range(1, max_col + 1)],
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(
            title="Lignes",
            autorange="reversed",
            showgrid=True,
            gridcolor='lightgray'
        ),
        height=400,
        plot_bgcolor='white',
        showlegend=True
    )
    return fig
    
def create_zone_summary_dataframe(zones: List[Dict]) -> pd.DataFrame:
    """
    Crée un DataFrame résumant toutes les zones
    """
    if not zones:
        return pd.DataFrame()
    
    data = []
    for zone in zones:
        data.append({
            'Zone ID': zone['id'],
            'Lignes': f"{zone['bounds']['min_row']}-{zone['bounds']['max_row']}",
            'Colonnes': f"{num_to_excel_col(zone['bounds']['min_col'])}-{num_to_excel_col(zone['bounds']['max_col'])}",
            'Nombre de cellules': zone['cell_count'],
            'Nombre de labels': len(zone.get('labels', []))
        })
    
    return pd.DataFrame(data)

def create_excel_visualization(workbook, sheet_name: str, zones: List[Dict] = None, 
                                   selected_zone: Optional[int] = None, 
                                   color_mapping: Optional[Dict] = None) -> go.Figure:
    """
    Version corrigée de la visualisation Excel avec alignement correct des coordonnées
    """
    ws = workbook[sheet_name]
    
    # Obtenir les dimensions de la feuille (limitées pour la performance)
    max_row = min(ws.max_row, 100)
    max_col = min(ws.max_column, 26)
    
    print(f"DEBUG: Dimensions affichées: {max_row} x {max_col}")
    
    # Créer les données pour la heatmap
    z_values = []
    text_values = []
    
    # Préparer les données pour l'affichage
    for row in range(1, max_row + 1):
        row_values = []
        row_text = []
        
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            row_text.append(str(value))
            row_values.append(1)  # Valeur uniforme pour la heatmap
        
        z_values.append(row_values)
        text_values.append(row_text)
    
    # Créer les labels pour les axes - UTILISER DES INDICES NUMÉRIQUES
    x_labels = [num_to_excel_col(i) for i in range(1, max_col + 1)]  # ["A", "B", "C", ...]
    y_labels = [str(i) for i in range(1, max_row + 1)]              # ["1", "2", "3", ...]
    
    # COORDONNÉES POUR PLOTLY : utiliser des indices 0-based
    x_coords = list(range(max_col))  # [0, 1, 2, ...]
    y_coords = list(range(max_row))  # [0, 1, 2, ...]
    
    print(f"DEBUG: x_coords: {x_coords[:5]}...")
    print(f"DEBUG: y_coords: {y_coords[:5]}...")
    print(f"DEBUG: x_labels: {x_labels[:5]}...")
    print(f"DEBUG: y_labels: {y_labels[:5]}...")
    
    # Créer la figure Plotly
    fig = go.Figure()
    
    # Ajouter la heatmap de base avec les COORDONNÉES NUMÉRIQUES
    fig.add_trace(go.Heatmap(
        z=z_values,
        x=x_coords,  # CHANGEMENT: utiliser des indices numériques
        y=y_coords,  # CHANGEMENT: utiliser des indices numériques
        showscale=False,
        hoverongaps=False,
        colorscale=[[0, 'white'], [1, 'white']],
        text=text_values,
        texttemplate="%{text}",
        textfont={"size": 10},
        #hovertemplate='Cellule: %{x}%{y}<br>Valeur: %{text}<extra></extra>',
        customdata=[[f"{x_labels[j]}{y_labels[i]}" for j in range(max_col)] for i in range(max_row)],
        hovertemplate='Cellule: %{customdata}<br>Valeur: %{text}<extra></extra>'
    ))
    
    # Ajouter les rectangles colorés pour les zones avec COORDONNÉES ALIGNÉES
    shapes = []
    annotations = []
    
    if zones and color_mapping:
        for zone in zones:
            bounds = zone['bounds']
            
            # VÉRIFIER QUE LES BOUNDS SONT DANS LES LIMITES D'AFFICHAGE
            if (bounds['min_row'] > max_row or bounds['min_col'] > max_col or
                bounds['max_row'] < 1 or bounds['max_col'] < 1):
                print(f"DEBUG: Zone {zone['id']} hors limites d'affichage")
                continue
            
            # CONVERTIR LES COORDONNÉES EXCEL EN COORDONNÉES PLOTLY (0-based)
            # Excel: colonne 1 = index 0, ligne 1 = index 0
            plot_min_col = bounds['min_col'] - 1  # Colonne 1 -> index 0
            plot_max_col = bounds['max_col'] - 1  # Colonne 2 -> index 1
            plot_min_row = bounds['min_row'] - 1  # Ligne 1 -> index 0
            plot_max_row = bounds['max_row'] - 1  # Ligne 2 -> index 1
            
            print(f"DEBUG: Zone {zone['id']}")
            print(f"  Excel bounds: cols {bounds['min_col']}-{bounds['max_col']}, rows {bounds['min_row']}-{bounds['max_row']}")
            print(f"  Plot coords: cols {plot_min_col}-{plot_max_col}, rows {plot_min_row}-{plot_max_row}")
            
            # Couleur de la zone
            zone_hex = color_mapping['zone_color']
            r, g, b = hex_to_rgb(zone_hex)
            zone_color = f'rgba({r}, {g}, {b}, 0.3)' if zone['id'] != selected_zone else 'rgba(0, 104, 201, 0.5)'
            
            # COORDONNÉES CORRIGÉES POUR LES RECTANGLES
            # Ajouter une marge de 0.5 pour bien centrer sur les cellules
            shapes.append(dict(
                type="rect",
                x0=plot_min_col - 0.5,   # CHANGEMENT: coordonnées alignées
                y0=plot_min_row - 0.5,   # CHANGEMENT: coordonnées alignées  
                x1=plot_max_col + 0.5,   # CHANGEMENT: coordonnées alignées
                y1=plot_max_row + 0.5,   # CHANGEMENT: coordonnées alignées
                line=dict(color=zone_color, width=2),
                fillcolor=zone_color,
                layer="below"
            ))
            
            # ANNOTATION AVEC COORDONNÉES CORRIGÉES
            annotations.append(dict(
                x=plot_min_col,  # CHANGEMENT: coordonnée alignée
                y=plot_min_row,  # CHANGEMENT: coordonnée alignée
                text=f"Zone {zone['id']}",
                showarrow=False,
                bgcolor="white",
                bordercolor="black",
                borderwidth=1,
                font=dict(size=10)
            ))
            
            # Ajouter des indicateurs pour les labels avec COORDONNÉES CORRIGÉES
            for label in zone.get('labels', []):
                # Vérifier que le label est dans les limites d'affichage
                if label['row'] > max_row or label['col'] > max_col:
                    continue
                
                # Déterminer la couleur du label
                label_color = '#888888'  # Couleur par défaut
                if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                    label_color = color_mapping['label_colors'][label['type']]['color']
                elif label['type'] == 'label1' and 'label1_color' in color_mapping:
                    label_color = color_mapping['label1_color']
                elif label['type'] == 'label2' and 'label2_color' in color_mapping:
                    label_color = color_mapping['label2_color']
                
                r, g, b = hex_to_rgb(label_color)
                
                # COORDONNÉES CORRIGÉES POUR LES LABELS
                plot_label_col = label['col'] - 1  # Convertir en 0-based
                plot_label_row = label['row'] - 1  # Convertir en 0-based
                
                shapes.append(dict(
                    type="rect",
                    x0=plot_label_col - 0.4,  # CHANGEMENT: coordonnée alignée
                    y0=plot_label_row - 0.4,  # CHANGEMENT: coordonnée alignée
                    x1=plot_label_col + 0.4,  # CHANGEMENT: coordonnée alignée
                    y1=plot_label_row + 0.4,  # CHANGEMENT: coordonnée alignée
                    line=dict(color=f'rgb({r}, {g}, {b})', width=2),
                    fillcolor=f'rgba({r}, {g}, {b}, 0.7)',
                ))
    
    # CONFIGURATION DES AXES CORRIGÉE
    fig.update_layout(
        shapes=shapes,
        annotations=annotations,
        xaxis=dict(
            title="Colonnes",
            side="top",
            tickmode='array',
            tickvals=x_coords,        # CHANGEMENT: indices numériques
            ticktext=x_labels,        # Labels Excel correspondants
            showgrid=True,
            gridcolor='lightgray',
            zeroline=False,
            range=[-0.5, max_col - 0.5]  # AJOUT: limiter la plage
        ),
        yaxis=dict(
            title="Lignes", 
            autorange="reversed",     # Garder l'inversion pour que ligne 1 soit en haut
            tickmode='array',
            tickvals=y_coords,        # CHANGEMENT: indices numériques
            ticktext=y_labels,        # Labels Excel correspondants
            showgrid=True,
            gridcolor='lightgray',
            zeroline=False,
            range=[max_row - 0.5, -0.5]  # AJOUT: limiter la plage (inversée)
        ),
        width=1000,
        height=600,
        plot_bgcolor='white',
        margin=dict(l=50, r=50, t=50, b=50),
        title="Vue Excel - Coordonnées alignées"
    )
    
    return fig

def create_color_preview_html(colors: List[Dict]) -> str:
    """
    Crée un tableau HTML avec aperçu des couleurs
    """
    if not colors:
        return "<p>Aucune couleur détectée</p>"
    
    html = """
    <table style="width:100%; border-collapse: collapse;">
        <thead>
            <tr style="background-color: #f0f0f0;">
                <th style="padding: 10px; text-align: left;">Aperçu</th>
                <th style="padding: 10px; text-align: left;">Nom</th>
                <th style="padding: 10px; text-align: left;">Code</th>
                <th style="padding: 10px; text-align: left;">Occurrences</th>
                <th style="padding: 10px; text-align: left;">Exemples</th>
            </tr>
        </thead>
        <tbody>
    """
    
    for color in colors:
        # Utiliser la liste d'exemples si disponible, sinon construire à partir des cellules
        if 'examples' in color:
            examples = ", ".join(color['examples'][:3])
            if len(color['examples']) > 3:
                examples += "..."
        else:
            examples = ", ".join([c['address'] for c in color.get('cells', [])[:3]])
            if len(color.get('cells', [])) > 3:
                examples += "..."
        
        # Ajouter une note sur les cellules fusionnées si présentes
        merged_note = ""
        if color.get('merged_count', 0) > 0:
            merged_note = f" <small>({color['merged_count']} fusionnées)</small>"
        
        html += f"""
        <tr>
            <td style="padding: 10px;">
                <div class="color-preview" style="background-color: #{color['hex']};"></div>
            </td>
            <td style="padding: 10px;">{color['name']}</td>
            <td style="padding: 10px;">#{color['hex']}</td>
            <td style="padding: 10px;">{color['count']}{merged_note}</td>
            <td style="padding: 10px; font-size: 0.9em;">{examples}</td>
        </tr>
        """
    
    html += """
        </tbody>
    </table>
    """
    
    return html

def create_zone_detail_view(workbook, sheet_name: str, zone: Dict, color_mapping: Dict) -> go.Figure:
    """
    Version corrigée de create_zone_detail_view avec alignement des coordonnées
    """
    ws = workbook[sheet_name]
    bounds = zone['bounds']
    
    # Ajouter une marge pour voir les labels autour
    margin = 3
    min_row = max(1, bounds['min_row'] - margin)
    max_row = min(ws.max_row, bounds['max_row'] + margin)
    min_col = max(1, bounds['min_col'] - margin)
    max_col = min(ws.max_column, bounds['max_col'] + margin)
    
    print(f"DEBUG Zone detail: Excel range rows {min_row}-{max_row}, cols {min_col}-{max_col}")
    
    # Créer les données pour la heatmap
    z_values = []
    text_values = []
    customdata = []
    
    # Préparer les données pour l'affichage
    for row in range(min_row, max_row + 1):
        row_values = []
        row_text = []
        row_custom = []
        
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            row_text.append(str(value))
            row_values.append(1)  # Valeur uniforme pour la heatmap
            row_custom.append(f"{num_to_excel_col(col)}{row}")  # Référence Excel
        
        z_values.append(row_values)
        text_values.append(row_text)
        customdata.append(row_custom)
    
    # Calculer les dimensions d'affichage
    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    
    # Créer les labels pour les axes
    x_labels = [num_to_excel_col(i) for i in range(min_col, max_col + 1)]
    y_labels = [str(i) for i in range(min_row, max_row + 1)]
    
    # Coordonnées numériques pour Plotly
    x_coords = list(range(num_cols))
    y_coords = list(range(num_rows))
    
    print(f"DEBUG: Display size: {num_rows} x {num_cols}")
    print(f"DEBUG: x_labels: {x_labels}")
    print(f"DEBUG: y_labels: {y_labels}")
    
    # Créer la figure
    fig = go.Figure()
    
    # Ajouter la heatmap de base avec le texte
    fig.add_trace(go.Heatmap(
        z=z_values,
        x=x_coords,  # Coordonnées numériques
        y=y_coords,  # Coordonnées numériques
        showscale=False,
        text=text_values,
        texttemplate="%{text}",
        textfont={"size": 12},
        customdata=customdata,
        hovertemplate='%{customdata}: %{text}<extra></extra>',
        colorscale=[[0, 'white'], [1, 'white']],
        zmin=0,
        zmax=1
    ))
    
    # Créer un mapping des cellules de la zone et des labels
    zone_cells = {(c['row'], c['col']) for c in zone['cells']}
    label_cells = {(l['row'], l['col']): l for l in zone.get('labels', [])}
    
    # Ajouter les rectangles colorés avec coordonnées corrigées
    shapes = []
    
    # Cellules de la zone
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) in zone_cells:
                # Convertir les coordonnées Excel en coordonnées Plotly
                plot_col = col - min_col  # Position relative dans l'affichage
                plot_row = row - min_row  # Position relative dans l'affichage
                
                r, g, b = hex_to_rgb(color_mapping['zone_color'])
                shapes.append(dict(
                    type="rect",
                    x0=plot_col - 0.45,
                    y0=plot_row - 0.45,
                    x1=plot_col + 0.45,
                    y1=plot_row + 0.45,
                    fillcolor=f'rgba({r},{g},{b},0.3)',
                    line=dict(width=0.5, color=f'rgb({r},{g},{b})'),
                    layer="below"
                ))
    
    # Labels (par-dessus les zones)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) in label_cells:
                label = label_cells[(row, col)]
                
                # Convertir les coordonnées Excel en coordonnées Plotly
                plot_col = col - min_col
                plot_row = row - min_row
                
                # Déterminer la couleur du label
                label_color = None
                if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                    label_color = color_mapping['label_colors'][label['type']]['color']
                elif label['type'] == 'horizontal' and 'horizontal' in color_mapping.get('label_colors', {}):
                    label_color = color_mapping['label_colors']['horizontal']['color']
                elif label['type'] == 'vertical' and 'vertical' in color_mapping.get('label_colors', {}):
                    label_color = color_mapping['label_colors']['vertical']['color']
                
                if label_color:
                    r, g, b = hex_to_rgb(label_color)
                    shapes.append(dict(
                        type="rect",
                        x0=plot_col - 0.45,
                        y0=plot_row - 0.45,
                        x1=plot_col + 0.45,
                        y1=plot_row + 0.45,
                        fillcolor=f'rgba({r},{g},{b},0.5)',
                        line=dict(width=1, color=f'rgb({r},{g},{b})'),
                        layer="below"
                    ))
    
    # Ajouter un cadre autour de la zone principale avec coordonnées corrigées
    zone_min_row_plot = bounds['min_row'] - min_row
    zone_max_row_plot = bounds['max_row'] - min_row
    zone_min_col_plot = bounds['min_col'] - min_col
    zone_max_col_plot = bounds['max_col'] - min_col
    
    shapes.append(dict(
        type="rect",
        x0=zone_min_col_plot - 0.5,
        y0=zone_min_row_plot - 0.5,
        x1=zone_max_col_plot + 0.5,
        y1=zone_max_row_plot + 0.5,
        fillcolor="rgba(0,0,0,0)",
        line=dict(width=3, color='blue')
    ))
    
    fig.update_layout(
        title=f"Zone {zone['id']} - Vue détaillée (coordonnées corrigées)",
        shapes=shapes,
        xaxis=dict(
            title="Colonnes",
            side="top",
            tickmode='array',
            tickvals=x_coords,
            ticktext=x_labels,
            showgrid=True,
            gridcolor='lightgray',
            zeroline=False,
            constrain="domain"
        ),
        yaxis=dict(
            title="Lignes",
            autorange="reversed",
            tickmode='array',
            tickvals=y_coords,
            ticktext=y_labels,
            showgrid=True,
            gridcolor='lightgray',
            zeroline=False,
            scaleanchor="x",
            scaleratio=1,
            constrain="domain"
        ),
        height=600,
        width=900,
        plot_bgcolor='white',
        margin=dict(l=50, r=50, t=80, b=50)
    )
    
    return fig
    
def create_dataframe_view(workbook, sheet_name: str, zones: List[Dict] = None, 
                         color_mapping: Optional[Dict] = None, max_rows: int = 50) -> pd.DataFrame:
    """
    Crée une vue DataFrame stylée de la feuille Excel avec coloration des zones
    """
    ws = workbook[sheet_name]
    
    # Limiter les dimensions pour la performance
    max_row = min(ws.max_row, max_rows)
    max_col = min(ws.max_column, 26)
    
    # Créer un mapping des cellules colorées
    colored_cells = {}
    if zones and color_mapping:
        for zone in zones:
            # Cellules de la zone
            for cell in zone['cells']:
                if cell['row'] <= max_row and cell['col'] <= max_col:
                    colored_cells[(cell['row'], cell['col'])] = {
                        'color': color_mapping['zone_color'],
                        'type': 'zone',
                        'zone_id': zone['id']
                    }
            
            # Labels de la zone
            for label in zone.get('labels', []):
                if label['row'] <= max_row and label['col'] <= max_col:
                    # Déterminer la couleur du label
                    label_color = '#888888'  # Couleur par défaut
                    if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                        label_color = color_mapping['label_colors'][label['type']]['color']
                    elif label['type'] == 'label1' and 'label1_color' in color_mapping:
                        label_color = color_mapping['label1_color']
                    elif label['type'] == 'label2' and 'label2_color' in color_mapping:
                        label_color = color_mapping['label2_color']
                    
                    colored_cells[(label['row'], label['col'])] = {
                        'color': label_color,
                        'type': 'label',
                        'label_type': label['type'],
                        'zone_id': zone['id']
                    }
    
    # Créer les données du DataFrame
    data = []
    columns = [num_to_excel_col(i) for i in range(1, max_col + 1)]
    
    for row in range(1, max_row + 1):
        row_data = []
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            row_data.append(str(value))
        data.append(row_data)
    
    # Créer le DataFrame
    df = pd.DataFrame(data, columns=columns, index=range(1, max_row + 1))
    
    # Si pas de zones ou de mapping de couleurs, retourner le DataFrame simple
    if not zones or not color_mapping:
        return df
    
    # Appliquer le style avec les couleurs
    def style_cells(val):
        """Fonction pour styler les cellules"""
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        
        for row_idx, row_num in enumerate(df.index, 1):
            for col_idx, col_name in enumerate(df.columns):
                col_num = excel_col_to_num(col_name)
                
                if (row_num, col_num) in colored_cells:
                    cell_info = colored_cells[(row_num, col_num)]
                    color = cell_info['color']
                    
                    # Calculer une couleur de texte contrastante
                    r, g, b = hex_to_rgb(color)
                    brightness = (r * 299 + g * 587 + b * 114) / 1000
                    text_color = 'white' if brightness < 128 else 'black'
                    
                    if cell_info['type'] == 'zone':
                        styles.iloc[row_idx-1, col_idx] = f'background-color: #{color}; color: {text_color}; border: 2px solid #{color};'
                    elif cell_info['type'] == 'label':
                        styles.iloc[row_idx-1, col_idx] = f'background-color: #{color}; color: {text_color}; border: 2px solid #{color}; font-weight: bold;'
        
        return styles
    
    # Appliquer le style
    try:
        styled_df = df.style.apply(style_cells, axis=None)
        styled_df = styled_df.set_table_attributes('style="border-collapse: collapse; font-size: 12px;"')
        return styled_df
    except Exception as e:
        # En cas d'erreur avec le style, retourner le DataFrame simple
        print(f"Erreur lors de l'application du style: {e}")
        return df


def excel_col_to_num(col_str: str) -> int:
    """
    Convertit une référence de colonne Excel (A, B, AA, etc.) en numéro
    """
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def create_zone_detail_table_view(workbook, sheet_name: str, zone: Dict, color_mapping: Dict) -> pd.DataFrame:
    """
    Crée une vue tableau détaillée pour une zone spécifique avec coloration
    """
    ws = workbook[sheet_name]
    bounds = zone['bounds']
    
    # Ajouter une marge pour voir les labels autour
    margin = 3
    min_row = max(1, bounds['min_row'] - margin)
    max_row = min(ws.max_row, bounds['max_row'] + margin)
    min_col = max(1, bounds['min_col'] - margin)
    max_col = min(ws.max_column, bounds['max_col'] + margin)
    
    print(f"Vue tableau zone {zone['id']}: lignes {min_row}-{max_row}, colonnes {min_col}-{max_col}")
    
    # Créer un mapping des cellules de la zone et des labels
    zone_cells = {(c['row'], c['col']) for c in zone['cells']}
    label_cells = {(l['row'], l['col']): l for l in zone.get('labels', [])}
    
    # Créer les données du DataFrame
    data = []
    columns = [num_to_excel_col(i) for i in range(min_col, max_col + 1)]
    
    for row in range(min_row, max_row + 1):
        row_data = []
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            row_data.append(str(value))
        data.append(row_data)
    
    # Créer le DataFrame
    df = pd.DataFrame(data, columns=columns, index=range(min_row, max_row + 1))
    
    
    def style_zone_cells(val):
        """Fonction pour styler les cellules de la zone"""
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        
        for row_idx in range(len(df)):
            actual_row = df.index[row_idx]  # Ligne réelle dans Excel
            
            for col_idx in range(len(df.columns)):
                col_name = df.columns[col_idx]  # Nom de colonne (A, B, C...)
                col_num = excel_col_to_num(col_name)  # Numéro de colonne
                
                # Vérifier si c'est une cellule de zone
                if (actual_row, col_num) in zone_cells:
                    zone_color = color_mapping['zone_color']
                    r, g, b = hex_to_rgb(zone_color)
                    brightness = (r * 299 + g * 587 + b * 114) / 1000
                    text_color = 'white' if brightness < 128 else 'black'
                    
                    styles.iloc[row_idx, col_idx] = f'background-color: #{zone_color}; color: {text_color}; font-weight: bold; border: 2px solid #{zone_color};'
                
                # Vérifier si c'est un label (priorité sur la zone)
                elif (actual_row, col_num) in label_cells:
                    label = label_cells[(actual_row, col_num)]
                    
                    # Déterminer la couleur du label
                    label_color = None
                    if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                        label_color = color_mapping['label_colors'][label['type']]['color']
                    
                    if label_color:
                        r, g, b = hex_to_rgb(label_color)
                        brightness = (r * 299 + g * 587 + b * 114) / 1000
                        text_color = 'white' if brightness < 128 else 'black'
                        
                        styles.iloc[row_idx, col_idx] = f'background-color: #{label_color}; color: {text_color}; font-weight: bold; border: 3px solid #{label_color}; box-shadow: 0 0 5px rgba({r},{g},{b},0.7);'
        
        return styles

 # Appliquer le style
    try:
        styled_df = df.style.apply(style_zone_cells, axis=None)
        styled_df = styled_df.set_table_attributes('style="border-collapse: collapse; font-size: 14px;"')
        styled_df = styled_df.set_caption(f"Zone {zone['id']} - Vue détaillée tableau")
        return styled_df
    except Exception as e:
        print(f"Erreur lors de l'application du style: {e}")
        return df


def create_zone_detail_table_view_enhanced(workbook, sheet_name: str, zone: Dict, color_mapping: Dict) -> pd.DataFrame:
    """
    Version améliorée de la vue tableau avec gestion avancée du style et marqueurs visuels
    """
    ws = workbook[sheet_name]
    bounds = zone['bounds']
    
    # Calculer la zone d'affichage avec marge
    margin = 3
    min_row = max(1, bounds['min_row'] - margin)
    max_row = min(ws.max_row, bounds['max_row'] + margin)
    min_col = max(1, bounds['min_col'] - margin)
    max_col = min(ws.max_column, bounds['max_col'] + margin)
    
    # Créer les mappings
    zone_cells = {(c['row'], c['col']) for c in zone['cells']}
    label_cells = {(l['row'], l['col']): l for l in zone.get('labels', [])}
    
    # Créer le DataFrame avec marqueurs visuels
    columns = [num_to_excel_col(i) for i in range(min_col, max_col + 1)]
    data = []
    
    for row in range(min_row, max_row + 1):
        row_data = []
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            
            # Ajouter des indicateurs visuels dans le texte
            if (row, col) in zone_cells:
                # Cellule de zone - ajouter un marqueur
                value = f"🔵 {value}" if value else "🔵"
            elif (row, col) in label_cells:
                # Label - ajouter un marqueur selon le type
                label = label_cells[(row, col)]
                marker = "🏷️" if label['type'] == 'horizontal' else "📍"
                value = f"{marker} {value}" if value else marker
            
            row_data.append(str(value))
        data.append(row_data)
    
    df = pd.DataFrame(data, columns=columns, index=range(min_row, max_row + 1))
    
    # Style avancé avec CSS
    def enhanced_style(x):
        """Style avancé pour le tableau"""
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        
        for row_idx in range(len(df)):
            actual_row = df.index[row_idx]
            
            for col_idx in range(len(df.columns)):
                col_name = df.columns[col_idx]
                col_num = excel_col_to_num(col_name)
                
                if (actual_row, col_num) in zone_cells:
                    # Style pour cellules de zone
                    zone_color = color_mapping['zone_color']
                    r, g, b = hex_to_rgb(zone_color)
                    
                    styles.iloc[row_idx, col_idx] = f'background-color: rgba({r}, {g}, {b}, 0.3); border: 3px solid #{zone_color}; font-weight: bold; text-align: center;'
                
                elif (actual_row, col_num) in label_cells:
                    # Style pour labels
                    label = label_cells[(actual_row, col_num)]
                    label_color = None
                    
                    if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                        label_color = color_mapping['label_colors'][label['type']]['color']
                    
                    if label_color:
                        r, g, b = hex_to_rgb(label_color)
                        styles.iloc[row_idx, col_idx] = f'background-color: rgba({r}, {g}, {b}, 0.5); border: 2px solid #{label_color}; font-weight: bold; font-style: italic; text-align: center;'
        
        return styles
    
    try:
        styled_df = df.style.apply(enhanced_style, axis=None)
        styled_df = styled_df.set_table_attributes('style="border-collapse: collapse; font-size: 12px;"')
        styled_df = styled_df.set_caption(f"<h3>Zone {zone['id']} - Vue détaillée avec légende</h3><p>🔵 = Cellules de zone | 🏷️📍 = Labels</p>")
        return styled_df
    except Exception as e:
        print(f"Erreur style avancé: {e}")
        return df


def display_zone_comparison_table(workbook, sheet_name: str, zone: Dict, color_mapping: Dict):
    """
    Affiche une comparaison entre les données détectées et la réalité Excel
    """
    ws = workbook[sheet_name]
    
    # Analyser les cellules de la zone
    zone_analysis = []
    
    print(f"Analyse comparative zone {zone['id']}")
    
    for cell_info in zone.get('cells', [])[:10]:  # Limiter à 10 cellules
        row, col = cell_info['row'], cell_info['col']
        excel_addr = f"{num_to_excel_col(col)}{row}"
        
        try:
            excel_cell = ws.cell(row=row, column=col)
            value = excel_cell.value
            detected_color = get_cell_color(excel_cell)
            expected_color = color_mapping['zone_color']
            
            zone_analysis.append({
                'Cellule': excel_addr,
                'Valeur': str(value) if value else "(vide)",
                'Couleur détectée': detected_color or "Aucune",
                'Couleur attendue': expected_color,
                'Correspondance': '✅' if (detected_color and detected_color.upper().replace('#', '') == expected_color.upper().replace('#', '')) else '❌',
                'Dans zone bounds': '✅' if (zone['bounds']['min_row'] <= row <= zone['bounds']['max_row'] and 
                                          zone['bounds']['min_col'] <= col <= zone['bounds']['max_col']) else '❌'
            })
        except Exception as e:
            zone_analysis.append({
                'Cellule': excel_addr,
                'Valeur': "ERREUR",
                'Couleur détectée': str(e),
                'Couleur attendue': expected_color,
                'Correspondance': '❌',
                'Dans zone bounds': '❌'
            })
    
    # Analyser les labels
    label_analysis = []
    for label in zone.get('labels', []):
        row, col = label['row'], label['col']
        excel_addr = f"{num_to_excel_col(col)}{row}"
        
        try:
            excel_cell = ws.cell(row=row, column=col)
            value = excel_cell.value
            detected_color = get_cell_color(excel_cell)
            
            expected_color = None
            if 'label_colors' in color_mapping and label['type'] in color_mapping['label_colors']:
                expected_color = color_mapping['label_colors'][label['type']]['color']
            
            label_analysis.append({
                'Cellule': excel_addr,
                'Type': label['type'],
                'Valeur': str(value) if value else "(vide)",
                'Couleur détectée': detected_color or "Aucune",
                'Couleur attendue': expected_color or "Non définie",
                'Correspondance': '✅' if (expected_color and detected_color and 
                                        detected_color.upper().replace('#', '') == expected_color.upper().replace('#', '')) else '❌'
            })
        except Exception as e:
            label_analysis.append({
                'Cellule': excel_addr,
                'Type': label['type'],
                'Valeur': "ERREUR",
                'Couleur détectée': str(e),
                'Couleur attendue': expected_color or "Non définie",
                'Correspondance': '❌'
            })
    
    return pd.DataFrame(zone_analysis), pd.DataFrame(label_analysis)