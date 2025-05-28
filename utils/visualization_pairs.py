"""
Module de visualisation adapté pour le système de paires de labels alternées
"""

import plotly.graph_objects as go
import pandas as pd
from typing import List, Dict, Optional
from .excel_utils import num_to_excel_col, get_cell_color
from .color_detector import hex_to_rgb

# Fonction helper pour adapter le format de color_palette
def create_excel_visualization_pairs(workbook, sheet_name, zones, selected_zone, color_palette):
    """Adapte le format de color_palette pour la visualisation"""
    adapted_palette = {
        'zone_color': color_palette['zone_color'],
        'zone_name': color_palette['zone_name'],
        'label_colors': {}
    }
    
    # Gérer les différents formats possibles
    if 'label_pairs' in color_palette:
        # Format avec paires
        if len(color_palette['label_pairs']) > 0:
            adapted_palette['label_colors']['h1'] = color_palette['label_pairs'][0]['horizontal']
            adapted_palette['label_colors']['v1'] = color_palette['label_pairs'][0]['vertical']
        if len(color_palette['label_pairs']) > 1:
            adapted_palette['label_colors']['h2'] = color_palette['label_pairs'][1]['horizontal']
            adapted_palette['label_colors']['v2'] = color_palette['label_pairs'][1]['vertical']
    elif 'h1_color' in color_palette:
        # Format direct
        adapted_palette['h1_color'] = color_palette.get('h1_color')
        adapted_palette['h2_color'] = color_palette.get('h2_color')
        adapted_palette['v1_color'] = color_palette.get('v1_color')
        adapted_palette['v2_color'] = color_palette.get('v2_color')
        adapted_palette['label_colors'] = {
            'h1': {'color': color_palette.get('h1_color'), 'name': 'H1'},
            'h2': {'color': color_palette.get('h2_color'), 'name': 'H2'},
            'v1': {'color': color_palette.get('v1_color'), 'name': 'V1'},
            'v2': {'color': color_palette.get('v2_color'), 'name': 'V2'}
        }
    
    return create_excel_visualization(workbook, sheet_name, zones, selected_zone, adapted_palette)

def create_zone_detail_view_with_pairs(workbook, sheet_name: str, zone: Dict, color_palette: Dict) -> go.Figure:
    """
    Vue détaillée d'une zone avec visualisation des paires de labels
    """
    ws = workbook[sheet_name]
    bounds = zone['bounds']
    
    # Marge pour voir les labels
    margin = 5  # Plus grande marge pour voir les alternances
    min_row = max(1, bounds['min_row'] - margin)
    max_row = min(ws.max_row, bounds['max_row'] + margin)
    min_col = max(1, bounds['min_col'] - margin)
    max_col = min(ws.max_column, bounds['max_col'] + margin)
    
    # Créer les données
    z_values = []
    text_values = []
    customdata = []
    
    for row in range(min_row, max_row + 1):
        row_values = []
        row_text = []
        row_custom = []
        
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            value = cell.value if cell.value is not None else ""
            row_text.append(str(value))
            row_values.append(1)
            row_custom.append(f"{num_to_excel_col(col)}{row}")
        
        z_values.append(row_values)
        text_values.append(row_text)
        customdata.append(row_custom)
    
    # Dimensions et coordonnées
    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    x_labels = [num_to_excel_col(i) for i in range(min_col, max_col + 1)]
    y_labels = [str(i) for i in range(min_row, max_row + 1)]
    x_coords = list(range(num_cols))
    y_coords = list(range(num_rows))
    
    # Créer la figure
    fig = go.Figure()
    
    # Heatmap de base
    fig.add_trace(go.Heatmap(
        z=z_values,
        x=x_coords,
        y=y_coords,
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
    
    # Créer les mappings
    zone_cells = {(c['row'], c['col']) for c in zone['cells']}
    label_map = {(l['row'], l['col']): l for l in zone.get('labels', [])}
    
    # Shapes pour les cellules et labels
    shapes = []
    
    # Cellules de la zone
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) in zone_cells:
                plot_col = col - min_col
                plot_row = row - min_row
                
                r, g, b = hex_to_rgb(color_palette['zone_color'])
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
    
    # Labels avec styles différenciés par paire et direction
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if (row, col) in label_map:
                label = label_map[(row, col)]
                plot_col = col - min_col
                plot_row = row - min_row
                
                # Déterminer la couleur et le style
                if 'pair_id' in label and label['pair_id'] < len(color_palette.get('label_pairs', [])):
                    pair = color_palette['label_pairs'][label['pair_id']]
                    
                    if label['direction'] == 'horizontal':
                        label_color = pair['horizontal']['color']
                        # Style horizontal : plus large, bordure épaisse en haut/bas
                        r, g, b = hex_to_rgb(label_color)
                        shapes.append(dict(
                            type="rect",
                            x0=plot_col - 0.48,
                            y0=plot_row - 0.38,
                            x1=plot_col + 0.48,
                            y1=plot_row + 0.38,
                            fillcolor=f'rgba({r},{g},{b},0.6)',
                            line=dict(width=2, color=f'rgb({r},{g},{b})'),
                            layer="below"
                        ))
                        # Indicateur de paire
                        shapes.append(dict(
                            type="rect",
                            x0=plot_col + 0.35,
                            y0=plot_row - 0.35,
                            x1=plot_col + 0.45,
                            y1=plot_row - 0.25,
                            fillcolor=f'rgb({r},{g},{b})',
                            line=dict(width=0),
                        ))
                    else:  # vertical
                        label_color = pair['vertical']['color']
                        # Style vertical : plus haut, bordure épaisse à gauche/droite
                        r, g, b = hex_to_rgb(label_color)
                        shapes.append(dict(
                            type="rect",
                            x0=plot_col - 0.38,
                            y0=plot_row - 0.48,
                            x1=plot_col + 0.38,
                            y1=plot_row + 0.48,
                            fillcolor=f'rgba({r},{g},{b},0.6)',
                            line=dict(width=2, color=f'rgb({r},{g},{b})'),
                            layer="below"
                        ))
                        # Indicateur de paire
                        shapes.append(dict(
                            type="rect",
                            x0=plot_col - 0.35,
                            y0=plot_row + 0.35,
                            x1=plot_col - 0.25,
                            y1=plot_row + 0.45,
                            fillcolor=f'rgb({r},{g},{b})',
                            line=dict(width=0),
                        ))
    
    # Cadre autour de la zone principale
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
    
    # Annotations pour les paires
    annotations = []
    
    # Ajouter une légende des paires dans le coin
    legend_text = "Paires:<br>"
    for i, pair in enumerate(color_palette.get('label_pairs', [])):
        h_color = pair['horizontal']['color']
        v_color = pair['vertical']['color']
        legend_text += f"P{i+1}: H=#{h_color[:6]} V=#{v_color[:6]}<br>"
    
    annotations.append(dict(
        x=num_cols - 1,
        y=0,
        text=legend_text,
        showarrow=False,
        bgcolor="white",
        bordercolor="black",
        borderwidth=1,
        font=dict(size=9),
        xanchor="right",
        yanchor="top"
    ))
    
    fig.update_layout(
        title=f"Zone {zone['id']} - Vue détaillée avec paires de labels",
        shapes=shapes,
        annotations=annotations,
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
        height=700,
        width=1000,
        plot_bgcolor='white',
        margin=dict(l=50, r=50, t=80, b=50)
    )
    
    return fig

def create_pair_analysis_chart(zones: List[Dict], color_palette: Dict) -> go.Figure:
    """
    Crée un graphique d'analyse des paires de labels
    """
    from collections import defaultdict
    
    # Analyser la distribution des labels par paire
    pair_stats = defaultdict(lambda: {
        'horizontal': {'count': 0, 'zones': set()},
        'vertical': {'count': 0, 'zones': set()}
    })
    
    for zone in zones:
        for label in zone.get('labels', []):
            if 'pair_id' in label and 'direction' in label:
                pair_id = label['pair_id']
                direction = label['direction']
                pair_stats[pair_id][direction]['count'] += 1
                pair_stats[pair_id][direction]['zones'].add(zone['id'])
    
    # Préparer les données pour le graphique
    data = []
    colors = []
    
    for pair_id in sorted(pair_stats.keys()):
        if pair_id < len(color_palette.get('label_pairs', [])):
            pair = color_palette['label_pairs'][pair_id]
            
            # Données horizontales
            data.append(go.Bar(
                name=f'P{pair_id+1} Horizontal',
                x=[f'Paire {pair_id+1}'],
                y=[pair_stats[pair_id]['horizontal']['count']],
                marker_color=f"#{pair['horizontal']['color']}",
                text=f"{len(pair_stats[pair_id]['horizontal']['zones'])} zones",
                textposition='auto',
            ))
            
            # Données verticales
            data.append(go.Bar(
                name=f'P{pair_id+1} Vertical',
                x=[f'Paire {pair_id+1}'],
                y=[pair_stats[pair_id]['vertical']['count']],
                marker_color=f"#{pair['vertical']['color']}",
                text=f"{len(pair_stats[pair_id]['vertical']['zones'])} zones",
                textposition='auto',
            ))
    
    # Créer la figure
    fig = go.Figure(data=data)
    
    fig.update_layout(
        title="Analyse des paires de labels",
        xaxis_title="Paires",
        yaxis_title="Nombre de labels",
        barmode='group',
        showlegend=True,
        height=400
    )
    
    return fig

def create_zone_pair_heatmap(zones: List[Dict], color_palette: Dict) -> go.Figure:
    """
    Crée une heatmap montrant la distribution des paires par zone
    """
    from collections import defaultdict
    
    # Préparer la matrice
    num_pairs = len(color_palette.get('label_pairs', []))
    zone_ids = [z['id'] for z in zones]
    
    # Matrice : zones x (paires * directions)
    matrix = []
    column_labels = []
    
    # Créer les labels de colonnes
    for i in range(num_pairs):
        column_labels.extend([f'P{i+1}_H', f'P{i+1}_V'])
    
    # Remplir la matrice
    for zone in zones:
        row = [0] * (num_pairs * 2)
        
        for label in zone.get('labels', []):
            if 'pair_id' in label and label['pair_id'] < num_pairs:
                col_idx = label['pair_id'] * 2
                if label['direction'] == 'horizontal':
                    row[col_idx] += 1
                else:
                    row[col_idx + 1] += 1
        
        matrix.append(row)
    
    # Créer la heatmap
    fig = go.Figure(data=go.Heatmap(
        z=matrix,
        x=column_labels,
        y=[f'Zone {id}' for id in zone_ids],
        colorscale='Blues',
        showscale=True,
        hoverongaps=False,
        hovertemplate='Zone: %{y}<br>Type: %{x}<br>Nombre: %{z}<extra></extra>'
    ))
    
    fig.update_layout(
        title="Distribution des labels par zone et paire",
        xaxis_title="Paires et directions",
        yaxis_title="Zones",
        height=400 + len(zones) * 20,  # Ajuster la hauteur selon le nombre de zones
        margin=dict(l=100)
    )
    
    return fig