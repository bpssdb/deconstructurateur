"""
Module d'export des données en JSON
"""

import json
from datetime import datetime
from typing import List, Dict
from .excel_utils import num_to_excel_col

def export_to_json(zones: List[Dict], sheet_name: str, color_palette: Dict) -> str:
    """
    Exporte les zones et leurs métadonnées en JSON
    """
    # Construire la palette de couleurs pour l'export
    export_palette = {
        "zone_color": f"#{color_palette['zone_color']}",
        "zone_name": color_palette['zone_name']
    }
    
    # Ajouter les couleurs de labels (format flexible)
    if 'label_colors' in color_palette:
        export_palette['label_colors'] = {}
        for label_type, label_info in color_palette['label_colors'].items():
            export_palette['label_colors'][label_type] = {
                "color": f"#{label_info['color']}",
                "name": label_info['name']
            }
    # Support ancien format (rétrocompatibilité)
    elif 'label1_color' in color_palette:
        export_palette['label1_color'] = f"#{color_palette['label1_color']}"
        export_palette['label1_name'] = color_palette.get('label1_name', 'Label 1')
        export_palette['label2_color'] = f"#{color_palette['label2_color']}"
        export_palette['label2_name'] = color_palette.get('label2_name', 'Label 2')
    
    export_data = {
        "date_export": datetime.now().isoformat(),
        "sheet_name": sheet_name,
        "color_palette": export_palette,
        "zones": []
    }
    
    for zone in zones:
        zone_data = {
            "id": zone['id'],
            "bounds": {
                "min_row": zone['bounds']['min_row'],
                "max_row": zone['bounds']['max_row'],
                "min_col": zone['bounds']['min_col'],
                "max_col": zone['bounds']['max_col'],
                "min_col_letter": num_to_excel_col(zone['bounds']['min_col']),
                "max_col_letter": num_to_excel_col(zone['bounds']['max_col'])
            },
            "cell_count": zone['cell_count'],
            "cells": format_cells_for_export(zone['cells']),
            "labels": format_labels_for_export(zone.get('labels', []))
        }
        export_data["zones"].append(zone_data)
    
    return json.dumps(export_data, indent=2, ensure_ascii=False)

def format_cells_for_export(cells: List[Dict]) -> List[Dict]:
    """
    Formate les cellules pour l'export JSON
    """
    formatted_cells = []
    for cell in cells:
        formatted_cells.append({
            "address": f"{num_to_excel_col(cell['col'])}{cell['row']}",
            "row": cell['row'],
            "col": cell['col'],
            "col_letter": num_to_excel_col(cell['col']),
            "value": str(cell['value']) if cell['value'] is not None else ""
        })
    return formatted_cells

def format_labels_for_export(labels: List[Dict]) -> List[Dict]:
    """
    Formate les labels pour l'export JSON
    """
    formatted_labels = []
    for label in labels:
        formatted_labels.append({
            "address": f"{num_to_excel_col(label['col'])}{label['row']}",
            "row": label['row'],
            "col": label['col'],
            "col_letter": num_to_excel_col(label['col']),
            "value": str(label['value']) if label['value'] is not None else "",
            "type": label['type'],
            "position": label['position']
        })
    return formatted_labels

def export_to_csv(zones: List[Dict]) -> str:
    """
    Exporte les zones en format CSV pour analyse
    """
    import csv
    import io
    
    output = io.StringIO()
    writer = csv.writer(output)
    
    # En-tête
    writer.writerow([
        'Zone ID', 
        'Min Row', 
        'Max Row', 
        'Min Col', 
        'Max Col', 
        'Cell Count', 
        'Label Count',
        'Label Values'
    ])
    
    # Données
    for zone in zones:
        label_values = "; ".join([
            f"{label['value']} ({label['type']})" 
            for label in zone.get('labels', [])
        ])
        
        writer.writerow([
            zone['id'],
            zone['bounds']['min_row'],
            zone['bounds']['max_row'],
            num_to_excel_col(zone['bounds']['min_col']),
            num_to_excel_col(zone['bounds']['max_col']),
            zone['cell_count'],
            len(zone.get('labels', [])),
            label_values
        ])
    
    return output.getvalue()

def create_zone_report(zones: List[Dict], color_palette: Dict) -> str:
    """
    Crée un rapport textuel des zones détectées
    """
    report = f"""
RAPPORT DE DÉTECTION DES ZONES
==============================

Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

PALETTE DE COULEURS
-------------------
- Zones de données: {color_palette['zone_name']} (#{color_palette['zone_color']})
- Labels type 1: {color_palette['label1_name']} (#{color_palette['label1_color']})
- Labels type 2: {color_palette['label2_name']} (#{color_palette['label2_color']})

RÉSUMÉ
------
Nombre total de zones: {len(zones)}
Nombre total de cellules: {sum(z['cell_count'] for z in zones)}
Nombre total de labels: {sum(len(z.get('labels', [])) for z in zones)}

DÉTAIL DES ZONES
----------------
"""
    
    for zone in zones:
        report += f"\nZone {zone['id']}:\n"
        report += f"  Position: Lignes {zone['bounds']['min_row']}-{zone['bounds']['max_row']}, "
        report += f"Colonnes {num_to_excel_col(zone['bounds']['min_col'])}-{num_to_excel_col(zone['bounds']['max_col'])}\n"
        report += f"  Nombre de cellules: {zone['cell_count']}\n"
        
        if zone.get('labels'):
            report += f"  Labels ({len(zone['labels'])}):\n"
            for label in zone['labels']:
                report += f"    - {label['value']} ({label['type']}, position: {label['position']})\n"
        else:
            report += "  Labels: Aucun\n"
    
    return report