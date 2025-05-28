"""
Module de détection et d'analyse des couleurs dans les fichiers Excel
"""

import colorsys
from collections import defaultdict, Counter
from typing import List, Dict, Tuple
from .excel_utils import get_cell_color, num_to_excel_col, get_cell_value, rgb_to_hex, get_merged_cells_info

def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convertit hexadécimal en RGB"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def get_color_name(hex_color: str) -> str:
    """Retourne un nom descriptif pour une couleur"""
    try:
        r, g, b = hex_to_rgb(hex_color)
        h, s, v = colorsys.rgb_to_hsv(r/255, g/255, b/255)
        
        # Déterminer la teinte principale
        if s < 0.1:  # Gris
            if v < 0.3:
                return "Noir"
            elif v > 0.7:
                return "Blanc"
            else:
                return "Gris"
        else:
            hue_deg = h * 360
            if hue_deg < 15 or hue_deg >= 345:
                return "Rouge"
            elif hue_deg < 30:
                return "Rouge-Orange"
            elif hue_deg < 45:
                return "Orange"
            elif hue_deg < 60:
                return "Orange-Jaune"
            elif hue_deg < 75:
                return "Jaune"
            elif hue_deg < 120:
                return "Vert-Jaune"
            elif hue_deg < 150:
                return "Vert"
            elif hue_deg < 180:
                return "Vert-Cyan"
            elif hue_deg < 210:
                return "Cyan"
            elif hue_deg < 240:
                return "Bleu-Cyan"
            elif hue_deg < 270:
                return "Bleu"
            elif hue_deg < 300:
                return "Violet"
            elif hue_deg < 330:
                return "Magenta"
            else:
                return "Rose"
    except:
        return "Inconnu"

def detect_all_colors(workbook, sheet_name: str) -> Tuple[List[Dict], Dict[str, List[Dict]]]:
    """
    Détecte toutes les couleurs présentes dans la feuille Excel
    Retourne un résumé des couleurs et un dictionnaire des cellules par couleur
    """
    ws = workbook[sheet_name]
    color_cells = defaultdict(list)
    color_counts = Counter()
    
    # Obtenir les informations sur les cellules fusionnées
    merged_info = get_merged_cells_info(ws)
    
    # Debug: compter les cellules analysées
    total_cells = 0
    cells_with_fill = 0
    
    # Parcourir toutes les cellules
    for row_idx, row in enumerate(ws.iter_rows()):
        for col_idx, cell in enumerate(row):
            total_cells += 1
            
            # Vérifier différentes propriétés de remplissage
            hex_color = None
            
            # Méthode 1: utiliser get_cell_color qui gère plus de cas
            hex_color = get_cell_color(cell)
            
            # Si on n'a pas trouvé avec get_cell_color, essayer d'autres méthodes
            if not hex_color and hasattr(cell, 'fill') and cell.fill:
                cells_with_fill += 1
            
            # Ignorer les cellules sans couleur, transparentes ou blanches
            if hex_color and hex_color not in ["FFFFFF", "00000000", None, ""]:
                # Nettoyer le code couleur
                hex_color = hex_color.upper().lstrip('#')
                if len(hex_color) == 8:  # ARGB
                    hex_color = hex_color[2:]  # Enlever le canal alpha
                if len(hex_color) == 6:  # RGB valide
                    # Vérifier si c'est vraiment blanc (tolérance pour les blancs cassés)
                    r, g, b = hex_to_rgb(hex_color)
                    if r > 250 and g > 250 and b > 250:  # Blanc ou presque blanc
                        continue
                    
                    # Vérifier si la cellule fait partie d'une fusion
                    merge_data = merged_info.get((cell.row, cell.column), {})
                    
                    cell_info = {
                        'row': cell.row,
                        'col': cell.column,
                        'value': get_cell_value(cell),
                        'address': f"{num_to_excel_col(cell.column)}{cell.row}",
                        'color': hex_color,
                        'is_merged': bool(merge_data),
                        'merge_info': merge_data
                    }
                    
                    color_cells[hex_color].append(cell_info)
                    color_counts[hex_color] += 1
    
    print(f"Debug - Cellules analysées: {total_cells}, avec fill: {cells_with_fill}, avec couleur: {sum(color_counts.values())}")
    print(f"Debug - Couleurs trouvées: {list(color_counts.keys())}")
                    
    # Créer un résumé des couleurs
    color_summary = []
    for hex_color, count in color_counts.most_common():
        # Exemples avec indication des cellules fusionnées
        examples = []
        merged_count = 0
        for cell in color_cells[hex_color][:5]:
            addr = cell['address']
            if cell.get('is_merged'):
                merged_count += 1
                merge_info = cell.get('merge_info', {})
                if merge_info.get('is_merged_cell'):
                    continue  # Ne pas montrer les cellules secondaires d'une fusion
                if 'span_rows' in merge_info:
                    addr += f" (fusionnée: {merge_info['span_rows']}x{merge_info['span_cols']})"
            examples.append(addr)
        
        color_summary.append({
            'hex': hex_color,
            'name': get_color_name(hex_color),
            'count': count,
            'cells': color_cells[hex_color][:5],
            'examples': examples,
            'merged_count': merged_count,
            'rgb': hex_to_rgb(hex_color)
        })
    
    # Si peu de couleurs détectées, essayer de grouper les couleurs similaires
    if len(color_summary) < 3:
        print("Debug - Peu de couleurs détectées, tentative de regroupement des couleurs similaires")
        color_summary = group_similar_colors(color_summary, tolerance=0.15)
    
    return color_summary, color_cells

def is_similar_color(color1: str, color2: str, tolerance: float = 0.1) -> bool:
    """
    Vérifie si deux couleurs sont similaires
    tolerance: différence maximale acceptée en HSV
    """
    try:
        r1, g1, b1 = hex_to_rgb(color1)
        r2, g2, b2 = hex_to_rgb(color2)
        
        h1, s1, v1 = colorsys.rgb_to_hsv(r1/255, g1/255, b1/255)
        h2, s2, v2 = colorsys.rgb_to_hsv(r2/255, g2/255, b2/255)
        
        # Calculer la différence
        h_diff = min(abs(h1 - h2), 1 - abs(h1 - h2))  # Gérer la circularité
        s_diff = abs(s1 - s2)
        v_diff = abs(v1 - v2)
        
        return h_diff < tolerance and s_diff < tolerance and v_diff < tolerance
    except:
        return False

def group_similar_colors(color_summary: List[Dict], tolerance: float = 0.1) -> List[Dict]:
    """
    Groupe les couleurs similaires ensemble
    Utile pour gérer les variations mineures de couleur
    """
    grouped = []
    used = set()
    
    for i, color in enumerate(color_summary):
        if i in used:
            continue
            
        group = {
            'hex': color['hex'],
            'name': color['name'],
            'count': color['count'],
            'cells': color['cells'][:],
            'variations': [color['hex']]
        }
        
        # Chercher les couleurs similaires
        for j, other in enumerate(color_summary[i+1:], i+1):
            if j not in used and is_similar_color(color['hex'], other['hex'], tolerance):
                group['count'] += other['count']
                group['cells'].extend(other['cells'][:2])  # Ajouter quelques exemples
                group['variations'].append(other['hex'])
                used.add(j)
        
        grouped.append(group)
    
    return grouped