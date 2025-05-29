"""
Module de détection et de groupement des zones dans les fichiers Excel
Support des 4 couleurs de headers indépendantes (2 horizontaux, 2 verticaux)
CORRECTION: V = vertical = au-dessus, H = horizontal = à gauche
"""

from typing import List, Dict, Set, Tuple, Optional
from collections import defaultdict

def find_labels_for_zone_with_colors(zone: Dict, label_data: Dict) -> List[Dict]:
    """
    Trouve les labels pour une zone selon la logique des 4 couleurs:
    - VERTICAL (V1/V2): on REMONTE dans la COLONNE (chercher au-dessus)
    - HORIZONTAL (H1/H2): on RECULE dans la LIGNE (chercher à gauche)
    """
    labels = []
    processed = set()
    
    # Créer des mappings position -> label pour un accès rapide
    h_positions = {}  # Headers horizontaux (à gauche)
    v_positions = {}  # Headers verticaux (au-dessus)
    
    # Mapper H1 (horizontal) - CORRECTION: créer une copie propre du dictionnaire
    for cell in label_data.get('h1', {}).get('cells', []):
        if isinstance(cell, dict) and 'row' in cell and 'col' in cell:
            h_positions[(cell['row'], cell['col'])] = {
                'row': cell['row'],
                'col': cell['col'],
                'value': cell.get('value', ''),
                'type': 'h1',
                'color': label_data['h1']['color'],
                'direction': 'horizontal'
            }
    
    # Mapper H2 (horizontal)
    for cell in label_data.get('h2', {}).get('cells', []):
        if isinstance(cell, dict) and 'row' in cell and 'col' in cell:
            h_positions[(cell['row'], cell['col'])] = {
                'row': cell['row'],
                'col': cell['col'],
                'value': cell.get('value', ''),
                'type': 'h2',
                'color': label_data['h2']['color'],
                'direction': 'horizontal'
            }
    
    # Mapper V1 (vertical)
    for cell in label_data.get('v1', {}).get('cells', []):
        if isinstance(cell, dict) and 'row' in cell and 'col' in cell:
            v_positions[(cell['row'], cell['col'])] = {
                'row': cell['row'],
                'col': cell['col'],
                'value': cell.get('value', ''),
                'type': 'v1',
                'color': label_data['v1']['color'],
                'direction': 'vertical'
            }
    
    # Mapper V2 (vertical)
    for cell in label_data.get('v2', {}).get('cells', []):
        if isinstance(cell, dict) and 'row' in cell and 'col' in cell:
            v_positions[(cell['row'], cell['col'])] = {
                'row': cell['row'],
                'col': cell['col'],
                'value': cell.get('value', ''),
                'type': 'v2',
                'color': label_data['v2']['color'],
                'direction': 'vertical'
            }
    
    print(f"\n  DEBUG find_labels: Zone {zone['id']}")
    print(f"    - Zone bounds: rows {zone['bounds']['min_row']}-{zone['bounds']['max_row']}, cols {zone['bounds']['min_col']}-{zone['bounds']['max_col']}")
    print(f"    - H positions mapped: {len(h_positions)} (horizontal - à gauche)")
    print(f"    - V positions mapped: {len(v_positions)} (vertical - au-dessus)")
    
    # Pour chaque cellule de la zone
    for zone_cell in zone['cells']:
        zone_row = zone_cell['row']
        zone_col = zone_cell['col']
        
        # 1. Chercher les headers VERTICAUX (remonter dans la COLONNE - au-dessus)
        first_v_color = None
        
        for check_row in range(zone_row - 1, 0, -1):  # Remonter
            if (check_row, zone_col) in v_positions:  # Même colonne
                v_label = v_positions[(check_row, zone_col)]
                current_color = v_label['color']
                
                # Si c'est le premier header V trouvé, on note sa couleur
                if first_v_color is None:
                    first_v_color = current_color
                
                # Si c'est la même couleur que le premier trouvé, on l'ajoute
                if current_color == first_v_color:
                    key = (v_label['row'], v_label['col'], 'vertical', v_label['type'])
                    if key not in processed:
                        labels.append({
                            'row': v_label['row'],
                            'col': v_label['col'],
                            'value': str(v_label.get('value', '')),  # CORRECTION: forcer en string
                            'type': v_label['type'],
                            'position': 'top',
                            'direction': 'vertical',
                            'distance': zone_row - check_row,
                            'color': v_label['color']
                        })
                        processed.add(key)
                # Si c'est une couleur V différente, on arrête
                else:
                    break
        
        # 2. Chercher les headers HORIZONTAUX (reculer dans la LIGNE - à gauche)
        first_h_color = None
        
        for check_col in range(zone_col - 1, 0, -1):  # Reculer
            if (zone_row, check_col) in h_positions:  # Même ligne
                h_label = h_positions[(zone_row, check_col)]
                current_color = h_label['color']
                
                # Si c'est le premier header H trouvé, on note sa couleur
                if first_h_color is None:
                    first_h_color = current_color
                
                # Si c'est la même couleur que le premier trouvé, on l'ajoute
                if current_color == first_h_color:
                    key = (h_label['row'], h_label['col'], 'horizontal', h_label['type'])
                    if key not in processed:
                        labels.append({
                            'row': h_label['row'],
                            'col': h_label['col'],
                            'value': str(h_label.get('value', '')),  # CORRECTION: forcer en string
                            'type': h_label['type'],
                            'position': 'left',
                            'direction': 'horizontal',
                            'distance': zone_col - check_col,
                            'color': h_label['color']
                        })
                        processed.add(key)
                # Si c'est une couleur H différente, on arrête
                else:
                    break
    
    print(f"\n  Total labels found for zone {zone['id']}: {len(labels)}")
    return labels


def detect_zones_with_two_colors(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """
    Détecte les zones avec le système à 4 couleurs (zone + 2H + 2V)
    
    color_palette format attendu:
    {
        'zone_color': 'RRGGBB',
        'h1_color': 'RRGGBB',  # Horizontal 1 (à gauche)
        'h2_color': 'RRGGBB',  # Horizontal 2 (à gauche)
        'v1_color': 'RRGGBB',  # Vertical 1 (au-dessus)
        'v2_color': 'RRGGBB',  # Vertical 2 (au-dessus)
        ... (et les noms correspondants)
    }
    """
    # Récupérer les cellules de zones
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    
    # Récupérer les cellules de headers
    h1_color = color_palette.get('h1_color')
    h2_color = color_palette.get('h2_color')
    v1_color = color_palette.get('v1_color')
    v2_color = color_palette.get('v2_color')
    
    h1_cells = color_cells.get(h1_color, []) if h1_color else []
    h2_cells = color_cells.get(h2_color, []) if h2_color else []
    v1_cells = color_cells.get(v1_color, []) if v1_color else []
    v2_cells = color_cells.get(v2_color, []) if v2_color else []
    
    print(f"DEBUG detect_zones_with_two_colors:")
    print(f"  - Zone cells: {len(zone_cells)}")
    print(f"  - H1 cells ({h1_color}): {len(h1_cells)} (horizontal - à gauche)")
    print(f"  - H2 cells ({h2_color}): {len(h2_cells)} (horizontal - à gauche)")
    print(f"  - V1 cells ({v1_color}): {len(v1_cells)} (vertical - au-dessus)")
    print(f"  - V2 cells ({v2_color}): {len(v2_cells)} (vertical - au-dessus)")
    
    # Grouper les cellules de zone en zones contiguës
    zones = group_contiguous_cells(zone_cells)
    
    # Préparer les données de labels
    label_data = {
        'h1': {'cells': h1_cells, 'color': h1_color},
        'h2': {'cells': h2_cells, 'color': h2_color},
        'v1': {'cells': v1_cells, 'color': v1_color},
        'v2': {'cells': v2_cells, 'color': v2_color}
    }
    
    # Associer les labels aux zones et ajouter le nom de la feuille
    for zone in zones:
        zone['labels'] = find_labels_for_zone_with_colors(zone, label_data)
        
        # Ajouter le nom de la feuille comme label
        zone['sheet_name'] = sheet_name
        
        print(f"  Zone {zone['id']}: {len(zone['labels'])} labels trouvés")
    
    return zones, label_data

def group_contiguous_cells(cells: List[Dict]) -> List[Dict]:
    """
    Groupe les cellules contiguës en zones
    Utilise un algorithme DFS pour trouver les composantes connexes
    """
    if not cells:
        return []
    
    zones = []
    visited = set()
    
    def get_neighbors(cell: Dict, all_cells: List[Dict]) -> List[Dict]:
        """Trouve les voisins d'une cellule"""
        neighbors = []
        for c in all_cells:
            if (c['row'], c['col']) in visited:
                continue
            # Cellules adjacentes (horizontalement ou verticalement)
            if (abs(c['row'] - cell['row']) == 1 and c['col'] == cell['col']) or \
               (abs(c['col'] - cell['col']) == 1 and c['row'] == cell['row']):
                neighbors.append(c)
        return neighbors
    
    # DFS pour trouver les zones contiguës
    for i, cell in enumerate(cells):
        cell_key = (cell['row'], cell['col'])
        if cell_key not in visited:
            zone_cells = []
            stack = [cell]
            
            while stack:
                current = stack.pop()
                if (current['row'], current['col']) in visited:
                    continue
                
                visited.add((current['row'], current['col']))
                zone_cells.append(current)
                stack.extend(get_neighbors(current, cells))
            
            if zone_cells:
                # Calculer les limites de la zone
                min_row = min(c['row'] for c in zone_cells)
                max_row = max(c['row'] for c in zone_cells)
                min_col = min(c['col'] for c in zone_cells)
                max_col = max(c['col'] for c in zone_cells)
                
                zones.append({
                    'id': len(zones) + 1,
                    'cells': zone_cells,
                    'bounds': {
                        'min_row': min_row,
                        'max_row': max_row,
                        'min_col': min_col,
                        'max_col': max_col
                    },
                    'cell_count': len(zone_cells)
                })
    
    return zones

def merge_zones(zones: List[Dict], max_gap: int = 1) -> List[Dict]:
    """
    Fusionne les zones proches (avec un écart maximum)
    Utile pour gérer les zones avec des espaces mineurs
    """
    if len(zones) <= 1:
        return zones
    
    merged = []
    used = set()
    
    for i, zone1 in enumerate(zones):
        if i in used:
            continue
            
        merged_zone = {
            'id': len(merged) + 1,
            'cells': zone1['cells'][:],
            'bounds': zone1['bounds'].copy(),
            'labels': zone1.get('labels', []),
            'sheet_name': zone1.get('sheet_name', '')  # Conserver le nom de la feuille
        }
        
        # Chercher les zones à fusionner
        for j, zone2 in enumerate(zones[i+1:], i+1):
            if j in used:
                continue
                
            # Vérifier si les zones sont proches
            if are_zones_adjacent(zone1['bounds'], zone2['bounds'], max_gap):
                # Fusionner les zones
                merged_zone['cells'].extend(zone2['cells'])
                merged_zone['labels'].extend(zone2.get('labels', []))
                
                # Mettre à jour les limites
                merged_zone['bounds']['min_row'] = min(
                    merged_zone['bounds']['min_row'], 
                    zone2['bounds']['min_row']
                )
                merged_zone['bounds']['max_row'] = max(
                    merged_zone['bounds']['max_row'], 
                    zone2['bounds']['max_row']
                )
                merged_zone['bounds']['min_col'] = min(
                    merged_zone['bounds']['min_col'], 
                    zone2['bounds']['min_col']
                )
                merged_zone['bounds']['max_col'] = max(
                    merged_zone['bounds']['max_col'], 
                    zone2['bounds']['max_col']
                )
                
                used.add(j)
        
        merged_zone['cell_count'] = len(merged_zone['cells'])
        merged.append(merged_zone)
    
    return merged

def are_zones_adjacent(bounds1: Dict, bounds2: Dict, max_gap: int = 1) -> bool:
    """
    Vérifie si deux zones sont adjacentes ou proches
    """
    # Vérifier l'adjacence horizontale
    if (bounds1['min_row'] <= bounds2['max_row'] + max_gap and 
        bounds1['max_row'] >= bounds2['min_row'] - max_gap):
        # Vérifier si elles sont alignées verticalement
        if (abs(bounds1['max_col'] - bounds2['min_col']) <= max_gap or
            abs(bounds2['max_col'] - bounds1['min_col']) <= max_gap):
            return True
    
    # Vérifier l'adjacence verticale
    if (bounds1['min_col'] <= bounds2['max_col'] + max_gap and 
        bounds1['max_col'] >= bounds2['min_col'] - max_gap):
        # Vérifier si elles sont alignées horizontalement
        if (abs(bounds1['max_row'] - bounds2['min_row']) <= max_gap or
            abs(bounds2['max_row'] - bounds1['min_row']) <= max_gap):
            return True
    
    return False

# Garder les fonctions de compatibilité pour ne pas casser le code existant
def detect_zones_with_alternating_pairs(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """
    Fonction de compatibilité qui convertit l'ancien format "pairs" vers le nouveau format
    """
    # Si on a le format avec label_pairs, convertir vers le format direct
    if 'label_pairs' in color_palette and len(color_palette['label_pairs']) >= 2:
        new_palette = {
            'zone_color': color_palette['zone_color'],
            'zone_name': color_palette['zone_name'],
            'h1_color': color_palette['label_pairs'][0]['horizontal']['color'],
            'h1_name': color_palette['label_pairs'][0]['horizontal']['name'],
            'h2_color': color_palette['label_pairs'][1]['horizontal']['color'],
            'h2_name': color_palette['label_pairs'][1]['horizontal']['name'],
            'v1_color': color_palette['label_pairs'][0]['vertical']['color'],
            'v1_name': color_palette['label_pairs'][0]['vertical']['name'],
            'v2_color': color_palette['label_pairs'][1]['vertical']['color'],
            'v2_name': color_palette['label_pairs'][1]['vertical']['name']
        }
        return detect_zones_with_two_colors(workbook, sheet_name, new_palette, color_cells)
    else:
        # Utiliser directement si déjà au bon format
        return detect_zones_with_two_colors(workbook, sheet_name, color_palette, color_cells)