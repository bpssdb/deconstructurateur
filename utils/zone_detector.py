"""
Module de détection et de groupement des zones dans les fichiers Excel
Support des paires de labels alternées
"""

from typing import List, Dict, Set, Tuple, Optional
from collections import defaultdict

def detect_zones_with_alternating_pairs(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """
    Détecte les zones et labels avec système de paires alternées
    
    color_palette format:
    {
        'zone_color': 'RRGGBB',
        'zone_name': 'Zones de données',
        'label_pairs': [
            {
                'horizontal': {'color': 'RRGGBB', 'name': 'Headers H1'},
                'vertical': {'color': 'RRGGBB', 'name': 'Headers V1'}
            },
            {
                'horizontal': {'color': 'RRGGBB', 'name': 'Headers H2'},
                'vertical': {'color': 'RRGGBB', 'name': 'Headers V2'}
            }
        ]
    }
    """
    # Récupérer les cellules de zones
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    
    # Récupérer toutes les cellules de labels par paire
    label_cells_by_pair = []
    all_label_cells = []
    
    for pair_idx, pair in enumerate(color_palette.get('label_pairs', [])):
        h_cells = color_cells.get(pair['horizontal']['color'], [])
        v_cells = color_cells.get(pair['vertical']['color'], [])
        
        label_cells_by_pair.append({
            'pair_id': pair_idx,
            'horizontal': h_cells,
            'vertical': v_cells,
            'h_color': pair['horizontal']['color'],
            'v_color': pair['vertical']['color']
        })
        
        all_label_cells.extend(h_cells)
        all_label_cells.extend(v_cells)
    
    # Grouper les cellules de zone en zones contiguës
    zones = group_contiguous_cells(zone_cells)
    
    # Associer les labels aux zones avec logique d'alternance
    for zone in zones:
        zone['labels'] = find_labels_with_alternating_pairs(
            zone, 
            label_cells_by_pair,
            color_palette
        )
    
    return zones, label_cells_by_pair

def find_labels_with_alternating_pairs(zone: Dict, label_pairs: List[Dict], color_palette: Dict) -> List[Dict]:
    """
    Trouve les labels pour une zone en utilisant la logique des paires alternées
    Version avec debug amélioré
    """
    labels = []
    
    # Créer un mapping position -> label pour accès rapide
    label_map = {}
    color_to_pair_and_type = {}
    
    print(f"\nDEBUG find_labels: Zone {zone['id']} - bounds: rows {zone['bounds']['min_row']}-{zone['bounds']['max_row']}, cols {zone['bounds']['min_col']}-{zone['bounds']['max_col']}")
    print(f"DEBUG: {len(label_pairs)} paires de labels à traiter")
    
    for pair_data in label_pairs:
        pair_id = pair_data['pair_id']
        
        # Labels horizontaux
        print(f"DEBUG: Paire {pair_id} - {len(pair_data['horizontal'])} labels horizontaux")
        for label in pair_data['horizontal']:
            label_map[(label['row'], label['col'])] = {
                **label,
                'pair_id': pair_id,
                'direction': 'horizontal',
                'color': pair_data['h_color']
            }
            color_to_pair_and_type[pair_data['h_color']] = (pair_id, 'horizontal')
        
        # Labels verticaux
        print(f"DEBUG: Paire {pair_id} - {len(pair_data['vertical'])} labels verticaux")
        for label in pair_data['vertical']:
            label_map[(label['row'], label['col'])] = {
                **label,
                'pair_id': pair_id,
                'direction': 'vertical',
                'color': pair_data['v_color']
            }
            color_to_pair_and_type[pair_data['v_color']] = (pair_id, 'vertical')
    
    print(f"DEBUG: Total de {len(label_map)} labels dans label_map")
    
    # Pour chaque cellule de la zone, chercher ses labels en remontant
    processed = set()
    cells_checked = 0
    
    for zone_cell in zone['cells']:
        zone_row = zone_cell['row']
        zone_col = zone_cell['col']
        cells_checked += 1
        
        if cells_checked <= 3:  # Debug pour les premières cellules
            print(f"DEBUG: Recherche labels pour cellule {zone_row},{zone_col}")
        
        # Chercher les labels horizontaux (remonter dans la colonne)
        horizontal_labels = find_horizontal_labels_alternating(
            zone_row, zone_col, label_map, color_to_pair_and_type
        )
        
        # Chercher les labels verticaux (reculer dans la ligne)
        vertical_labels = find_vertical_labels_alternating(
            zone_row, zone_col, label_map, color_to_pair_and_type
        )
        
        if cells_checked <= 3 and (horizontal_labels or vertical_labels):
            print(f"  → Trouvé {len(horizontal_labels)} H, {len(vertical_labels)} V")
        
        # Ajouter les labels trouvés (éviter les doublons)
        for label in horizontal_labels + vertical_labels:
            key = (label['row'], label['col'], label['direction'], label.get('pair_id'))
            if key not in processed:
                # Ajouter le nom de la paire depuis la palette
                if label.get('pair_id') is not None and label['pair_id'] < len(color_palette.get('label_pairs', [])):
                    pair = color_palette['label_pairs'][label['pair_id']]
                    if label['direction'] == 'horizontal':
                        label['pair_name'] = pair['horizontal'].get('name', f'H{label["pair_id"]+1}')
                    else:
                        label['pair_name'] = pair['vertical'].get('name', f'V{label["pair_id"]+1}')
                labels.append(label)
                processed.add(key)
    
    print(f"DEBUG: Zone {zone['id']} - {cells_checked} cellules vérifiées, {len(labels)} labels trouvés")
    
    return labels

def find_horizontal_labels_alternating(row: int, col: int, label_map: Dict, color_to_pair: Dict) -> List[Dict]:
    """
    Remonte dans une colonne pour trouver tous les labels horizontaux
    S'arrête quand on rencontre un label de l'autre type de la même paire
    Version avec debug
    """
    labels = []
    current_pair_id = None
    labels_checked = 0
    
    # Remonter dans la colonne
    for check_row in range(row - 1, 0, -1):
        labels_checked += 1
        if labels_checked <= 5 and (check_row, col) in label_map:
            print(f"    DEBUG H: Vérif ({check_row},{col}) - label trouvé")
        
        if (check_row, col) in label_map:
            label = label_map[(check_row, col)]
            
            # Si c'est un label horizontal
            if label['direction'] == 'horizontal':
                # Si c'est la première fois ou si c'est la même paire
                if current_pair_id is None or label['pair_id'] == current_pair_id:
                    labels.append({
                        'row': label['row'],
                        'col': label['col'],
                        'value': label.get('value', ''),
                        'type': f"h_pair_{label['pair_id']}",
                        'position': 'top',
                        'direction': 'horizontal',
                        'pair_id': label['pair_id'],
                        'distance': row - check_row,
                        'color': label['color']
                    })
                    current_pair_id = label['pair_id']
                    if labels_checked <= 5:
                        print(f"      → Ajouté label H paire {current_pair_id}")
                else:
                    # Changement de paire, on continue
                    labels.append({
                        'row': label['row'],
                        'col': label['col'],
                        'value': label.get('value', ''),
                        'type': f"h_pair_{label['pair_id']}",
                        'position': 'top',
                        'direction': 'horizontal',
                        'pair_id': label['pair_id'],
                        'distance': row - check_row,
                        'color': label['color']
                    })
                    current_pair_id = label['pair_id']
            
            # Si c'est un label vertical de la même paire, on s'arrête
            elif label['direction'] == 'vertical' and label['pair_id'] == current_pair_id:
                if labels_checked <= 5:
                    print(f"      → Stop: label V de la même paire {current_pair_id}")
                break
    
    return labels

def find_vertical_labels_alternating(row: int, col: int, label_map: Dict, color_to_pair: Dict) -> List[Dict]:
    """
    Recule dans une ligne pour trouver tous les labels verticaux
    S'arrête quand on rencontre un label de l'autre type de la même paire
    """
    labels = []
    current_pair_id = None
    
    # Reculer dans la ligne
    for check_col in range(col - 1, 0, -1):
        if (row, check_col) in label_map:
            label = label_map[(row, check_col)]
            
            # Si c'est un label vertical
            if label['direction'] == 'vertical':
                # Si c'est la première fois ou si c'est la même paire
                if current_pair_id is None or label['pair_id'] == current_pair_id:
                    labels.append({
                        'row': label['row'],
                        'col': label['col'],
                        'value': label.get('value', ''),
                        'type': f"v_pair_{label['pair_id']}",
                        'position': 'left',
                        'direction': 'vertical',
                        'pair_id': label['pair_id'],
                        'distance': col - check_col,
                        'color': label['color']
                    })
                    current_pair_id = label['pair_id']
                else:
                    # Changement de paire, on continue
                    labels.append({
                        'row': label['row'],
                        'col': label['col'],
                        'value': label.get('value', ''),
                        'type': f"v_pair_{label['pair_id']}",
                        'position': 'left',
                        'direction': 'vertical',
                        'pair_id': label['pair_id'],
                        'distance': col - check_col,
                        'color': label['color']
                    })
                    current_pair_id = label['pair_id']
            
            # Si c'est un label horizontal de la même paire, on s'arrête
            elif label['direction'] == 'horizontal' and label['pair_id'] == current_pair_id:
                break
    
    return labels

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
            'labels': zone1.get('labels', [])
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

# Garder les anciennes fonctions pour compatibilité
def detect_zones_with_palette(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """Version compatible avec l'ancienne API (3 couleurs)"""
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    label1_cells = color_cells.get(color_palette.get('label1_color'), [])
    label2_cells = color_cells.get(color_palette.get('label2_color'), [])
    
    zones = group_contiguous_cells(zone_cells)
    
    for zone in zones:
        zone['labels'] = find_labels_for_zone(
            zone, 
            label1_cells + label2_cells, 
            color_palette.get('label1_color'), 
            color_palette.get('label2_color')
        )
    
    return zones, {'label1': label1_cells, 'label2': label2_cells}

def detect_zones_with_flexible_palette(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """Version avec support pour un nombre variable de couleurs de labels"""
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    
    all_label_cells = []
    label_cells_by_type = {}
    
    if 'label_colors' in color_palette:
        for label_type, label_info in color_palette['label_colors'].items():
            cells = color_cells.get(label_info['color'], [])
            all_label_cells.extend(cells)
            label_cells_by_type[label_type] = cells
    
    zones = group_contiguous_cells(zone_cells)
    
    for zone in zones:
        zone['labels'] = find_labels_for_zone_flexible(
            zone, 
            all_label_cells,
            color_palette
        )
    
    return zones, label_cells_by_type

def find_labels_for_zone(zone: Dict, label_cells: List[Dict], label1_color: str, label2_color: str) -> List[Dict]:
    """Trouve les labels associés à une zone (ancienne version)"""
    labels = []
    bounds = zone['bounds']
    
    for label in label_cells:
        if (label['row'] == bounds['min_row'] - 1 and 
            bounds['min_col'] <= label['col'] <= bounds['max_col']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'top',
                'type': label_type
            })
        elif (label['col'] == bounds['min_col'] - 1 and 
              bounds['min_row'] <= label['row'] <= bounds['max_row']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'left',
                'type': label_type
            })
    
    return labels

def find_labels_for_zone_flexible(zone: Dict, label_cells: List[Dict], color_palette: Dict) -> List[Dict]:
    """Trouve les labels associés à une zone (version flexible)"""
    labels = []
    bounds = zone['bounds']
    
    color_to_type = {}
    if 'label_colors' in color_palette:
        for label_type, label_info in color_palette['label_colors'].items():
            color_to_type[label_info['color']] = label_type
    
    label_map = {}
    for label in label_cells:
        label_map[(label['row'], label['col'])] = label
    
    processed_labels = set()
    
    for zone_cell in zone['cells']:
        zone_row = zone_cell['row']
        zone_col = zone_cell['col']
        
        for check_row in range(zone_row - 1, 0, -1):
            if (check_row, zone_col) in label_map:
                label = label_map[(check_row, zone_col)]
                label_key = (label['row'], label['col'], 'horizontal')
                
                if label_key not in processed_labels:
                    label_color = label.get('color', '')
                    label_type = color_to_type.get(label_color, 'unknown')
                    
                    labels.append({
                        **label,
                        'position': 'top',
                        'type': label_type,
                        'for_cells': [(zone_row, zone_col)],
                        'distance': zone_row - check_row
                    })
                    processed_labels.add(label_key)
                break
        
        for check_col in range(zone_col - 1, 0, -1):
            if (zone_row, check_col) in label_map:
                label = label_map[(zone_row, check_col)]
                label_key = (label['row'], label['col'], 'vertical')
                
                if label_key not in processed_labels:
                    label_color = label.get('color', '')
                    label_type = color_to_type.get(label_color, 'unknown')
                    
                    labels.append({
                        **label,
                        'position': 'left',
                        'type': label_type,
                        'for_cells': [(zone_row, zone_col)],
                        'distance': zone_col - check_col
                    })
                    processed_labels.add(label_key)
                break
    
    consolidated_labels = {}
    for label in labels:
        key = (label['row'], label['col'], label['type'], label['position'])
        if key in consolidated_labels:
            consolidated_labels[key]['for_cells'].extend(label['for_cells'])
        else:
            consolidated_labels[key] = label
    
    return list(consolidated_labels.values())