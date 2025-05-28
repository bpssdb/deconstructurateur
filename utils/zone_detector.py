"""
Module de détection et de groupement des zones dans les fichiers Excel
"""

from typing import List, Dict, Set, Tuple

def detect_zones_with_palette(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """
    Détecte les zones et labels basés sur la palette définie par l'utilisateur
    Version compatible avec l'ancienne API (3 couleurs)
    """
    # Récupérer les cellules pour chaque type
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    label1_cells = color_cells.get(color_palette.get('label1_color'), [])
    label2_cells = color_cells.get(color_palette.get('label2_color'), [])
    
    # Grouper les cellules de zone en zones contiguës
    zones = group_contiguous_cells(zone_cells)
    
    # Associer les labels aux zones
    for zone in zones:
        zone['labels'] = find_labels_for_zone(
            zone, 
            label1_cells + label2_cells, 
            color_palette.get('label1_color'), 
            color_palette.get('label2_color')
        )
    
    return zones, {'label1': label1_cells, 'label2': label2_cells}

def detect_zones_with_flexible_palette(workbook, sheet_name: str, color_palette: Dict, color_cells: Dict) -> Tuple[List[Dict], Dict]:
    """
    Détecte les zones et labels avec support pour un nombre variable de couleurs de labels
    """
    # Récupérer les cellules de zones
    zone_cells = color_cells.get(color_palette['zone_color'], [])
    
    # Récupérer toutes les cellules de labels
    all_label_cells = []
    label_cells_by_type = {}
    
    if 'label_colors' in color_palette:
        for label_type, label_info in color_palette['label_colors'].items():
            cells = color_cells.get(label_info['color'], [])
            all_label_cells.extend(cells)
            label_cells_by_type[label_type] = cells
    
    # Grouper les cellules de zone en zones contiguës
    zones = group_contiguous_cells(zone_cells)
    
    # Associer les labels aux zones
    for zone in zones:
        zone['labels'] = find_labels_for_zone_flexible(
            zone, 
            all_label_cells,
            color_palette
        )
    
    return zones, label_cells_by_type

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

def find_labels_for_zone(zone: Dict, label_cells: List[Dict], label1_color: str, label2_color: str) -> List[Dict]:
    """
    Trouve les labels associés à une zone
    Cherche les cellules colorées immédiatement adjacentes à la zone
    """
    labels = []
    bounds = zone['bounds']
    
    for label in label_cells:
        # Chercher les labels au-dessus de la zone
        if (label['row'] == bounds['min_row'] - 1 and 
            bounds['min_col'] <= label['col'] <= bounds['max_col']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'top',
                'type': label_type
            })
        
        # Chercher les labels à gauche de la zone
        elif (label['col'] == bounds['min_col'] - 1 and 
              bounds['min_row'] <= label['row'] <= bounds['max_row']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'left',
                'type': label_type
            })
        
        # Chercher les labels en bas de la zone
        elif (label['row'] == bounds['max_row'] + 1 and 
              bounds['min_col'] <= label['col'] <= bounds['max_col']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'bottom',
                'type': label_type
            })
        
        # Chercher les labels à droite de la zone
        elif (label['col'] == bounds['max_col'] + 1 and 
              bounds['min_row'] <= label['row'] <= bounds['max_row']):
            label_type = 'label1' if label.get('color', '') == label1_color else 'label2'
            labels.append({
                **label,
                'position': 'right',
                'type': label_type
            })
    
    return labels

def find_labels_for_zone_flexible(zone: Dict, label_cells: List[Dict], color_palette: Dict) -> List[Dict]:
    """
    Trouve les labels associés à une zone
    Remonte/recule jusqu'à trouver une cellule de couleur label
    """
    labels = []
    bounds = zone['bounds']
    
    # Créer un mapping couleur -> type de label
    color_to_type = {}
    if 'label_colors' in color_palette:
        for label_type, label_info in color_palette['label_colors'].items():
            color_to_type[label_info['color']] = label_type
    
    # Créer un dictionnaire pour accès rapide aux labels par position
    label_map = {}
    for label in label_cells:
        label_map[(label['row'], label['col'])] = label
    
    # Pour chaque cellule de la zone, chercher ses labels
    processed_labels = set()
    
    for zone_cell in zone['cells']:
        zone_row = zone_cell['row']
        zone_col = zone_cell['col']
        
        # Chercher le label horizontal (remonter dans la même colonne)
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
                break  # Arrêter dès qu'on trouve un label
        
        # Chercher le label vertical (reculer dans la même ligne)
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
                break  # Arrêter dès qu'on trouve un label
    
    # Consolider les labels qui s'appliquent à plusieurs cellules
    consolidated_labels = {}
    for label in labels:
        key = (label['row'], label['col'], label['type'], label['position'])
        if key in consolidated_labels:
            consolidated_labels[key]['for_cells'].extend(label['for_cells'])
        else:
            consolidated_labels[key] = label
    
    return list(consolidated_labels.values())
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