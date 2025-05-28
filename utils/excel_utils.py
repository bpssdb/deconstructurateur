"""
Utilitaires pour la manipulation des fichiers Excel
Supporte les formats .xlsx et .xls
"""

import openpyxl
import xlrd
import tempfile
import os
from typing import Union, List, Tuple, Any, Dict

def num_to_excel_col(n: int) -> str:
    """Convertit un numéro de colonne en lettre Excel"""
    if n <= 0:
        return "?"
    col = ""
    while n > 0:
        n -= 1
        col = chr(ord('A') + n % 26) + col
        n //= 26
    return col

def excel_col_to_num(col_str: str) -> int:
    """Convertit une lettre de colonne Excel en numéro"""
    num = 0
    for char in col_str:
        num = num * 26 + (ord(char.upper()) - ord('A') + 1)
    return num

def load_workbook(file):
    """
    Charge un fichier Excel (.xlsx ou .xls)
    Retourne un workbook openpyxl
    """
    # Déterminer le type de fichier
    filename = file.name.lower()
    
    if filename.endswith('.xlsx'):
        # Fichier .xlsx - utiliser openpyxl directement
        return openpyxl.load_workbook(file, data_only=False)
    
    elif filename.endswith('.xls'):
        # Fichier .xls - convertir via xlrd
        return convert_xls_to_openpyxl(file)
    
    else:
        raise ValueError("Format de fichier non supporté. Utilisez .xlsx ou .xls")

def convert_xls_to_openpyxl(file):
    """
    Convertit un fichier .xls en workbook openpyxl
    Préserve les couleurs et le formatage
    """
    # Lire le fichier .xls avec xlrd
    xls_book = xlrd.open_workbook(file_contents=file.read(), formatting_info=True)
    
    # Créer un nouveau workbook openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Supprimer la feuille par défaut
    
    # Obtenir les informations de formatage
    xf_list = xls_book.format_map
    
    # Parcourir toutes les feuilles
    for sheet_idx, sheet_name in enumerate(xls_book.sheet_names()):
        xls_sheet = xls_book.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        
        # Copier les données et le formatage
        for row_idx in range(xls_sheet.nrows):
            for col_idx in range(xls_sheet.ncols):
                cell = xls_sheet.cell(row_idx, col_idx)
                
                # Écrire la valeur
                ws.cell(row=row_idx + 1, column=col_idx + 1, value=cell.value)
                
                # Appliquer le formatage si disponible
                if cell.xf_index is not None and cell.xf_index < len(xf_list):
                    xf = xf_list[cell.xf_index]
                    
                    # Extraire la couleur de fond si elle existe
                    if xf.background and hasattr(xf.background, 'pattern_colour_index'):
                        color_idx = xf.background.pattern_colour_index
                        if color_idx and color_idx < len(xls_book.colour_map):
                            rgb = xls_book.colour_map.get(color_idx)
                            if rgb:
                                # Convertir RGB en hex
                                hex_color = '%02x%02x%02x' % rgb[:3]
                                from openpyxl.styles import PatternFill
                                fill = PatternFill(start_color=hex_color, 
                                                 end_color=hex_color, 
                                                 fill_type="solid")
                                ws.cell(row=row_idx + 1, column=col_idx + 1).fill = fill
    
    return wb

def get_sheet_names(workbook) -> List[str]:
    """Retourne la liste des noms de feuilles"""
    return workbook.sheetnames

def get_cell_color(cell) -> Union[str, None]:
    """
    Extrait la couleur d'une cellule
    Retourne le code hex ou None
    """
    try:
        # Vérifier si la cellule a un remplissage
        if not hasattr(cell, 'fill') or not cell.fill:
            return None
        
        # Vérifier le type de remplissage
        if hasattr(cell.fill, 'patternType'):
            # Si pas de pattern ou pattern none, pas de couleur
            if cell.fill.patternType in [None, 'none']:
                return None
        
        color = None
        
        # Essayer différentes propriétés de couleur
        # 1. fgColor (couleur de premier plan)
        if hasattr(cell.fill, 'fgColor') and cell.fill.fgColor:
            if hasattr(cell.fill.fgColor, 'rgb') and cell.fill.fgColor.rgb:
                color = cell.fill.fgColor.rgb
            elif hasattr(cell.fill.fgColor, 'value') and cell.fill.fgColor.value:
                color = str(cell.fill.fgColor.value)
            elif hasattr(cell.fill.fgColor, 'index'):
                # Pour les couleurs indexées, essayer de les convertir
                # Les couleurs indexées sont parfois stockées différemment
                pass
        
        # 2. start_color (pour les dégradés)
        if not color and hasattr(cell.fill, 'start_color') and cell.fill.start_color:
            if hasattr(cell.fill.start_color, 'index'):
                # Couleur indexée
                if hasattr(cell.fill.start_color, 'rgb') and cell.fill.start_color.rgb:
                    color = cell.fill.start_color.rgb
                elif hasattr(cell.fill.start_color, 'value'):
                    color = str(cell.fill.start_color.value)
            elif hasattr(cell.fill.start_color, 'rgb') and cell.fill.start_color.rgb:
                color = cell.fill.start_color.rgb
            elif hasattr(cell.fill.start_color, 'value') and cell.fill.start_color.value:
                color = str(cell.fill.start_color.value)
        
        # 3. bgColor (couleur de fond)
        if not color and hasattr(cell.fill, 'bgColor') and cell.fill.bgColor:
            if hasattr(cell.fill.bgColor, 'rgb') and cell.fill.bgColor.rgb:
                color = cell.fill.bgColor.rgb
            elif hasattr(cell.fill.bgColor, 'value') and cell.fill.bgColor.value:
                color = str(cell.fill.bgColor.value)
        
        # 4. patternType solid avec une couleur
        if not color and hasattr(cell.fill, 'patternType') and cell.fill.patternType == 'solid':
            # Pour les patterns solid, la couleur peut être dans fgColor ou start_color
            if hasattr(cell.fill, 'fgColor') and cell.fill.fgColor:
                # Essayer toutes les propriétés possibles
                for attr in ['rgb', 'value', 'theme', 'tint']:
                    if hasattr(cell.fill.fgColor, attr):
                        val = getattr(cell.fill.fgColor, attr)
                        if val:
                            color = str(val)
                            break
        
        # 5. Gestion des couleurs theme + tint
        if not color and hasattr(cell.fill, 'fgColor') and cell.fill.fgColor:
            if hasattr(cell.fill.fgColor, 'theme') and cell.fill.fgColor.theme is not None:
                # Les couleurs theme sont des couleurs du thème Excel
                theme = cell.fill.fgColor.theme
                tint = getattr(cell.fill.fgColor, 'tint', 0)
                # Créer un code couleur basé sur le theme pour le debug
                color = f"THEME{theme:02d}"
                if abs(tint) > 0.01:  # Si tint significatif
                    tint_hex = int((tint + 1) * 127.5)
                    color = f"T{theme:02d}{tint_hex:02X}"
        
        if color:
            # Nettoyer et valider la couleur
            hex_color = rgb_to_hex(color)
            return hex_color
            
    except Exception as e:
        print(f"Erreur lors de l'extraction de la couleur: {e}")
    
    return None

def rgb_to_hex(rgb: Union[str, Tuple[int, int, int]]) -> str:
    """Convertit RGB en hexadécimal"""
    if isinstance(rgb, str):
        # Nettoyer la chaîne
        rgb = rgb.strip().upper()
        
        # Si c'est déjà en hex (avec ou sans #)
        if rgb.startswith('#'):
            rgb = rgb[1:]
        
        # Si c'est un hex valide
        if len(rgb) == 6 and all(c in '0123456789ABCDEF' for c in rgb):
            return rgb
        elif len(rgb) == 8 and all(c in '0123456789ABCDEF' for c in rgb):
            return rgb[2:]  # Enlever le canal alpha
        
        # Si c'est une valeur numérique (parfois Excel stocke les couleurs comme des entiers)
        try:
            if rgb.isdigit():
                # Convertir l'entier en hex
                color_int = int(rgb)
                hex_color = f"{color_int:06X}"
                return hex_color
        except:
            pass
            
    elif isinstance(rgb, (tuple, list)) and len(rgb) >= 3:
        return '%02X%02X%02X' % (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    elif isinstance(rgb, int):
        # Parfois les couleurs sont stockées comme des entiers
        return f"{rgb:06X}"
    
    return "FFFFFF"

def get_merged_cells_info(worksheet) -> Dict[Tuple[int, int], Dict]:
    """
    Retourne les informations sur les cellules fusionnées
    Format: {(row, col): {'min_row': x, 'max_row': y, 'min_col': a, 'max_col': b}}
    """
    merged_info = {}
    
    if hasattr(worksheet, 'merged_cells'):
        for merged_range in worksheet.merged_cells.ranges:
            min_row = merged_range.min_row
            max_row = merged_range.max_row
            min_col = merged_range.min_col
            max_col = merged_range.max_col
            
            # Ajouter l'info pour la cellule principale (top-left)
            merged_info[(min_row, min_col)] = {
                'min_row': min_row,
                'max_row': max_row,
                'min_col': min_col,
                'max_col': max_col,
                'span_rows': max_row - min_row + 1,
                'span_cols': max_col - min_col + 1
            }
            
            # Marquer toutes les cellules de la plage comme faisant partie d'une fusion
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    if (row, col) != (min_row, min_col):
                        merged_info[(row, col)] = {
                            'is_merged_cell': True,
                            'main_cell': (min_row, min_col),
                            'min_row': min_row,
                            'max_row': max_row,
                            'min_col': min_col,
                            'max_col': max_col
                        }
    
    return merged_info

def get_cell_value(cell) -> Any:
    """Retourne la valeur d'une cellule de manière sûre"""
    if cell.value is None:
        return ""
    return cell.value

def get_worksheet_dimensions(worksheet) -> Tuple[int, int]:
    """Retourne les dimensions (lignes, colonnes) d'une feuille"""
    return worksheet.max_row, worksheet.max_column