import json 
import os 

from pathlib import Path as path 

def open_json(fils : str = "abscisse.json", folder = "Curves"):

    file_path = path.home() / "Documents/Matrix One/Influence Line" / folder / fils
    with open(file_path , "r", encoding="utf-8") as file:
        return json.load(file)
    
# Configuration par défaut
DEFAULT_CONFIG = {
    "grid": True,
    "travee": True,
    "noeud": True,
    "vitesse": 10,
    "vitesse_bridge": 0.005,
    "legend": True,
    "axe_y_inverser": False,
    "default_matplotlib_color":True, 
    "default_matplotlib_style":True,
    "style": {
        "line_color": "royalblue",
        "grid_color": "#E0E0E0",
        "minor_grid_color": "#F5F5F5",
        "noeud_color": "#FF4444",
        "noeud_size": 100,
        "line_width": 2.5,
        "line_style": "-",
        "background_color": "#FFFFFF",
        "edge_color": "#2C3E50",
        "shadow_color": "gray",
        "shadow_alpha": 0.3,
        "title": "",
        "xlabel": "Longueur des travées",
        "ylabel": "Valeur",
        "axis_fontsize": 12,
        "font_family": "Times New Roman",
        "legend_position": "best",
        "marker_style": "o",
        "legend_fontsize": 10
    }
}

x_normal = open_json("abscisse.json")  
x_forces = open_json("abscisse_T.json")

neouds = open_json("noeud_lengths.json")

distances = []
for i in range(len(neouds)):
    distance = round(neouds[i], 5)
    distances.append(f"{distance}")

# Fusionner avec la configuration par défaut
current_config = DEFAULT_CONFIG.copy()
text_format = {"family":"Times New Roman", "size":20} 