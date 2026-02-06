"""
Script Python pour ouvrir et lire un fichier DWG à partir d'un chemin spécifié
Nécessite: pip install pywin32
"""

import win32com.client
import os

def ouvrir_et_lire_dwg(chemin_dwg):
    """
    Ouvre un fichier DWG et lit ses informations
    
    Args:
        chemin_dwg: Chemin complet du fichier .dwg
    """
    try:
        # Vérifier si le fichier existe
        if not os.path.exists(chemin_dwg):
            print(f"ERREUR: Le fichier n'existe pas: {chemin_dwg}")
            return None
        
        # Vérifier l'extension
        if not chemin_dwg.lower().endswith('.dwg'):
            print(f"ERREUR: Le fichier doit être un .dwg")
            return None
        
        print(f"Ouverture du fichier: {chemin_dwg}")
        print("Veuillez patienter...\n")
        
        # Connexion à AutoCAD
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True  # Rendre AutoCAD visible
        
        # Ouvrir le document
        doc = acad.Documents.Open(chemin_dwg)
        
        print("="*70)
        print("FICHIER OUVERT AVEC SUCCÈS")
        print("="*70)
        
        # Afficher les informations du fichier
        print(f"\nChemin complet: {doc.FullName}")
        print(f"Nom du fichier: {doc.Name}")
        print(f"Répertoire: {doc.Path}")
        
        # Informations sur le contenu
        print(f"\n--- CONTENU DU DESSIN ---")
        
        # Compter les entités dans ModelSpace
        modelspace_count = doc.ModelSpace.Count
        print(f"Entités dans ModelSpace: {modelspace_count}")
        
        # Compter les entités dans PaperSpace
        paperspace_count = doc.PaperSpace.Count
        print(f"Entités dans PaperSpace: {paperspace_count}")
        
        # Compter les calques (layers)
        layers_count = doc.Layers.Count
        print(f"Nombre de calques: {layers_count}")
        
        # Compter les blocs
        blocks_count = doc.Blocks.Count
        print(f"Nombre de blocs: {blocks_count}")
        
        return doc
        
    except Exception as e:
        print(f"ERREUR: {str(e)}")
        return None


def lire_blocs_dynamiques_du_fichier(chemin_dwg):
    """
    Ouvre un fichier DWG et lit tous ses blocs dynamiques
    
    Args:
        chemin_dwg: Chemin complet du fichier .dwg
    """
    try:
        if not os.path.exists(chemin_dwg):
            print(f"ERREUR: Le fichier n'existe pas: {chemin_dwg}")
            return
        
        # Connexion à AutoCAD
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
        
        # Ouvrir le document
        doc = acad.Documents.Open(chemin_dwg)
        
        print(f"\n=== LECTURE DES BLOCS DYNAMIQUES ===")
        print(f"Fichier: {doc.Name}\n")
        
        blocs_dynamiques_trouves = 0
        
        # Parcourir ModelSpace
        for entity in doc.ModelSpace:
            if entity.ObjectName == "AcDbBlockReference":
                if entity.IsDynamicBlock:
                    blocs_dynamiques_trouves += 1
                    
                    print(f"\n--- BLOC DYNAMIQUE #{blocs_dynamiques_trouves} ---")
                    print(f"Nom: {entity.Name}")
                    print(f"Nom effectif: {entity.EffectiveName}")
                    
                    # Propriétés dynamiques
                    dynamic_props = entity.GetDynamicBlockProperties()
                    print(f"Propriétés dynamiques: {len(dynamic_props)}")
                    
                    for prop in dynamic_props:
                        print(f"  • {prop.PropertyName}: {prop.Value}")
                    
                    # Position
                    print(f"Position: X={entity.InsertionPoint[0]:.2f}, Y={entity.InsertionPoint[1]:.2f}, Z={entity.InsertionPoint[2]:.2f}")
        
        if blocs_dynamiques_trouves == 0:
            print("Aucun bloc dynamique trouvé dans ce dessin.")
        else:
            print(f"\n=== TOTAL: {blocs_dynamiques_trouves} bloc(s) dynamique(s) trouvé(s) ===")
        
        return doc
        
    except Exception as e:
        print(f"ERREUR: {str(e)}")
        return None


def lire_infos_detaillees(chemin_dwg):
    """
    Lit toutes les informations détaillées d'un fichier DWG
    
    Args:
        chemin_dwg: Chemin complet du fichier .dwg
    """
    try:
        if not os.path.exists(chemin_dwg):
            print(f"ERREUR: Le fichier n'existe pas: {chemin_dwg}")
            return
        
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
        doc = acad.Documents.Open(chemin_dwg)
        
        print("\n" + "="*70)
        print("INFORMATIONS DÉTAILLÉES DU FICHIER DWG")
        print("="*70)
        
        print(f"\n📁 FICHIER:")
        print(f"   Chemin: {doc.FullName}")
        print(f"   Nom: {doc.Name}")
        
        print(f"\n📊 STATISTIQUES:")
        print(f"   Entités ModelSpace: {doc.ModelSpace.Count}")
        print(f"   Entités PaperSpace: {doc.PaperSpace.Count}")
        print(f"   Calques: {doc.Layers.Count}")
        print(f"   Blocs: {doc.Blocks.Count}")
        
        # Lister les calques
        print(f"\n🎨 CALQUES:")
        for layer in doc.Layers:
            status = "ON" if layer.LayerOn else "OFF"
            print(f"   • {layer.Name} [{status}]")
        
        # Compter les types d'entités
        print(f"\n📐 TYPES D'ENTITÉS:")
        entity_types = {}
        for entity in doc.ModelSpace:
            obj_name = entity.ObjectName
            entity_types[obj_name] = entity_types.get(obj_name, 0) + 1
        
        for obj_type, count in sorted(entity_types.items()):
            print(f"   • {obj_type}: {count}")
        
        # Limites du dessin
        print(f"\n📏 LIMITES:")
        try:
            extmin = doc.GetVariable("EXTMIN")
            extmax = doc.GetVariable("EXTMAX")
            print(f"   Min: ({extmin[0]:.2f}, {extmin[1]:.2f}, {extmin[2]:.2f})")
            print(f"   Max: ({extmax[0]:.2f}, {extmax[1]:.2f}, {extmax[2]:.2f})")
        except:
            print("   Limites non disponibles")
        
        return doc
        
    except Exception as e:
        print(f"ERREUR: {str(e)}")
        return None


if __name__ == "__main__":
    # ====================================================================
    # METTEZ VOTRE CHEMIN ICI
    # ====================================================================
    
    # Exemple Windows:
    chemin_mon_fichier = r"D:\desktop\Correction Memoir Lycence\Plan Maison R+1.dwg"
    
    # Exemple avec double backslash:
    # chemin_mon_fichier = "C:\\Users\\TSARALOHA Christinot\\Documents\\Matrix One\\Draw One.dwg"
    
    # Exemple avec slash (fonctionne aussi):
    # chemin_mon_fichier = "C:/Users/TSARALOHA Christinot/Documents/Matrix One/Draw One.dwg"
    
    # ====================================================================
    
    print("="*70)
    print("LECTURE D'UN FICHIER DWG SPÉCIFIQUE")
    print("="*70)
    
    # Option 1: Ouvrir et lire les informations de base
    doc = ouvrir_et_lire_dwg(chemin_mon_fichier)
    
    # Option 2: Lire les blocs dynamiques (décommenter pour utiliser)
    # doc = lire_blocs_dynamiques_du_fichier(chemin_mon_fichier)
    
    # Option 3: Lire toutes les informations détaillées (décommenter pour utiliser)
    # doc = lire_infos_detaillees(chemin_mon_fichier)
    
    if doc:
        print("\n✅ Fichier ouvert avec succès dans AutoCAD!")
        print("Vous pouvez maintenant travailler avec le fichier.")
    else:
        print("\n❌ Échec de l'ouverture du fichier.")