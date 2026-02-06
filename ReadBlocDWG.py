"""
Script Python pour extraire TOUTES les informations des BLOCS uniquement
Nécessite: pip install pywin32
"""

import win32com.client
import os
import json
import time
from datetime import datetime

class ExtracteurBlocs:
    """Classe pour extraire toutes les informations des blocs d'un fichier DWG"""
    
    def __init__(self, chemin_dwg):
        self.chemin_dwg = chemin_dwg
        self.acad = None
        self.doc = None
        self.blocs_info = {
            'definitions_blocs': [],
            'instances_blocs': [],
            'blocs_dynamiques': [],
            'statistiques': {}
        }
        
    def ouvrir_fichier(self):
        """Ouvre le fichier DWG"""
        try:
            if not os.path.exists(self.chemin_dwg):
                raise FileNotFoundError(f"Le fichier n'existe pas: {self.chemin_dwg}")
            
            print(f"Ouverture du fichier: {self.chemin_dwg}")
            print("Veuillez patienter...\n")
            
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.acad.Visible = True
            
            # Ouvrir le document
            self.doc = self.acad.Documents.Open(self.chemin_dwg)
            
            # IMPORTANT: Attendre que le document soit complètement chargé
            print("⏳ Chargement du document en cours...")
            time.sleep(3)  # Attendre 3 secondes
            
            # Forcer AutoCAD à terminer le chargement
            try:
                self.acad.Update()
            except:
                pass
            
            time.sleep(2)  # Attendre encore 2 secondes
            
            print("✅ Fichier ouvert avec succès!\n")
            return True
            
        except Exception as e:
            print(f"❌ ERREUR lors de l'ouverture: {str(e)}")
            return False
    
    def extraire_definitions_blocs(self):
        """Extrait toutes les définitions de blocs (Block Definitions)"""
        print("🔲 Extraction des définitions de blocs...")
        
        # Plusieurs tentatives en cas d'erreur COM
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # Convertir en liste pour éviter les problèmes d'itération COM
                blocks_list = []
                blocks_count = self.doc.Blocks.Count
                
                for i in range(blocks_count):
                    try:
                        block = self.doc.Blocks.Item(i)
                        blocks_list.append(block)
                    except:
                        continue
                
                # Maintenant extraire les informations
                for block in blocks_list:
                    try:
                        bloc_def = {
                            'nom': block.Name,
                            'est_dynamique': False,
                            'est_xref': block.IsXRef,
                            'est_layout': block.IsLayout,
                            'nombre_entites': block.Count,
                            'origine': None,
                            'entites_contenues': [],
                            'attributs': []
                        }
                        
                        # Origine du bloc
                        try:
                            bloc_def['origine'] = {
                                'x': block.Origin[0],
                                'y': block.Origin[1],
                                'z': block.Origin[2]
                            }
                        except:
                            pass
                        
                        # Vérifier si c'est un bloc dynamique
                        try:
                            bloc_def['est_dynamique'] = block.IsDynamicBlock
                        except:
                            pass
                        
                        # Lister les entités contenues dans le bloc
                        entites_types = {}
                        try:
                            for j in range(block.Count):
                                try:
                                    entity = block.Item(j)
                                    obj_type = entity.ObjectName
                                    entites_types[obj_type] = entites_types.get(obj_type, 0) + 1
                                    
                                    # Détails de l'entité
                                    entite_info = {
                                        'type': obj_type,
                                        'calque': entity.Layer if hasattr(entity, 'Layer') else None
                                    }
                                    
                                    # Informations spécifiques pour les attributs
                                    if obj_type == "AcDbAttributeDefinition":
                                        entite_info['tag'] = entity.TagString if hasattr(entity, 'TagString') else None
                                        entite_info['prompt'] = entity.PromptString if hasattr(entity, 'PromptString') else None
                                        entite_info['valeur_defaut'] = entity.TextString if hasattr(entity, 'TextString') else None
                                        entite_info['constant'] = entity.Constant if hasattr(entity, 'Constant') else None
                                        entite_info['invisible'] = entity.Invisible if hasattr(entity, 'Invisible') else None
                                        entite_info['preset'] = entity.Preset if hasattr(entity, 'Preset') else None
                                        entite_info['verification'] = entity.Verify if hasattr(entity, 'Verify') else None
                                        
                                        bloc_def['attributs'].append(entite_info)
                                    
                                    # Info géométrique selon le type
                                    elif obj_type == "AcDbLine":
                                        entite_info['debut'] = list(entity.StartPoint)
                                        entite_info['fin'] = list(entity.EndPoint)
                                        entite_info['longueur'] = entity.Length
                                    
                                    elif obj_type == "AcDbCircle":
                                        entite_info['centre'] = list(entity.Center)
                                        entite_info['rayon'] = entity.Radius
                                    
                                    elif obj_type == "AcDbArc":
                                        entite_info['centre'] = list(entity.Center)
                                        entite_info['rayon'] = entity.Radius
                                        entite_info['angle_debut'] = entity.StartAngle
                                        entite_info['angle_fin'] = entity.EndAngle
                                    
                                    elif obj_type == "AcDbText" or obj_type == "AcDbMText":
                                        entite_info['texte'] = entity.TextString
                                        entite_info['hauteur'] = entity.Height
                                    
                                    bloc_def['entites_contenues'].append(entite_info)
                                    
                                except Exception as e:
                                    continue
                        except:
                            pass
                        
                        bloc_def['types_entites'] = entites_types
                        bloc_def['nombre_attributs'] = len(bloc_def['attributs'])
                        
                        self.blocs_info['definitions_blocs'].append(bloc_def)
                        
                    except Exception as e:
                        continue
                
                print(f"  ✓ {len(self.blocs_info['definitions_blocs'])} définitions de blocs extraites")
                return  # Succès, sortir de la fonction
                
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"  ⚠️  Tentative {attempt + 1} échouée, nouvelle tentative...")
                    time.sleep(2)
                else:
                    print(f"  ❌ ERREUR après {max_retries} tentatives: {str(e)}")
                    print(f"  ℹ️  {len(self.blocs_info['definitions_blocs'])} blocs extraits avant l'erreur")
    
    def extraire_instances_blocs(self):
        """Extrait toutes les instances de blocs dans le dessin"""
        print("\n📍 Extraction des instances de blocs...")
        
        # Dans ModelSpace
        print("  → Parcours du ModelSpace...")
        for entity in self.doc.ModelSpace:
            try:
                if entity.ObjectName == "AcDbBlockReference":
                    instance = self._extraire_info_instance(entity, "ModelSpace")
                    if instance:
                        self.blocs_info['instances_blocs'].append(instance)
            except:
                pass
        
        # Dans PaperSpace
        print("  → Parcours du PaperSpace...")
        for entity in self.doc.PaperSpace:
            try:
                if entity.ObjectName == "AcDbBlockReference":
                    instance = self._extraire_info_instance(entity, "PaperSpace")
                    if instance:
                        self.blocs_info['instances_blocs'].append(instance)
            except:
                pass
        
        print(f"  ✓ {len(self.blocs_info['instances_blocs'])} instances de blocs extraites")
    
    def _extraire_info_instance(self, entity, espace):
        """Extrait les informations d'une instance de bloc"""
        try:
            instance = {
                'nom_bloc': entity.Name,
                'espace': espace,
                'est_dynamique': entity.IsDynamicBlock,
                'position': {
                    'x': entity.InsertionPoint[0],
                    'y': entity.InsertionPoint[1],
                    'z': entity.InsertionPoint[2]
                },
                'rotation': entity.Rotation,
                'rotation_degres': entity.Rotation * 180 / 3.14159265359,
                'echelle': {
                    'x': entity.XScaleFactor,
                    'y': entity.YScaleFactor,
                    'z': entity.ZScaleFactor
                },
                'calque': entity.Layer,
                'couleur': entity.Color if hasattr(entity, 'Color') else None,
                'type_ligne': entity.Linetype if hasattr(entity, 'Linetype') else None,
                'epaisseur_ligne': entity.Lineweight if hasattr(entity, 'Lineweight') else None,
                'visible': entity.Visible if hasattr(entity, 'Visible') else None,
                'attributs': [],
                'handle': entity.Handle if hasattr(entity, 'Handle') else None
            }
            
            # Extraire les attributs de l'instance
            try:
                attributes = entity.GetAttributes()
                for attr in attributes:
                    attr_info = {
                        'tag': attr.TagString,
                        'valeur': attr.TextString,
                        'invisible': attr.Invisible,
                        'hauteur': attr.Height,
                        'position': {
                            'x': attr.InsertionPoint[0],
                            'y': attr.InsertionPoint[1],
                            'z': attr.InsertionPoint[2]
                        }
                    }
                    instance['attributs'].append(attr_info)
            except:
                pass
            
            instance['nombre_attributs'] = len(instance['attributs'])
            
            # Si c'est un bloc dynamique, extraire le nom effectif
            if instance['est_dynamique']:
                try:
                    instance['nom_effectif'] = entity.EffectiveName
                except:
                    pass
            
            return instance
            
        except Exception as e:
            return None
    
    def extraire_blocs_dynamiques(self):
        """Extrait toutes les informations des blocs dynamiques"""
        print("\n🔄 Extraction des blocs dynamiques...")
        
        # Dans ModelSpace
        for entity in self.doc.ModelSpace:
            try:
                if entity.ObjectName == "AcDbBlockReference" and entity.IsDynamicBlock:
                    bloc_dyn = self._extraire_info_bloc_dynamique(entity, "ModelSpace")
                    if bloc_dyn:
                        self.blocs_info['blocs_dynamiques'].append(bloc_dyn)
            except:
                pass
        
        # Dans PaperSpace
        for entity in self.doc.PaperSpace:
            try:
                if entity.ObjectName == "AcDbBlockReference" and entity.IsDynamicBlock:
                    bloc_dyn = self._extraire_info_bloc_dynamique(entity, "PaperSpace")
                    if bloc_dyn:
                        self.blocs_info['blocs_dynamiques'].append(bloc_dyn)
            except:
                pass
        
        print(f"  ✓ {len(self.blocs_info['blocs_dynamiques'])} blocs dynamiques extraits")
    
    def _extraire_info_bloc_dynamique(self, entity, espace):
        """Extrait les informations complètes d'un bloc dynamique"""
        try:
            bloc_dyn = {
                'nom': entity.Name,
                'nom_effectif': entity.EffectiveName,
                'espace': espace,
                'position': {
                    'x': entity.InsertionPoint[0],
                    'y': entity.InsertionPoint[1],
                    'z': entity.InsertionPoint[2]
                },
                'rotation': entity.Rotation,
                'rotation_degres': entity.Rotation * 180 / 3.14159265359,
                'echelle': {
                    'x': entity.XScaleFactor,
                    'y': entity.YScaleFactor,
                    'z': entity.ZScaleFactor
                },
                'calque': entity.Layer,
                'couleur': entity.Color if hasattr(entity, 'Color') else None,
                'handle': entity.Handle if hasattr(entity, 'Handle') else None,
                'proprietes_dynamiques': [],
                'attributs': []
            }
            
            # Extraire les propriétés dynamiques
            dynamic_props = entity.GetDynamicBlockProperties()
            for prop in dynamic_props:
                try:
                    prop_info = {
                        'nom': prop.PropertyName,
                        'valeur': prop.Value,
                        'lecture_seule': prop.ReadOnly,
                        'type_unite': prop.UnitsType,
                        'description': prop.Description if hasattr(prop, 'Description') else None
                    }
                    
                    # Valeurs autorisées
                    try:
                        allowed = prop.AllowedValues
                        if allowed:
                            prop_info['valeurs_autorisees'] = list(allowed)
                            prop_info['nombre_valeurs_autorisees'] = len(list(allowed))
                    except:
                        pass
                    
                    bloc_dyn['proprietes_dynamiques'].append(prop_info)
                except:
                    pass
            
            bloc_dyn['nombre_proprietes_dynamiques'] = len(bloc_dyn['proprietes_dynamiques'])
            
            # Extraire les attributs
            try:
                attributes = entity.GetAttributes()
                for attr in attributes:
                    attr_info = {
                        'tag': attr.TagString,
                        'valeur': attr.TextString,
                        'invisible': attr.Invisible,
                        'hauteur': attr.Height
                    }
                    bloc_dyn['attributs'].append(attr_info)
            except:
                pass
            
            bloc_dyn['nombre_attributs'] = len(bloc_dyn['attributs'])
            
            return bloc_dyn
            
        except Exception as e:
            return None
    
    def calculer_statistiques(self):
        """Calcule des statistiques sur les blocs"""
        print("\n📊 Calcul des statistiques...")
        
        # Statistiques des définitions
        total_defs = len(self.blocs_info['definitions_blocs'])
        defs_dynamiques = sum(1 for b in self.blocs_info['definitions_blocs'] if b['est_dynamique'])
        defs_xref = sum(1 for b in self.blocs_info['definitions_blocs'] if b['est_xref'])
        defs_avec_attributs = sum(1 for b in self.blocs_info['definitions_blocs'] if b['nombre_attributs'] > 0)
        
        # Statistiques des instances
        total_instances = len(self.blocs_info['instances_blocs'])
        instances_modelspace = sum(1 for b in self.blocs_info['instances_blocs'] if b['espace'] == 'ModelSpace')
        instances_paperspace = sum(1 for b in self.blocs_info['instances_blocs'] if b['espace'] == 'PaperSpace')
        instances_dynamiques = sum(1 for b in self.blocs_info['instances_blocs'] if b['est_dynamique'])
        
        # Blocs les plus utilisés
        utilisation_blocs = {}
        for instance in self.blocs_info['instances_blocs']:
            nom = instance['nom_bloc']
            utilisation_blocs[nom] = utilisation_blocs.get(nom, 0) + 1
        
        top_blocs = sorted(utilisation_blocs.items(), key=lambda x: x[1], reverse=True)[:10]
        
        # Calques utilisés par les blocs
        calques_blocs = set(b['calque'] for b in self.blocs_info['instances_blocs'] if b['calque'])
        
        self.blocs_info['statistiques'] = {
            'definitions': {
                'total': total_defs,
                'dynamiques': defs_dynamiques,
                'xref': defs_xref,
                'avec_attributs': defs_avec_attributs
            },
            'instances': {
                'total': total_instances,
                'modelspace': instances_modelspace,
                'paperspace': instances_paperspace,
                'dynamiques': instances_dynamiques
            },
            'top_10_blocs_utilises': [{'nom': nom, 'nombre': count} for nom, count in top_blocs],
            'nombre_calques_utilises': len(calques_blocs),
            'calques_utilises': sorted(list(calques_blocs))
        }
        
        print(f"  ✓ Statistiques calculées")
    
    def extraire_tout(self):
        """Extrait toutes les informations des blocs"""
        if not self.ouvrir_fichier():
            return False
        
        print("\n" + "="*80)
        print("EXTRACTION DE TOUTES LES INFORMATIONS DES BLOCS")
        print("="*80)
        
        self.extraire_definitions_blocs()
        self.extraire_instances_blocs()
        self.extraire_blocs_dynamiques()
        self.calculer_statistiques()
        
        print("\n" + "="*80)
        print("✅ EXTRACTION TERMINÉE!")
        print("="*80)
        
        return True
    
    def afficher_resume(self):
        """Affiche un résumé des blocs extraits"""
        stats = self.blocs_info['statistiques']
        
        print("\n" + "="*80)
        print("📊 RÉSUMÉ DES BLOCS")
        print("="*80)
        
        print(f"\n🔲 DÉFINITIONS DE BLOCS:")
        print(f"  Total: {stats['definitions']['total']}")
        print(f"  Dynamiques: {stats['definitions']['dynamiques']}")
        print(f"  XRef: {stats['definitions']['xref']}")
        print(f"  Avec attributs: {stats['definitions']['avec_attributs']}")
        
        print(f"\n📍 INSTANCES DE BLOCS:")
        print(f"  Total: {stats['instances']['total']}")
        print(f"  Dans ModelSpace: {stats['instances']['modelspace']}")
        print(f"  Dans PaperSpace: {stats['instances']['paperspace']}")
        print(f"  Dynamiques: {stats['instances']['dynamiques']}")
        
        print(f"\n🏆 TOP 10 DES BLOCS LES PLUS UTILISÉS:")
        for i, bloc in enumerate(stats['top_10_blocs_utilises'][:10], 1):
            print(f"  {i}. {bloc['nom']}: {bloc['nombre']} fois")
        
        print(f"\n🎨 CALQUES:")
        print(f"  Nombre de calques utilisés par les blocs: {stats['nombre_calques_utilises']}")
    
    def sauvegarder_json(self, fichier_sortie=None):
        """Sauvegarde toutes les informations des blocs dans un fichier JSON"""
        if fichier_sortie is None:
            nom_base = os.path.splitext(os.path.basename(self.chemin_dwg))[0]
            fichier_sortie = f"{nom_base}_blocs.json"
        
        with open(fichier_sortie, 'w', encoding='utf-8') as f:
            json.dump(self.blocs_info, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"\n💾 Données des blocs sauvegardées dans: {os.path.abspath(fichier_sortie)}")
        return fichier_sortie
    
    def sauvegarder_rapport(self, fichier_sortie=None):
        """Sauvegarde un rapport détaillé sur les blocs"""
        if fichier_sortie is None:
            nom_base = os.path.splitext(os.path.basename(self.chemin_dwg))[0]
            fichier_sortie = f"{nom_base}_rapport_blocs.txt"
        
        with open(fichier_sortie, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("RAPPORT DÉTAILLÉ DES BLOCS\n")
            f.write("="*80 + "\n\n")
            
            # Statistiques
            stats = self.blocs_info['statistiques']
            f.write("📊 STATISTIQUES:\n")
            f.write(f"  Définitions de blocs: {stats['definitions']['total']}\n")
            f.write(f"  Instances de blocs: {stats['instances']['total']}\n")
            f.write(f"  Blocs dynamiques: {stats['instances']['dynamiques']}\n\n")
            
            # Définitions de blocs
            f.write("="*80 + "\n")
            f.write("🔲 DÉFINITIONS DE BLOCS\n")
            f.write("="*80 + "\n\n")
            
            for bloc_def in self.blocs_info['definitions_blocs']:
                f.write(f"Bloc: {bloc_def['nom']}\n")
                f.write(f"  Dynamique: {bloc_def['est_dynamique']}\n")
                f.write(f"  XRef: {bloc_def['est_xref']}\n")
                f.write(f"  Nombre d'entités: {bloc_def['nombre_entites']}\n")
                f.write(f"  Nombre d'attributs: {bloc_def['nombre_attributs']}\n")
                
                if bloc_def['attributs']:
                    f.write(f"  Attributs:\n")
                    for attr in bloc_def['attributs']:
                        f.write(f"    - {attr.get('tag', 'N/A')}: {attr.get('prompt', 'N/A')}\n")
                
                if bloc_def['types_entites']:
                    f.write(f"  Types d'entités:\n")
                    for type_ent, count in bloc_def['types_entites'].items():
                        f.write(f"    - {type_ent}: {count}\n")
                
                f.write("\n")
            
            # Blocs dynamiques
            f.write("="*80 + "\n")
            f.write("🔄 BLOCS DYNAMIQUES\n")
            f.write("="*80 + "\n\n")
            
            for bloc_dyn in self.blocs_info['blocs_dynamiques']:
                f.write(f"Bloc: {bloc_dyn['nom_effectif']}\n")
                f.write(f"  Nom original: {bloc_dyn['nom']}\n")
                f.write(f"  Position: X={bloc_dyn['position']['x']:.2f}, Y={bloc_dyn['position']['y']:.2f}, Z={bloc_dyn['position']['z']:.2f}\n")
                f.write(f"  Rotation: {bloc_dyn['rotation_degres']:.2f}°\n")
                f.write(f"  Calque: {bloc_dyn['calque']}\n")
                f.write(f"  Propriétés dynamiques:\n")
                
                for prop in bloc_dyn['proprietes_dynamiques']:
                    f.write(f"    • {prop['nom']}: {prop['valeur']}")
                    if prop.get('valeurs_autorisees'):
                        f.write(f" (Valeurs: {prop['valeurs_autorisees']})")
                    f.write("\n")
                
                if bloc_dyn['attributs']:
                    f.write(f"  Attributs:\n")
                    for attr in bloc_dyn['attributs']:
                        f.write(f"    • {attr['tag']}: {attr['valeur']}\n")
                
                f.write("\n")
            
            # Top blocs
            f.write("="*80 + "\n")
            f.write("🏆 TOP 10 DES BLOCS LES PLUS UTILISÉS\n")
            f.write("="*80 + "\n\n")
            
            for i, bloc in enumerate(stats['top_10_blocs_utilises'][:10], 1):
                f.write(f"{i}. {bloc['nom']}: {bloc['nombre']} instances\n")
        
        print(f"📄 Rapport des blocs sauvegardé dans: {os.path.abspath(fichier_sortie)}")
        return fichier_sortie


if __name__ == "__main__":
    # ====================================================================
    # METTEZ VOTRE CHEMIN ICI
    # ====================================================================
    
    chemin_dwg = r"D:\desktop\Correction Memoir Lycence\Plan Maison R+1.dwg"
    
    # ====================================================================
    
    # Créer l'extracteur
    extracteur = ExtracteurBlocs(chemin_dwg)
    
    # Extraire toutes les informations des blocs
    if extracteur.extraire_tout():
        # Afficher le résumé
        extracteur.afficher_resume()
        
        # Sauvegarder en JSON
        fichier_json = extracteur.sauvegarder_json()
        
        # Sauvegarder le rapport
        fichier_txt = extracteur.sauvegarder_rapport()
        
        print(f"\n✅ Extraction des blocs terminée!")
        print(f"   - JSON: {fichier_json}")
        print(f"   - Rapport: {fichier_txt}")