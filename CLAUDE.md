# CLAUDE.md — Contexte projet automatisation rapports Ansys Mechanical

> **Version :** 1.0  
> **Date :** 2025-07-17  
> **Type :** Fichier d'instructions pour Claude Code  
> **Langue du code :** English (phase 1 de développement)

---

## 1. Contexte & Objectif du Projet

### Problème métier

Les ingénieurs calculs utilisent Ansys Mechanical pour effectuer des analyses par éléments finis. Une fois l'analyse terminée, ils doivent créer des rapports PDF/PPTX intégrant :

- **Captures d'écran** des résultats visuels (déformations, contraintes, etc.)
- **Tableaux de résultats** copiés manuellement depuis les tables de données d'Ansys
- **Données multi-cas de charge** (plusieurs Steps / load cases) — très chronophage à extraire case par case

Cette tâche de post-traitement est actuellement **manuelle et répétitive**, ralentissant significativement le workflow de l'ingénieur.

### Solution

Création d'un script Python (IronPython 2.7 + .NET) exécutable directement dans la **console de scripting d'Ansys Mechanical 2025 R2**. Le script expose une **interface graphique Windows Forms** avec un bouton unique qui permet :

1. De sélectionner les données à extraire (géométrie, maillage, conditions aux limites, résultats)
2. D'exporter les données (CSV pour les tableaux, images pour les visuels)
3. De peupler automatiquement un **PowerPoint corporate** à partir d'un template imposé
4. De générer des tableaux Excel via CSV pour archivage

### Livrable attendu

- **Un seul fichier script** exécutable dans la console de scripting Ansys Mechanical.
- **Interface Windows Forms** intégrée (pas de fenêtre externe, tout dans Mechanical).
- **Fonctionnement en un clic** : l'ingénieur configure, appuie sur le bouton, le rapport est généré.

---

## 2. Contraintes Techniques Strictes

> ⚠️ **Ces contraintes sont ABSOLUES. Tout écart sera cause d'erreur immédiate dans Ansys Mechanical.**

### 2.1 Environnement d'exécution

| Élément | Valeur imposée |
|---------|----------------|
| Runtime Python | **IronPython 2.7** (embarqué dans Ansys Mechanical) |
| Framework .NET | Utilisable directement via `clr.AddReference()` |
| Interface graphique | **Windows Forms** (`System.Windows.Forms`) via .NET |
| Communication Office | **COM Interop .NET** (`Microsoft.Office.Interop.PowerPoint`, `.Excel`) |

### 2.2 Contraintes IronPython 2.7 — Liste des interdictions absolues

```
À NE PAS FAIRE → Python 3 uniquement (incompatible avec IronPython 2.7)
```

- ❌ **F-strings** — utiliser `.format()` à la place
  ```python
  # INTERDIT
  name = f"result_{i}"
  
  # CORRECT
  name = "result_{}".format(i)
  ```

- ❌ **`print()` comme fonction** — c'est une instruction, pas une fonction
  ```python
  # INTERDIT
  print("Done")
  
  # CORRECT
  print "Done"
  ```

- ❌ **`pathlib`** — non disponible dans IronPython 2.7
  ```python
  # INTERDIT
  from pathlib import Path
  
  # CORRECT
  import os
  path = os.path.join(folder, file)
  ```

- ❌ **Async / await / asyncio**

- ❌ **Type hints syntax** (PEP 526: `var: str = ""`) — IronPython ne les supporte pas

- ❌ **`collections.abc`** au lieu de `collections` (IronPython 2.7 a `collections`)

- ❌ **`functools.lru_cache`** peut varier, tester avant utilisation

### 2.3 Contraintes Bibliothèques Python

> ⚠️ **Ces bibliothèques NE FONCTIONNENT PAS dans IronPython 2.7 — ne PAS les importer**

- ❌ `pandas` — incompatible
- ❌ `openpyxl` — incompatible
- ❌ `python-pptx` — incompatible
- ❌ `xlrd` / `xlwt` — incompatible (certaines versions peuvent fonctionner via IronClaw, mais non testé)

### 2.4 Contraintes Bibliothèque .NET Disponibles

```
BONNE PRATIQUE → Libraries .NET directement accessibles depuis IronPython
```

- ✅ `import clr` puis `clr.AddReference("Microsoft.Office.Interop.PowerPoint")`
- ✅ `import csv` — fonctionne nativement ✅
- ✅ `System.Windows.Forms` — Windows Forms via .NET
- ✅ `System.Drawing` — pour la manipulation d'images

### 2.5 Méthode CSV (choisie pour les données tabulaires)

L'export de données suit le chemin suivant :

```
Ansys Mechanical API → Extraction данных → import csv (fichier .csv sur disque) 
→ Lecture CSV dans PowerPoint/Excel via COM Interop (quand nécessaire)
→ Fichier CSV conservé comme archive à côté du rapport
```

Cette méthode :
- Fonctionne à 100% dans IronPython 2.7
- Donne un fichier de données autonome exploitable séparément
- Permet une relecture/modification facile

---

## 3. Architecture & Organisation du Projet

### 3.1 État actuel

L'utilisateur a déjà écrit une **base de code partielle** (majorité des actions de base) en самостоятельно. Claude Code est sollicité pour :

1. **Actions complexes** non encore implémentées
2. **Cohérence globale** et refactoring
3. **Maintenabilité** du code existant

### 3.2 Structure modulaire recommandée

```
📁 Projet AnsysReportGenerator/
│
├── 📄 AnsysReportGenerator.py        # Script principal — point d'entrée, GUI, orchestration
│   ├── Classe MainWindow (Windows Forms)
│   ├── Constantes et paths
│   └── Orchestration des modules
│
├── 📄 export_image.py                # Module extraction d'images (API Mechanical)
│   ├── export_view_screenshot()
│   ├── export_named_view()
│   └── setup_graphics_settings()
│
├── 📄 export_data.py                 # Module extraction de données (CSV)
│   ├── extract_result_table()
│   ├── extract_probe_results()
│   ├── extract_solution_info()
│   └── extract_contact_results()
│
├── 📄 pptx_manager.py                # Module PowerPoint (COM Interop)
│   ├── open_template()
│   ├── populate_slide_from_template()
│   ├── insert_image_into_slide()
│   ├── insert_table_from_csv()
│   └── save_presentation()
│
├── 📄 csv_utils.py                   # Utilitaires CSV
│   ├── write_data_to_csv()
│   └── read_csv_for_display()
│
├── 📄 ui_components.py               # Composants UI réutilisables (Windows Forms)
│   ├── FilePicker()
│   ├── StepSelector()
│   └── ResultSelector()
│
├── 📄 constants.py                   # Constantes, paths, settings globaux
│
└── 📁 data/                          # Dossier contenant les CSV générés
    ├── geometry_context.csv
    ├── mesh_context.csv
    ├── results_loadcase_001.csv
    └── ...
```

> **Note :** Si l'utilisateur préfère un **fichier unique** (tout dans un seul script), la structure interne doit suivre des regions / classes bien délimitées avec des commentaires clairs. Dans les deux cas, le fichier `CLAUDE.md` reste le point de référence unique.

### 3.3 Principes architecturaux

- **Réutilisabilité** : chaque fonction fait une chose, et une seule
- **Clarté** : nommage explicite en anglais, docstrings détaillés
- **Maintenabilité** : code lisible par un tiers sans connaissance préalable du projet
- **Pas de dépendances externes** : tout doit fonctionner dans le contexte IronPython d'Ansys Mechanical

---

## 4. Conventions de Code

### 4.1 Langue

| Élément | Langue |
|---------|--------|
| Noms de variables | English |
| Noms de fonctions/méthodes | English |
| Noms de classes | English (PascalCase) |
| Commentaires | English |
| Docstrings | English |
| Messages GUI | English (interface utilisateur, visible par l'ingénieur — bascule décidée le 2026-08-19 ; documentation en français conservée dans README.md, voir aussi README_EN.md) |

### 4.2 Style de codage IronPython 2.7

```python
# === CLASS DEFINITION ===
class ResultExporter(object):
    """
    Handles extraction and export of result data from Ansys Mechanical.
    
    Attributes:
        data_folder (str): Path to the folder where CSV files are stored.
        csv_delimiter (str): Delimiter used in CSV files (default: ';').
    """
    
    def __init__(self, data_folder):
        self.data_folder = data_folder
        self.csv_delimiter = ";"
    
    def extract_probe_results(self, selection, steps):
        """
        Extract probe results for selected load cases.
        
        Args:
            selection (list): List of result items to extract.
            steps (list): List of load case step IDs to include.
        
        Returns:
            bool: True if extraction succeeded, False otherwise.
        """
        result_data = []
        for step_id in steps:
            step_label = "Step_{}".format(step_id)
            # TODO: integrate with Mechanical API
            pass
        return result_data


# === FUNCTION DEFINITION ===
def export_view_screenshot(view_name, output_path):
    """
    Export the current graphics view as an image file.
    
    Args:
        view_name (str): Name or ID of the view to export.
        output_path (str): Full path including filename and extension (.png).
    
    Returns:
        bool: True if export succeeded.
    """
    try:
        graphics = DataModel.Project.Model.Analyses[0].Graphics
        export_settings = GraphicsImageExportSettings()
        export_settings.Resolution = GraphicsResolution.EffectiveResolution
        graphics.ExportImage(output_path, export_settings)
        return True
    except Exception as e:
        print "Error exporting view: {}".format(e)
        return False


# === STRING FORMATTING ===
# ALWAYS use .format(), never f-strings
project_name = "MyAnalysis"
output_file = "result_{}_export.csv".format(project_name)

# === PRINT STATEMENT ===
print "Starting export for step {}".format(step_id)
print "Done."
```

### 4.3 Conventions de nommage

| Type | Convention | Exemple |
|------|------------|---------|
| Variable | snake_case | `result_data`, `image_path` |
| Constante | UPPER_SNAKE_CASE | `MAX_STEPS`, `DEFAULT_FOLDER` |
| Fonction | snake_case | `export_view_screenshot()` |
| Méthode de classe | snake_case | `extract_probe_results()` |
| Classe | PascalCase | `MainWindow`, `ResultExporter` |
| Enum .NET | exact comme dans .NET | `GraphicsResolution.MediumResolution` |

### 4.4 Docstrings obligatoires

Chaque fonction/méthode publique doit avoir un docstring contenant :
- **Description** courte de ce que fait la fonction
- **Args** : paramètres avec types
- **Returns** : type et description de la valeur de retour
- **Raises** : exceptions potentielles (si applicable)

---

## 5. Workflow de Données

### 5.1 Pipeline complet

```
┌─────────────────────────────────────────────────────────────┐
│                 Ansys Mechanical (IronPython)                │
│                                                               │
│  ┌─────────────┐   ┌──────────────┐   ┌─────────────────┐  │
│  │ API Extract │ → │ CSV Writer   │ → │  Fichiers CSV   │  │
│  │ (images +   │   │ (import csv) │   │  (data/ folder) │  │
│  │  tables)    │   │              │   │                 │  │
│  └─────────────┘   └──────────────┘   └────────┬────────┘  │
│                                                 │           │
│                                                 ▼           │
│  ┌───────────────────────────────────────────────────────┐  │
│  │                  PowerPoint (COM Interop)              │  │
│  │  ┌─────────────┐   ┌──────────────┐   ┌────────────┐  │  │
│  │  │ Open        │ → │ Insert       │ → │ Populate   │  │  │
│  │  │ Template    │   │ Image from   │   │ Table from │  │  │
│  │  │             │   │ CSV path     │   │ CSV        │  │  │
│  │  └─────────────┘   └──────────────┘   └────────────┘  │  │
│  └───────────────────────────────────────────────────────┘  │
│                                                               │
│  ┌───────────────────────────────────────────────────────┐  │
│  │                  Excel (COM Interop)                   │  │
│  │  Generate separate data workbook for archival purposes  │  │
│  └───────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### 5.2 Séparation des préoccupations

1. **Extraction** (Ansys API → CSV/image) : fait partie du module `export_data.py` / `export_image.py`
2. **Stockage CSV** : fichiers `.csv` écrits sur le disque dans un dossier `data/` — ces fichiers sont **conservés** et servent d'archive
3. **Population PowerPoint** : lecture des CSV et insertion dans le template — module `pptx_manager.py`
4. **Excel archival** : génération d'un classeur Excel séparé à partir des CSV — module `excel_manager.py`

### 5.3 Chemins par défaut

```python
# Constantes à définir dans constants.py
TEMPLATE_PATH = r"C:\Users\<username>\VsCode\template\Master Template Perso.pptx"
OUTPUT_FOLDER = r"C:\Users\<username>\AnsysReports"
DATA_FOLDER = r"C:\Users\<username>\AnsysReports\data"
DEFAULT_IMAGE_FORMAT = ".png"
DEFAULT_CSV_DELIMITER = ";"
```

> ⚠️ Ces chemins sont des exemples — à confirmer avec l'utilisateur.

---

## 6. Spécification Fonctionnelle de l'Interface (3 Parties)

L'interface Windows Forms doit permettre la création de slides PowerPoint en **3 phases séquentielles** :

### PART 1 — Slides de Contexte (Contexte de l'analyse)

#### 1.1 Slide Géométrie
- Afficher la vue principale (une image)
- Afficher 4 vues individuelles (idealement)
- Bouton **Confirm** pour valider et passer à l'étape suivante

#### 1.2 Slide Maillage (Mesh)
- Afficher la vue principale du maillage
- Afficher 4 vues individuelles (idealement)
- Bouton **Confirm**

#### 1.3 Sélection des Conditions aux Limites (Boundary Conditions)
- Liste des BC actives dans le modèle
- Sélection des BC à afficher
- Sélection de la vue à associer
- Bouton **Confirm**

#### 1.4 Sélection Bolt Pretension
- Liste des Bolt Pretension définies
- Sélection de celles à afficher
- Sélection de la vue associée
- Bouton **Confirm**

#### 1.5 Slide Contexte Contact (table)
- Extraire la table des contacts (via API)
- L'afficher pour confirmation
- Exporter en CSV
- Bouton **Confirm**

#### 1.6 Slide Contexte Analyse (image + table paramètres)
- Image de l'analyse (vue d'ensemble)
- Table des paramètres :
  - **Steps** (liste des load cases/steps)
  - **Global Analysis Settings** (paramètres globaux : time range, solver type, etc.)
- Bouton **Confirm**

---

### PART 2 — Slides de Résultats

#### 2.1 Contact Tool Focus
- Sélection des **Contact Tools** à afficher (liste checkbox)
- Sélection de la vue associée
- Sélection de la partie de l'arbre :
  - `Contacts` (définition) ou
  - `Solution` (résultats)
- Bouton **Confirm**

#### 2.2 Résultats Globaux d'Analyse (Global Analysis Result)
- Sélection des **load cases** à afficher (multi-select)
- Méthode d'affichage :
  - **Condensed slide** : tous les load cases sur une seule slide (tableau condensé)
  - **Individual slides** : une slide par load case
- Image + table combinées
- Bouton **Confirm**

#### 2.3 Résultats Probe (principalement tables)
- Sélection des **load cases** (multi-select)
- Méthode d'affichage (condensed vs individual)
- Images de contexte (si applicable)
- Table principale des résultats
- Bouton **Confirm**

#### 2.4 Résultats Solution Information (extraction du graphe/chart)
- Détection automatique du graphique Solution Information
- Extraction de l'image du graphique
- Insertion dans la slide
- Bouton **Confirm**

#### 2.5 Résultats Bolt Tool
- Sélection des **load cases** (multi-select)
- Méthode d'affichage (condensed vs individual)
- Table des résultats bolt
- Bouton **Confirm**

#### 2.6 Available Focus Results
- Liste des résultats "Focus" disponibles dans l'arbre
- Sélection des focus à afficher
- Vue associée
- Bouton **Confirm**

---

### PART 3 — Gestion des Graphiques (Options transversales)

#### 3.1 Transparence
- Activation/désactivation d'effets de transparence sur le modèle
- Sélection du niveau de transparence (slider 0-100%)

#### 3.2 Section Cuts (Coupes)
- Gestion des coupes de section pour révéler l'intérieur du modèle
- Ajout/suppression de coupes
- Orientation de la coupe (plan XYZ, custom)

#### 3.3 Zoom par Partie
- Zoom sur des zones spécifiques :
  - Géométrie
  - Maillage
  - Résultats

---

## 7. Glossaire Métier Ansys

| Terme | Définition |
|-------|------------|
| **Step** | Un "step" est une étape de chargement dans une analyse. Une analyse peut contenir plusieurs steps (ex: Step 1 = preload, Step 2 = working load). Chaque step = un load case. |
| **Load Case** | Cas de charge — résultat d'un step. Synonyme de "solution step". |
| **Boundary Conditions (BC)** | Conditions aux limites — contraintes ou déplacements imposés (encastrement, pression, force, etc.). |
| **Bolt Pretension** | Précontrainte boulon — condition spécifique pour les connexions boulonnées (force de serrage initiale). |
| **Contact Tool** | Outil d'analyse de contact — permet d'évaluer la qualité du contact entre faces (gap, pressure, slip, etc.). |
| **Probe** | Sonde — outil permettant d'extraire la valeur d'un résultat (contrainte, déplacement) en un point ou une sélection de nodes/eléments. |
| **Solution Information** | Informations de solution — données de convergence du solveur (graphique de convergence, nombre d'itérations, résiduel, etc.). |
| **Bolt Tool** | Outil spécifique pour l'analyse des boulons — donne l'effort dans les bolts (axial, shear, etc.). |
| **Focus** | Résultat "Focus" — résultat agrégé filtré par une sélection (ex: focus sur une pièce). |
| **Named View** | Vue nommée — vue sauvegardée avec un nom (ex: "Vue Iso", "Vue Face", etc.). Utilisable pour l'export d'images cohérent. |
| **Graphics Export** | Export d'images depuis la vue graphique d'Ansys Mechanical via l'API `Graphics.ExportImage()`. |
| **GraphicsImageExportSettings** | Classe .NET permettant de configurer les paramètres d'export (résolution, format, etc.). |

---

## 8. Pièges & Bonnes Pratiques Spécifiques

### 8.1 IronPython 2.7

| Piège | Conséquence | Bonne pratique |
|--------|-------------|----------------|
| Utiliser `print("x")` au lieu de `print "x"` | Erreur de syntaxe | Vérifier toujours la syntaxe Python 2 |
| Utiliser des f-strings | Erreur de syntaxe | Utiliser `.format()` uniquement |
| Importer pandas/openpyxl | ImportError à runtime | Ne pas utiliser ces libs, passer par CSV + COM |
| Utiliser `pathlib` | ImportError | Utiliser `os.path.join()` |
| Type hints | SyntaxError | Ne pas utiliser |

### 8.2 COM Interop — PowerPoint / Excel

| Piège | Conséquence | Bonne pratique |
|--------|-------------|----------------|
| Ne pas libérer les objets COM | Mémoire qui s'accumule, Excel bloqué en arrière-plan | Toujours appeler `.Quit()` et `del` sur les objets |
| Utiliser des chemins relatifs | Fichier non trouvé | Utiliser des chemins absolus ou construir le chemin avec `os.path` |
| Modifier le template original | Template corrompu | Toujours travailler sur une **copie** du template |
| Tableaux mal dimensionnés | Texte tronqué ou chevauchement | Calculer dynamiquement la taille des tableaux en fonction du nombre de lignes/colonnes |

**Pattern COM cleanup :**
```python
try:
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = 1
    # ... operations ...
finally:
    if ppt_app:
        ppt_app.Quit()
        del ppt_app
```

### 8.3 API Mechanical — Export d'Images

| Piège | Conséquence | Bonne pratique |
|--------|-------------|----------------|
| Vue non sélectionnée avant export | Mauvaise vue exportée | S'assurer que la vue/graphique est actif avant export |
| Résolution trop haute | Fichier énorme, plantage | Commencer avec `GraphicsResolution.MediumResolution` |
| Format non spécifié | Comportement inattendu | Toujours specify `GraphicsImageExportFormat.Png` ou `.Jpeg` |

**Pattern export image :**
```python
def export_mechanical_view(analysis, output_path):
    graphics = analysis.Graphics
    settings = GraphicsImageExportSettings()
    settings.Resolution = GraphicsResolution.MediumResolution
    settings.BackgroundColor = GraphicsBackgroundColor.WhiteBackground
    graphics.ExportImage(output_path, settings)
```

### 8.4 Template PowerPoint Corporate

| Piège | Conséquence | Bonne pratique |
|--------|-------------|----------------|
| Modifier le template original | Perte du template pour les prochaines sessions | **Copier le template** vers un nouveau fichier avant modification |
| Placeholder mal orthographié | Slide non peuplée | Documente les noms exacts des placeholders dans le fichier constants |
| Layouts mal identifiés | Erreur à l'exécution | Identifier les IDs de layout dans le template avant de coder |

**Identification des layouts dans le template :**
```python
def list_template_layouts(ppt_app, template_path):
    pres = ppt_app.Presentations.Open(template_path)
    for i, layout in enumerate(pres.SlideMaster.CustomLayouts):
        print "Layout {}: {}".format(i, layout.Name)
    pres.Close()
```

### 8.5 Bonnes pratiques générales

- **Logging** : ajouter des `print` pour le debug dans la console Ansys (les `print` apparaissent dans la console Mechanical)
- **Gestion d'erreurs** : try/except autour de chaque opération risquée (accès COM, export)
- **Chemins** : toujours vérifier l'existence d'un fichier/répertoire avant de l'utiliser (`os.path.exists()`)
- **Valeurs par défaut** : fournir des valeurs par défaut pour tous les paramètres optionnels
- **Pas de code hardcodé** : toute valeur susceptible de changer doit être dans `constants.py`

---

## 9. Ressources & Chemins

### 9.1 Template PowerPoint Corporate

```
Chemin exact : .../VsCode/template/Master Template Persi.pptx
```

> ⚠️ **À CONFIRMER PAR L'UTILISATEUR** : chemin exact complet (lettre de lecteur, structure de dossiers).

Le template contient les layouts corporate (logo, titre, zones de contenu, etc.). Tous les slides générés doivent respecter ce template.

### 9.2 Chemins à confirmer

| Ressource | Chemin estimé | Statut |
|-----------|---------------|--------|
| Template PPTX | `.../VsCode/template/Master Template Persi.pptx` | ⚠️ À confirmer |
| Dossier de sortie | `.../AnsysReports/` | ⚠️ À définir |
| Dossier data (CSV) | `.../AnsysReports/data/` | ⚠️ À créer par le script |
| Dossier temporaire | `%TEMP%\AnsysReportGen\` | ✅ OK (Windows temp) |

### 9.3 Réferences .NET COM Interop

```
# PowerPoint
Microsoft.Office.Interop.PowerPoint (via COM)

# Excel  
Microsoft.Office.Interop.Excel (via COM)

# Windows Forms
System.Windows.Forms (référence .NET native)
System.Drawing (pour ImageList, icons, etc.)
```

### 9.4 API Ansys Mechanical principales

```python
# Accès au modèle
model = DataModel.Project.Model

# Accès aux analyses
analyses = model.Analyses

# Accès à laGraphics
graphics = analyses[0].Graphics

# Export d'image
GraphicsImageExportSettings()
graphics.ExportImage(path, settings)

# Résultats
results = analyses[0].Solution
```

---

## 10. TODO — Points à Compléter par l'Utilisateur

### 10.1 Points ouverts

| # | Question | Importance |
|---|----------|------------|
| 1 | Quel est le **chemin exact complet** du template PowerPoint ? | 🔴 Critique |
| 2 | Quelle est la **version d'Office** installée (2016, 2019, 365) ? Permet d'ajuster les IDs de méthodes COM. | 🔴 Critique |
| 3 | Le code existant de l'utilisateur est-il dans un **dépôt / dossier spécifique** à importer dans le projet ? | 🟡 Important |
| 4 | Quelle est la **convention de nommage** des variables/fonctions déjà utilisée dans le code existant ? | 🟡 Important |
| 5 | Y a-t-il des **noms de placeholders** PowerPoint déjà identifiés dans le template ? | 🟡 Important |
| 6 | Quelle est l'**arborescence exacte** du dossier VsCode / projet ? | 🟡 Important |
| 7 | Faut-il générer aussi un **rapport PDF** en sortie, ou uniquement le PPTX ? | 🟢 Optionnel |
| 8 | Faut-il supporter l'export vers **plusieurs formats d'image** (PNG, JPEG, TIFF) ? | 🟢 Optionnel |
| 9 | Y a-t-il des **analyses spécifiques** (modal, harmonic, etc.) qui nécessitent un traitement différent ? | 🟢 Optionnel |

### 10.2 Ordre de développement suggéré

```
Phase 1 : Base — Export d'image + CSV basique
Phase 2 : GUI — Interface Windows Forms complète
Phase 3 : PowerPoint — Population du template
Phase 4 : Intégration — Orchestration des 3 parties
Phase 5 : Validation — Tests dans Ansys Mechanical réel
```

---

## Résumé pour Claude Code

Ce projet est un script **IronPython 2.7** exécuté dans **Ansys Mechanical 2025 R2**. Il utilise **Windows Forms** pour l'interface, **COM Interop** pour piloter PowerPoint/Excel, et **l'API native Mechanical** pour extraire images et données. La méthode de données tabulaires est le **CSV** (pas de pandas/openpyxl).

Le code doit être :
- **English** pour les noms de fonctions, classes, variables, commentaires
- **Réutilisable** et **maintenable** par un tiers
- **Sans dépendances externes** (hors libs standard IronPython + COM .NET)

Le template PowerPoint corporate est dans `.../VsCode/template/` et doit être **copié avant modification** (ne jamais toucher à l'original).

**Premier pas suggéré** : Identifier la structure exacte du code existant de l'utilisateur, puis créer un plan de travail pour intégrer les modules manquants de manière cohérente.

---

*Document généré pour servir de contexte à Claude Code. Mettre à jour ce document au fur et à mesure de l'évolution du projet.*