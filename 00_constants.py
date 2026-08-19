# 00_constants.py : constantes globales et helpers generiques (chemins, fichiers). A executer EN PREMIER dans la console de scripting Mechanical.

import os
import re
import shutil

# === Chemins racine, calcules a partir de PROJECT_DIR (deja defini par AnsysReportGenerator_WPF.py avant l'execfile() de ce fichier, = le dossier "Report Generator" du projet Ansys) ===
# Aucun chemin code en dur pour un poste/projet particulier : PROJECT_DIR est localise via l'API
# Ansys elle-meme (ExtAPI.DataModel.Project.ProjectDirectory), pas via l'emplacement du script -
# un nouveau projet qui regroupe tous les .py, le .xaml et le template dans un dossier "Report
# Generator" a cote de "user_files" fonctionne directement, et les dossiers de stockage
# ci-dessous sont crees automatiquement a la premiere execution s'ils n'existent pas encore
# (voir les ensure_folder_exists() en bas de ce fichier).
DATA_ROOT = os.path.join(PROJECT_DIR, "data")

IMAGE_EXPORT_FOLDER = os.path.join(DATA_ROOT, "image_export")
CSV_EXPORT_FOLDER = os.path.join(DATA_ROOT, "csv_export")
EXPORT_3D_FOLDER = os.path.join(DATA_ROOT, "export_3D")

    # Volontairement HORS de DATA_ROOT (donc jamais concerne par le nettoyage de l'onglet Fichiers,
    # ni cree automatiquement ci-dessous) : "user_files" est le dossier standard du projet Ansys, a
    # cote de "Report Generator" (voir PROJECT_DIR ci-dessus) - les legendes y sont deposees et
    # entretenues manuellement par l'ingenieur, ce script ne fait que les LIRE, jamais les generer.
PROJECT_ROOT = os.path.dirname(PROJECT_DIR)
LEGEND_FOLDER = os.path.join(PROJECT_ROOT, "user_files", "legend")

    # Dossier de la copie de travail du template (voir PPTReportBuilder dans 03_ppt_utils.py) : le template original n'est jamais ouvert directement, pour ne jamais risquer d'etre ecrase par un Ctrl+S accidentel.
REPORT_OUTPUT_FOLDER = os.path.join(DATA_ROOT, "reports")

    # Directement dans PROJECT_DIR (structure a plat, pas de sous-dossier "templates") : contrairement aux dossiers ci-dessus, ce fichier ne peut pas etre cree automatiquement s'il manque (voir l'avertissement plus bas).
TEMPLATE_PATH = os.path.join(PROJECT_DIR, "Master Template_def.pptx")

    # Logo entreprise affiche en haut a droite de la fenetre (voir ReportGeneratorApp._find_controls,
    # imgLogo dans le XAML) : comme TEMPLATE_PATH, ne peut pas etre cree automatiquement s'il manque.
LOGO_PATH = os.path.join(PROJECT_DIR, "logo", "Liebherr-Emblem.png")


# === Index des layouts personnalises du template PowerPoint ===
LAYOUT_IMAGE_TABLE = 10    # title[2] / subtitle[4] / image[3] / table[1] / comment[8]
LAYOUT_TABLE_ONLY = 8      # title[1] / subtitle[3] / table[2]
LAYOUT_MESH_MULTI = 11     # images[5,6,7,8] (haut) / tables[9,10,11,12] (bas) -- indices sur la SLIDE generee, pas sur le layout (voir MESH_MULTI_*_SHAPE_INDICES ci-dessous)

DEFAULT_IMAGE_WIDTH = 1920
DEFAULT_IMAGE_HEIGHT = 1920

# === Garde-fou d'affichage des tableaux dans PowerPoint ===
    # Le CSV est toujours exporte quelle que soit sa taille ; seule son insertion en table PowerPoint est bloquee au-dela de ces limites (tableau illisible une fois insere).
MAX_TABLE_ROWS = 50
MAX_TABLE_COLUMNS = 50

# === Mesh par piece isolee (slide multi-image, voir LAYOUT_MESH_MULTI) ===
    # Le layout 11 du template ("Disposition personnalisee") contient dans SlideMaster.CustomLayouts
    # une forme Table supplementaire (pas un placeholder) qui n'est PAS heritee par les slides creees
    # a partir de ce layout : sur layout.Shapes elle occupe l'index 5 et decale tout ce qui suit
    # (images en 6-9, tables en 10-13), mais sur la slide reellement generee (report.presentation.Slides.AddSlide),
    # cette forme est absente et tout remonte d'un cran (images en 5-8, tables en 9-12). Les indices
    # ci-dessous sont ceux vus sur la SLIDE (ce que le code manipule reellement), pas sur le layout.
MESH_MULTI_IMAGE_SHAPE_INDICES = [5, 6, 7, 8]
MESH_MULTI_TABLE_SHAPE_INDICES = [9, 10, 11, 12]
MAX_MESH_MULTI_BODIES = 4  # nombre d'emplacements image/table disponibles sur ce layout


def ensure_folder_exists(folder_path):
    """
    Fait : cree le dossier folder_path (et ses parents) s'il n'existe pas deja.
    Depend de : os.path.exists / os.makedirs.
    Retourne : rien (effet de bord sur le systeme de fichiers).
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def safe_file_name(name):
    """
    Fait : remplace les caracteres interdits dans un nom de fichier Windows (dont "/" et "\\") par un underscore.
    Depend de : le module re (regex).
    Retourne : str, le nom nettoye, utilisable tel quel dans un chemin de fichier.
    """
    # Sans ce nettoyage, un nom Mechanical type "Part/Solid" cree un faux sous-dossier ("Mesh_Part\Solid.csv") et fait echouer l'ecriture.
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "object"


def get_unique_file_path(folder, base_name, extension):
    """
    Fait : construit un chemin de fichier qui n'entre pas en collision avec un fichier existant, en ajoutant un suffixe incremental si besoin.
    Depend de : os.path.exists / os.path.join.
    Retourne : str, un chemin absolu garanti inexistant au moment de l'appel.
    """
    path = os.path.join(folder, base_name + extension)
    counter = 1
    while os.path.exists(path):
        path = os.path.join(folder, base_name + "_" + str(counter) + extension)
        counter += 1
    return path


def list_data_cleanup_folders():
    """
    Fait : liste les sous-dossiers directs de DATA_ROOT proposables au nettoyage (onglet Fichiers),
    en excluant le dossier des legendes (jamais concerne par le nettoyage - contrairement au reste,
    ce ne sont pas des exports mais des fichiers de configuration reutilises d'une generation a l'autre).
    Depend de : DATA_ROOT, LEGEND_FOLDER, os.listdir/os.path.isdir.
    Retourne : list de tuples (nom_affiche, chemin_absolu), tries par nom (vide si DATA_ROOT n'existe pas).
    """
    if not os.path.isdir(DATA_ROOT):
        return []
    legend_name = os.path.basename(os.path.normpath(LEGEND_FOLDER))
    folders = []
    for name in os.listdir(DATA_ROOT):
        path = os.path.join(DATA_ROOT, name)
        if os.path.isdir(path) and name != legend_name:
            folders.append((name, path))
    return sorted(folders, key=lambda item: item[0].lower())


def get_folder_stats(folder_path):
    """
    Fait : calcule la taille totale et le nombre de fichiers d'un dossier (recursif, sous-dossiers inclus).
    Depend de : os.walk, os.path.getsize.
    Retourne : tuple (total_size_bytes, file_count) - (0, 0) si le dossier n'existe pas.
    """
    total_size = 0
    file_count = 0
    if not os.path.isdir(folder_path):
        return total_size, file_count
    for root, _dirs, files in os.walk(folder_path):
        for name in files:
            try:
                total_size += os.path.getsize(os.path.join(root, name))
                file_count += 1
            except OSError:
                pass
    return total_size, file_count


def format_folder_size(size_bytes):
    """
    Fait : formate une taille en octets en chaine lisible (bytes/KB/MB/GB).
    Depend de : rien (calcul pur).
    Retourne : str, la taille formatee (ex : "12.4 MB").
    """
    size = float(size_bytes)
    for unit in ("bytes", "KB", "MB"):
        if size < 1024.0:
            if unit == "bytes":
                return "{} {}".format(int(size), unit)
            return "{:.1f} {}".format(size, unit)
        size /= 1024.0
    return "{:.1f} GB".format(size)


def clear_folder_contents(folder_path):
    """
    Fait : supprime tout le contenu (fichiers et sous-dossiers) d'un dossier, sans supprimer le dossier lui-meme.
    Depend de : os.listdir, os.remove, shutil.rmtree.
    Retourne : rien (effet de bord sur le systeme de fichiers ; ne fait rien si le dossier n'existe pas).
    """
    if not os.path.isdir(folder_path):
        return
    for name in os.listdir(folder_path):
        path = os.path.join(folder_path, name)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception as e:
            print "Unable to delete: {} ({})".format(path, str(e))


def clean_cell_text(text):
    """
    Fait : normalise le texte d'une cellule du Tabular Data pane pour l'export CSV.
    Depend de : rien (traitement de chaine pur).
    Retourne : str, le texte nettoye ("" si l'entree etait None).
    """
    if text is None:
        return ""
    return text.replace("=", "").strip().rstrip(",").strip()


def to_csv_cell(value):
    """
    Fait : convertit une valeur quelconque (texte .NET unicode, nombre, None) en str encodee UTF-8 pour csv.writer.
    Depend de : le type unicode d'IronPython 2.7.
    Retourne : str encodee UTF-8 ("" si value est None).
    """
    # Certaines unites renvoyees par Mechanical (mm3, degre, micro...) contiennent des caracteres speciaux qui font planter l'ecriture si l'encodage n'est pas fixe explicitement.
    if value is None:
        return ""
    if isinstance(value, unicode):
        return value.encode("utf-8")
    return str(value)


# Premiere execution sur un nouveau projet : ces dossiers de stockage n'existent pas encore,
# on les cree ici une bonne fois pour toutes avant que le reste de l'application ne s'en serve.
# LEGEND_FOLDER n'en fait PAS partie : hors de DATA_ROOT, entretenu manuellement par l'ingenieur
# dans "user_files" (voir sa definition ci-dessus) - le creer automatiquement ici masquerait une
# vraie absence de legendes plutot que d'avertir l'utilisateur.
ensure_folder_exists(IMAGE_EXPORT_FOLDER)
ensure_folder_exists(CSV_EXPORT_FOLDER)
ensure_folder_exists(REPORT_OUTPUT_FOLDER)
ensure_folder_exists(EXPORT_3D_FOLDER)

# Le template ne peut pas etre cree automatiquement (fichier de contenu, pas juste un dossier) :
# on avertit seulement en console pour que l'utilisateur sache tout de suite pourquoi la
# generation echouerait, sans bloquer le chargement des modules suivants.
if not os.path.isfile(TEMPLATE_PATH):
    print "WARNING: PowerPoint template not found at the expected location: " + TEMPLATE_PATH

# Meme logique pour le dossier des legendes (voir LEGEND_FOLDER ci-dessus) : plus cree
# automatiquement, on avertit simplement si l'emplacement attendu dans "user_files" n'existe pas.
if not os.path.isdir(LEGEND_FOLDER):
    print "WARNING: legend folder not found at the expected location: " + LEGEND_FOLDER

# Meme logique pour le logo (voir LOGO_PATH ci-dessus) : absence non bloquante, juste un
# avertissement (l'emplacement imgLogo dans le XAML reste alors simplement vide).
if not os.path.isfile(LOGO_PATH):
    print "WARNING: logo not found at the expected location: " + LOGO_PATH
