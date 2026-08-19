# AnsysReportGenerator_WPF.py : point d'entree WPF pour l'application de generation de rapport. Charge les modules 00_constants.py a 05_interactive_slides.py (meme dossier que ce script) via execfile(), puis construit la fenetre depuis AnsysReportGenerator_WPF.xaml.

import os
import shutil
import xml.etree.ElementTree as ET

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference("System")

from System.Diagnostics import Process

# System.Windows.Forms/Color restent necessaires ici : 05_interactive_slides.py appelle
# SWF.Application.DoEvents() (_set_result_display_time) et utilise des Color nommees pour
# CURVE_COLOR_OPTIONS, meme si CE script batit son interface en WPF.
import System.Windows.Forms as SWF
from System.Drawing import Color

from System.IO import StreamReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader
from System.Windows import (
    Thickness, CornerRadius, GridLength, GridUnitType, TextTrimming, TextWrapping, TextAlignment,
    VerticalAlignment, HorizontalAlignment, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult,
    FontWeights, Point, Visibility
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, RowDefinition, StackPanel, Orientation, TextBlock, CheckBox, Button, TextBox,
    ComboBox, RadioButton, Slider, WrapPanel, Border, DockPanel, Dock, ScrollViewer, ScrollBarVisibility, Canvas
)
from System.Windows.Controls.Primitives import Popup, PlacementMode
from System.Windows.Shapes import Line
from System.Windows.Media import (
    SolidColorBrush, Color as WpfColor, VisualBrush, VisualTreeHelper, Brushes, LinearGradientBrush, GradientStop,
    PenLineCap
)
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Input import Key, MouseButtonState, Cursors
from System import Uri, UriKind


# --- SECTION 1 - Chargement des modules du projet (00 -> 05) ---
# execfile() execute chaque fichier dans le namespace global de ce script, comme si son
# contenu avait ete copie-colle dans la console a la suite des autres.

# PROJECT_DIR = dossier "Report Generator" du projet Ansys courant, qui regroupe TOUT : ce
# script, AnsysReportGenerator_WPF.xaml, les modules 00_constants.py -> 05_interactive_slides.py
# et le template PowerPoint (structure volontairement a plat, un seul dossier a placer a cote de
# "user_files" dans le repertoire de fichiers du projet Ansys).
#
# Localise via l'API Ansys elle-meme (ExtAPI.DataModel.Project.ProjectDirectory, le dossier
# "<Projet>_files" du projet Mechanical courant) plutot que via __file__/os.getcwd()/sys.argv :
# ces derniers se sont averes peu fiables selon la facon dont Mechanical execute le script (ils
# peuvent pointer vers un chemin propre a la session plutot que l'emplacement reel du script),
# alors qu'ExtAPI est garanti disponible (utilise partout ailleurs dans ce projet) et donne
# toujours le vrai repertoire du projet, quel que soit le mode de lancement du script.
try:
    _ansys_project_directory = ExtAPI.DataModel.Project.ProjectDirectory
except Exception as _ex:
    raise RuntimeError(
        "Impossible de lire ExtAPI.DataModel.Project.ProjectDirectory : {}. Le projet Ansys "
        "a-t-il ete enregistre au moins une fois ?".format(str(_ex))
    )

if not _ansys_project_directory:
    raise RuntimeError(
        "ExtAPI.DataModel.Project.ProjectDirectory est vide : enregistrez le projet Ansys avant "
        "de lancer ce script."
    )

PROJECT_DIR = os.path.join(_ansys_project_directory, "Report Generator")

if not os.path.isfile(os.path.join(PROJECT_DIR, "00_constants.py")):
    raise IOError(
        "Dossier 'Report Generator' introuvable ou incomplet : {}. Verifiez qu'il existe bien a "
        "cote du dossier 'user_files' du projet (ExtAPI.DataModel.Project.ProjectDirectory), et "
        "qu'il contient AnsysReportGenerator_WPF.py/.xaml, les modules 00_constants.py -> "
        "05_interactive_slides.py et le template PowerPoint.".format(PROJECT_DIR)
    )

_MODULE_FILES = [
    "00_constants.py",
    "01_data_export.py",
    "02_image_export.py",
    "03_ppt_utils.py",
    "04_slides.py",
    "05_interactive_slides.py",
]

for _module_file in _MODULE_FILES:
    _module_path = os.path.join(PROJECT_DIR, _module_file)
    if not os.path.exists(_module_path):
        raise IOError(
            "Module introuvable : {}. Verifiez PROJECT_DIR en haut de "
            "AnsysReportGenerator_WPF.py.".format(_module_path)
        )
    print "Chargement du module : " + _module_path
    execfile(_module_path)

print "Tous les modules sont charges."


# --- Chemins de fichiers modifiables depuis l'onglet "Fichiers" ---
# Valeurs d'origine de 00_constants.py, capturees une seule fois ici (avant toute
# modification depuis l'UI) pour que le bouton "Reinitialiser les chemins" puisse
# toujours y revenir. Cle = nom du global correspondant dans 00_constants.py, reaffecte
# directement via globals()[cle] = ... : tous les modules 00_constants.py -> 05_interactive_slides.py
# lisent ce meme global au moment de l'appel, aucun autre changement necessaire ailleurs.
FILE_PATH_SETTINGS = [
    ("TEMPLATE_PATH", "txtPathTemplate", "btnBrowseTemplate", "file"),
    ("IMAGE_EXPORT_FOLDER", "txtPathImages", "btnBrowseImages", "folder"),
    ("CSV_EXPORT_FOLDER", "txtPathCsv", "btnBrowseCsv", "folder"),
    ("LEGEND_FOLDER", "txtPathLegends", "btnBrowseLegends", "folder"),
    ("REPORT_OUTPUT_FOLDER", "txtPathReports", "btnBrowseReports", "folder"),
]

_DEFAULT_FILE_PATHS = dict((name, globals()[name]) for name, _, _, _ in FILE_PATH_SETTINGS)


# --- SECTION 2 - Helpers partages (couleurs de statut, recherche) ---

# --- Couleurs de statut des lignes de selection (3 etats) ---
# 3 etats bases sur la selection ET la configuration :
#   - non selectionnee (quel que soit son etat de configuration)
#   - selectionnee, pas encore configuree via "..."
#   - selectionnee ET configuree via "..."
# (voir les 3 brushes ci-dessous pour les couleurs exactes de chaque etat)

ROW_STATUS_NOT_SELECTED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xFF, 0xD0, 0x00))
ROW_STATUS_SELECTED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xBE, 0xE3, 0xDB))
ROW_STATUS_CONFIGURED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x7D, 0xCE, 0x82))


def _row_status_brush(row):
    """
    Fait : determine la couleur de fond d'une ligne de selection selon son etat coche/configure.
    Depend de : row.checkbox.IsChecked, row.row_config.configured, les 3 brushes ROW_STATUS_*.
    Retourne : SolidColorBrush, la couleur de fond a appliquer a row.border.
    """
    if not row.checkbox.IsChecked:
        return ROW_STATUS_NOT_SELECTED_BRUSH
    if row.row_config.configured:
        return ROW_STATUS_CONFIGURED_BRUSH
    return ROW_STATUS_SELECTED_BRUSH


def _general_slide_status_text(row_config):
    """
    Fait : construit le texte de statut affiche sous le titre des cartes Geometrie/Maillage
    ("Slides d'ensemble", onglet 01) - etat de configuration et vue effective.
    Depend de : row_config.configured/view_name.
    Retourne : str, ex. "a configurer - vue courante" ou "configure - vue=Vue ISO".
    """
    state = "configure" if row_config.configured else "a configurer"
    vue = "vue={}".format(row_config.view_name) if row_config.view_name else "vue courante"
    return "{} - {}".format(state, vue)


# --- Filtre par type de contact (section "Contacts a afficher") ---
# Base sur le PREFIXE du nom (pas sur contact.ContactType, l'API Ansys) : un contact renomme par
# l'ingenieur (nom "personnalise") doit tomber dans "Autres" meme si son type reste Frictional/Bonded
# cote solveur - c'est le nom affiche dans la liste, pas le type technique, que ce filtre trie.

CONTACTS_FILTER_OPTIONS = ["Tous", "Frictional", "Bonded", "Autres"]


def _classify_contact_name(name):
    """
    Fait : classe un nom de Contact Region selon son prefixe ("Frictional-...", "Bonded-...", ou personnalise).
    Depend de : rien (comparaison de chaine, insensible a la casse).
    Retourne : str, "Frictional"/"Bonded"/"Autres".
    """
    lowered = (name or "").strip().lower()
    if lowered.startswith("frictional"):
        return "Frictional"
    if lowered.startswith("bonded"):
        return "Bonded"
    return "Autres"


# --- Champs de recherche : texte indicatif grise ---

SEARCH_PLACEHOLDER = "Rechercher..."
SEARCH_PLACEHOLDER_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x79, 0x7E, 0x8A))  # meme gris que TextMutedBrush (xaml)
SEARCH_TEXT_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x00, 0x00))  # meme noir que TextPrimaryBrush (xaml)
SEARCH_BOX_DEFAULT_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xFA, 0xFB, 0xFC))
SEARCH_BOX_NO_MATCH_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xF8, 0xD9, 0xDC))
SEARCH_HIGHLIGHT_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x8D, 0xD5))

# --- Cartes d'apercu : couleur au survol (au lieu d'un zoom) ---

CARD_NORMAL_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xFF, 0xFF, 0xFF))
CARD_HOVER_BACKGROUND = SolidColorBrush(WpfColor.FromRgb(0xE7, 0xEC, 0xF8))  # meme bleu que HoverBrush (xaml)

# Largeur de la carte.
CARD_WIDTH = 340

# Conteneur de liste (voir _build_preview_list_container) : hauteur FIXE (pas un plafond) pour que
# toutes les cartes d'apercu partagent la meme taille, qu'elles aient 1 ou 50 elements coches -
# sinon une carte a peu d'elements (ex: Maillage) parait minuscule a cote d'une carte pleine (ex:
# Boundary Conditions). Au-dela de cette hauteur, la liste reste consultable via defilement,
# signale par un fondu (PREVIEW_LIST_FADE_HEIGHT) en bas.
PREVIEW_LIST_DEFAULT_HEIGHT = 130
PREVIEW_LIST_FADE_HEIGHT = 26
PREVIEW_LIST_BACKGROUND_COLOR = WpfColor.FromRgb(0xF1, 0xF2, 0xF5)
PREVIEW_LIST_BACKGROUND = SolidColorBrush(PREVIEW_LIST_BACKGROUND_COLOR)

# Fondu bas des listes de selection a cocher (onglets 01/02/03, cartes CardBorder) : meme principe
# que PREVIEW_LIST_FADE_HEIGHT ci-dessus (visible uniquement si la liste deborde reellement), mais
# applique via OpacityMask directement sur le ScrollViewer existant plutot qu'un Border de recouvrement
# separe - inutile ici puisque ces listes reposent toujours sur un fond CardBorder blanc uni (voir
# ReportGeneratorApp._attach_list_fade).
ITEM_LIST_FADE_HEIGHT = 26

# --- Ressources partagees avec AnsysReportGenerator_WPF.xaml ---
# Les champs des panneaux de configuration ("..." de chaque ligne, panneau lateral global -
# voir SECTION 4/5/5bis/6 et ReportGeneratorApp._open_config_panel) sont batis en code Python
# (Border/TextBlock/ComboBox... crees directement, pas charges depuis le XAML) et ne peuvent donc
# pas resoudre les {StaticResource ...} de la fenetre principale via le markup XAML. Plutot que de
# redefinir ces styles/couleurs a la main cote Python (source constatee de desynchronisation avec
# le .xaml), _shared_resources reference directement le MEME ResourceDictionary que la fenetre
# principale, assigne une seule fois dans ReportGeneratorApp.__init__ (une seule instance d'app
# par execution, voir SECTION 8) - fonctionne meme si ces controles rejoignent ensuite le meme
# arbre visuel que la fenetre principale (ce qui est le cas depuis le passage au panneau lateral).

_shared_resources = None  # assigne dans ReportGeneratorApp.__init__


def _make_field_label(text):
    """
    Fait : cree un TextBlock utilise comme etiquette de champ dans les dialogues de configuration.
    Depend de : _shared_resources["TextPrimaryBrush"].
    Retourne : TextBlock, l'etiquette prete a etre ajoutee au panneau.
    """
    label = TextBlock()
    label.Text = text
    label.FontWeight = FontWeights.SemiBold
    label.Foreground = _shared_resources["TextPrimaryBrush"]
    label.Margin = Thickness(0, 0, 0, 2)
    return label


def _themed_textbox():
    """
    Fait : cree un TextBox stylee comme le champ generique des boites de dialogue.
    Depend de : _shared_resources["DialogTextBox"] (x:Key defini dans le XAML).
    Retourne : TextBox, le champ stylise pret a l'emploi.
    """
    box = TextBox()
    box.Style = _shared_resources["DialogTextBox"]
    return box


def _themed_button(primary=False):
    """
    Fait : cree un Button style PrimaryButton (accent) ou SecondaryButton (neutre).
    Depend de : _shared_resources["PrimaryButton"/"SecondaryButton"], memes ressources que la fenetre principale.
    Retourne : Button, le bouton stylise pret a l'emploi.
    """
    btn = Button()
    btn.Style = _shared_resources["PrimaryButton" if primary else "SecondaryButton"]
    return btn


def _build_close_icon(size=10, thickness=1.4):
    """
    Fait : construit un petit "x" vectoriel (2 lignes croisees) a utiliser comme Button.Content
    pour un bouton de fermeture "x" - un simple TextBlock("x") n'est jamais parfaitement centre
    verticalement dans son cadre (metriques de police, ascender/descender), meme avec
    HorizontalAlignment/VerticalAlignment=Center sur le ContentPresenter.
    Depend de : Canvas/Line (System.Windows.Shapes), _shared_resources["TextPrimaryBrush"].
    Retourne : Canvas, l'icone prete a etre assignee a Button.Content (une instance neuve a chaque appel).
    """
    canvas = Canvas()
    canvas.Width = size
    canvas.Height = size

    for x1, y1, x2, y2 in ((0, 0, size, size), (size, 0, 0, size)):
        line = Line()
        line.X1 = x1
        line.Y1 = y1
        line.X2 = x2
        line.Y2 = y2
        line.Stroke = _shared_resources["TextPrimaryBrush"]
        line.StrokeThickness = thickness
        line.StrokeStartLineCap = PenLineCap.Round
        line.StrokeEndLineCap = PenLineCap.Round
        canvas.Children.Add(line)

    return canvas


# --- Messages console formates (etapes cles) ---
# Remplace les boites de dialogue bloquantes (MessageBox) pour les evenements de routine :
# une MessageBox.Show() bloque a la fois cette fenetre ET Mechanical jusqu'a sa fermeture
# manuelle, ce qui casse l'enchainement quand on genere plusieurs rapports a la suite.

CONSOLE_BANNER_WIDTH = 70


def _print_console_banner(title):
    """
    Fait : affiche un titre encadre dans la console Mechanical (etapes cles de la generation).
    Depend de : CONSOLE_BANNER_WIDTH.
    Retourne : rien (effet de bord : imprime dans la console).
    """
    border = "=" * CONSOLE_BANNER_WIDTH
    print border
    print title
    print border


# --- SECTION 3 - Ligne de selection WPF (case a cocher + nom + config) ---

class SectionRow(object):
    """
    Une ligne de selection WPF : case a cocher + nom de l'objet + bouton de configuration
    optionnel ("..."), liee a un row_config (SlideRowConfig / GeometryPartRowConfig /
    ContactRowConfig / MeshPartRowConfig / SolutionInfoRowConfig, voir 05_interactive_slides.py).

    Attributes:
        border (Border): Conteneur exterieur de la ligne (fond de statut, surlignage recherche).
        checkbox (CheckBox): Case a cocher "inclure cette slide".
        text_block (TextBlock): Nom affiche de l'objet.
        config_button (Button): Bouton "..." (None si la categorie n'a rien a configurer).
        row_config: Objet de configuration associe.
        display_name_func (callable): Fonction de texte d'apercu pour ce row_config.
        panel_kind (str): Etat du panneau lateral global a afficher pour ce row_config au clic sur
            "..." ("result"/"geometry_part"/"mesh_part"/"solution_info", None si aucune categorie).
            Voir ReportGeneratorApp._open_config_panel.
    """

    def __init__(self, border, checkbox, text_block, config_button, row_config,
                 display_name_func, panel_kind):
        """
        Fait : stocke les references des controles WPF et de la configuration associee a la ligne.
        Depend de : rien (simple assignation des parametres recus).
        Retourne : rien (initialise les attributs de self).
        """
        self.border = border
        self.checkbox = checkbox
        self.text_block = text_block
        self.config_button = config_button
        self.row_config = row_config
        self.display_name_func = display_name_func
        self.panel_kind = panel_kind


# --- SECTION 4 - Champs partages : vue / coupe / scale factor / steps ---
# Ces champs vivaient a l'origine dans 4 boites de dialogue modales ("..." de chaque ligne de
# selection). Elles ont ete remplacees par UN panneau lateral global (voir "PARAMETRES" dans le
# XAML, ReportGeneratorApp._open_config_panel et les methodes _on_config_panel_*) : plus aucune
# fenetre separee, seulement 4 "kinds" de contenu (result/geometry_part/mesh_part/solution_info)
# affiches tour a tour dans le meme panneau. Chaque paire de fonctions ci-dessous construit
# (_build_*) puis relit (_apply_*) un jeu de champs sur un `target` generique (voir
# _ConfigFieldsHolder) : ce decouplage est ce qui permet au meme code de servir a la fois au
# panneau lateral global ET au panneau de case de l'onglet "Slide combinee" (_build_row_config_fields
# uniquement, sans steps - un resultat fige par case).

def _build_row_config_fields(target, root, row_config, views, section_plane_labels, legend_names):
    """
    Fait : construit les champs de configuration graphique communs (vue, coupe, legende, apparence,
    scoping, scale factor - sans la section steps) et les ajoute a root. Partage par le panneau
    lateral global ("kind"="result", voir ReportGeneratorApp._open_config_panel) et par le panneau
    de case inline de l'onglet "Slide combinee" (un resultat fige par case, jamais de notion de
    step - voir ReportGeneratorApp._show_multi_result_editor).
    Depend de : _make_field_label, get_result_display_unit, CONTOUR_VIEW_OPTIONS/LEGEND_ORIENTATION_OPTIONS/SCOPING_DISPLAY_OPTIONS.
    Retourne : rien (pose sur target : cmb_view/cmb_section/cmb_legend/cmb_contour_view/cmb_legend_orientation/cmb_scoping_display/txt_scale).
    """
    root.Children.Add(_make_field_label("Vue (View Manager) :"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)

    root.Children.Add(_make_field_label("Coupe (Section Plane) :"))
    target.cmb_section = ComboBox()
    target.cmb_section.Margin = Thickness(0, 4, 0, 12)
    target.cmb_section.Items.Add(NO_SECTION_LABEL)
    for name in section_plane_labels:
        target.cmb_section.Items.Add(name)
    if row_config.section_name and row_config.section_name in section_plane_labels:
        target.cmb_section.SelectedItem = row_config.section_name
    else:
        target.cmb_section.SelectedIndex = 0
    root.Children.Add(target.cmb_section)

    # L'unite affichee ici est EXACTEMENT celle qui sera passee a ExtAPI.Graphics.ImportLegend()
    # lors de la generation : diagnostic visuel immediat, sans passer par la console.
    # force_evaluate=False : lecture indicative seule (pas de reevaluation couteuse du
    # resultat) pour que l'ouverture de cette fenetre reste instantanee ; la vraie application
    # de la legende (apply_legend_if_exists) reevalue toujours a fond le resultat.
    detected_unit = get_result_display_unit(row_config.obj, force_evaluate=False)
    lbl_unit = TextBlock()
    lbl_unit.Text = "Unite detectee pour ImportLegend : " + (detected_unit if detected_unit else "aucune")
    lbl_unit.Foreground = _shared_resources["DiagnosticLabelBrush"]
    lbl_unit.FontWeight = FontWeights.Bold
    lbl_unit.TextWrapping = TextWrapping.Wrap
    lbl_unit.Margin = Thickness(0, 0, 0, 6)
    root.Children.Add(lbl_unit)

    root.Children.Add(_make_field_label("Legende :"))
    target.cmb_legend = ComboBox()
    target.cmb_legend.Margin = Thickness(0, 4, 0, 12)
    target.cmb_legend.Items.Add(NO_LEGEND_LABEL)
    for name in legend_names:
        target.cmb_legend.Items.Add(name)
    if row_config.legend_name and row_config.legend_name in legend_names:
        target.cmb_legend.SelectedItem = row_config.legend_name
    else:
        target.cmb_legend.SelectedIndex = 0
    root.Children.Add(target.cmb_legend)

    root.Children.Add(_make_field_label("Affichage des couleurs (Contour View) :"))
    target.cmb_contour_view = ComboBox()
    target.cmb_contour_view.Margin = Thickness(0, 4, 0, 12)
    for label, value in CONTOUR_VIEW_OPTIONS:
        target.cmb_contour_view.Items.Add(label)
    target.cmb_contour_view.SelectedItem = contour_view_label(row_config.contour_view)
    root.Children.Add(target.cmb_contour_view)

    root.Children.Add(_make_field_label("Orientation de la legende :"))
    target.cmb_legend_orientation = ComboBox()
    target.cmb_legend_orientation.Margin = Thickness(0, 4, 0, 12)
    for label, value in LEGEND_ORIENTATION_OPTIONS:
        target.cmb_legend_orientation.Items.Add(label)
    target.cmb_legend_orientation.SelectedItem = legend_orientation_label(row_config.legend_orientation)
    root.Children.Add(target.cmb_legend_orientation)

    root.Children.Add(_make_field_label("Affichage du scoping :"))
    target.cmb_scoping_display = ComboBox()
    target.cmb_scoping_display.Margin = Thickness(0, 4, 0, 12)
    for label, value in SCOPING_DISPLAY_OPTIONS:
        target.cmb_scoping_display.Items.Add(label)
    target.cmb_scoping_display.SelectedItem = scoping_display_label(row_config.scoping_display)
    root.Children.Add(target.cmb_scoping_display)

    root.Children.Add(_make_field_label("Echelle de deformation :"))
    target.cmb_deformation_scale_mode = ComboBox()
    target.cmb_deformation_scale_mode.Margin = Thickness(0, 4, 0, 12)
    for label, value in DEFORMATION_SCALE_MODE_OPTIONS:
        target.cmb_deformation_scale_mode.Items.Add(label)
    target.cmb_deformation_scale_mode.SelectedItem = deformation_scale_mode_label(row_config.deformation_scale_mode)
    root.Children.Add(target.cmb_deformation_scale_mode)

    root.Children.Add(_make_field_label("Scale factor deformation (mode Manuel uniquement, defaut = 1) :"))
    target.txt_scale = _themed_textbox()
    target.txt_scale.Width = 100
    target.txt_scale.HorizontalAlignment = HorizontalAlignment.Left
    target.txt_scale.Margin = Thickness(0, 4, 0, 12)
    target.txt_scale.Text = "1" if row_config.scale_factor == 1.0 else str(row_config.scale_factor)
    root.Children.Add(target.txt_scale)


def _apply_row_config_fields(target, row_config):
    """
    Fait : lit les champs communs (vue/coupe/legende/apparence/scoping/scale factor) depuis target
    et les applique a row_config. Partage par le panneau lateral global (_on_config_panel_apply) et
    par le panneau de case inline de l'onglet "Slide combinee" (voir _build_row_config_fields).
    Depend de : target.cmb_view/cmb_section/cmb_legend/cmb_contour_view/cmb_legend_orientation/cmb_scoping_display/
        cmb_deformation_scale_mode/txt_scale.
    Retourne : rien (effet de bord sur row_config uniquement ; ne touche pas row_config.configured, ni les steps).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view

    selected_section = unicode(target.cmb_section.SelectedItem)
    row_config.section_name = None if selected_section == NO_SECTION_LABEL else selected_section

    selected_legend = unicode(target.cmb_legend.SelectedItem)
    row_config.legend_name = None if selected_legend == NO_LEGEND_LABEL else selected_legend

    row_config.contour_view = contour_view_from_label(unicode(target.cmb_contour_view.SelectedItem))
    row_config.legend_orientation = legend_orientation_from_label(unicode(target.cmb_legend_orientation.SelectedItem))
    row_config.scoping_display = scoping_display_from_label(unicode(target.cmb_scoping_display.SelectedItem))
    row_config.deformation_scale_mode = deformation_scale_mode_from_label(
        unicode(target.cmb_deformation_scale_mode.SelectedItem))

    # La valeur du champ n'est lue/validee qu'en mode Manuel : en mode Auto Scale x1/x2, le
    # multiplicateur applique est fixe (voir apply_scale_factor), ce champ est ignore - inutile
    # d'avertir sur une valeur invalide qui ne sera de toute facon pas utilisee.
    if row_config.deformation_scale_mode == "manual":
        try:
            scale_value = float(target.txt_scale.Text.strip().replace(",", "."))
            if scale_value <= 0:
                raise ValueError("Le scale factor doit etre positif.")
            row_config.scale_factor = scale_value
        except ValueError:
            row_config.scale_factor = 1.0
            MessageBox.Show("Valeur de scale factor invalide : la valeur par defaut (1) a ete appliquee.",
                             "Scale factor invalide", MessageBoxButton.OK, MessageBoxImage.Warning)


def _build_steps_section_fields(target, root, row_config, step_count):
    """
    Fait : construit la section "Loadcases" (steps + mode d'affichage individuel/combine) et l'ajoute a root.
    Depend de : row_config.selected_steps/step_display_mode, _shared_resources.
    Retourne : rien (pose sur target : step_checkboxes/radio_individual/radio_combined).
    """
    group = Border()
    group.BorderBrush = _shared_resources["CardBorderBrush"]
    group.BorderThickness = Thickness(1)
    group.CornerRadius = CornerRadius(0)
    group.Padding = Thickness(10)
    group.Margin = Thickness(0, 4, 0, 0)

    panel = StackPanel()
    group.Child = panel

    lbl_info = TextBlock()
    lbl_info.Text = "Loadcase disponible : {}".format(step_count)
    lbl_info.FontWeight = FontWeights.SemiBold
    lbl_info.Margin = Thickness(0, 0, 0, 6)
    panel.Children.Add(lbl_info)

    wrap = WrapPanel()
    selected_steps = row_config.selected_steps or []
    target.step_checkboxes = []
    for step in range(1, step_count + 1):
        cb = CheckBox()
        cb.Content = "Step {}".format(step)
        cb.Tag = step
        cb.IsChecked = step in selected_steps
        cb.Width = 90
        cb.Margin = Thickness(0, 2, 10, 2)
        wrap.Children.Add(cb)
        target.step_checkboxes.append(cb)
    panel.Children.Add(wrap)

    lbl_hint = TextBlock()
    lbl_hint.Text = "(aucun coche = etat actuel)"
    lbl_hint.FontSize = 11
    lbl_hint.Foreground = SEARCH_PLACEHOLDER_BRUSH
    lbl_hint.Margin = Thickness(0, 6, 0, 2)
    panel.Children.Add(lbl_hint)

    steps_buttons = StackPanel()
    steps_buttons.Orientation = Orientation.Horizontal
    steps_buttons.Margin = Thickness(0, 0, 0, 10)

    def _on_select_all_steps(sender, e):
        for cb in target.step_checkboxes:
            cb.IsChecked = True

    def _on_deselect_all_steps(sender, e):
        for cb in target.step_checkboxes:
            cb.IsChecked = False

    btn_select_all_steps = _themed_button()
    btn_select_all_steps.Content = "Tout selectionner"
    btn_select_all_steps.Padding = Thickness(8, 2, 8, 2)
    btn_select_all_steps.FontSize = 11
    btn_select_all_steps.Click += _on_select_all_steps
    steps_buttons.Children.Add(btn_select_all_steps)

    btn_deselect_all_steps = _themed_button()
    btn_deselect_all_steps.Content = "Deselectionner"
    btn_deselect_all_steps.Padding = Thickness(8, 2, 8, 2)
    btn_deselect_all_steps.FontSize = 11
    btn_deselect_all_steps.Margin = Thickness(6, 0, 0, 0)
    btn_deselect_all_steps.Click += _on_deselect_all_steps
    steps_buttons.Children.Add(btn_deselect_all_steps)

    panel.Children.Add(steps_buttons)

    target.radio_individual = RadioButton()
    target.radio_individual.Content = "Slides individuelles (1 par step)"
    target.radio_individual.GroupName = "StepDisplayMode"
    target.radio_individual.IsChecked = (row_config.step_display_mode != "combined")
    target.radio_individual.Margin = Thickness(0, 0, 0, 2)
    panel.Children.Add(target.radio_individual)

    target.radio_combined = RadioButton()
    target.radio_combined.Content = "Slide combinee (si template disponible)"
    target.radio_combined.GroupName = "StepDisplayMode"
    target.radio_combined.IsChecked = (row_config.step_display_mode == "combined")
    panel.Children.Add(target.radio_combined)

    root.Children.Add(group)


def _apply_steps_section_fields(target, row_config):
    """
    Fait : lit la selection de steps et le mode d'affichage depuis target et les applique a row_config.
    Depend de : target.step_checkboxes/radio_combined, get_multi_step_template, MULTI_STEP_SLIDE_TEMPLATES.
    Retourne : rien (effet de bord sur row_config.selected_steps/step_display_mode).
    """
    checked_steps = [cb.Tag for cb in target.step_checkboxes if cb.IsChecked]
    row_config.selected_steps = checked_steps if checked_steps else None

    if not checked_steps:
        row_config.step_display_mode = "individual"
    elif target.radio_combined.IsChecked:
        if get_multi_step_template(len(checked_steps)):
            row_config.step_display_mode = "combined"
        else:
            row_config.step_display_mode = "individual"
            supported = ", ".join(str(n) for n in sorted(MULTI_STEP_SLIDE_TEMPLATES.keys()))
            MessageBox.Show(
                "Aucun template de slide combinee n'existe pour {} step(s) (nombres supportes : {}). "
                "Des slides individuelles seront generees a la place.".format(len(checked_steps), supported),
                "Mode combine indisponible", MessageBoxButton.OK, MessageBoxImage.Warning
            )
    else:
        row_config.step_display_mode = "individual"


# --- SECTION 5 - Champs partages : vue / coupe / opacite (geometrie par piece) ---

def _build_geometry_part_fields(target, root, row_config, views, section_plane_labels):
    """
    Fait : construit les champs vue/coupe/opacite du contexte pour une piece isolee (geometrie) et les ajoute a root.
    Depend de : _make_field_label.
    Retourne : rien (pose sur target : cmb_view/cmb_section/slider_opacity/lbl_opacity_value).
    """
    root.Children.Add(_make_field_label("Vue (View Manager) :"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)

    root.Children.Add(_make_field_label("Coupe (Section Plane) :"))
    target.cmb_section = ComboBox()
    target.cmb_section.Margin = Thickness(0, 4, 0, 12)
    target.cmb_section.Items.Add(NO_SECTION_LABEL)
    for name in section_plane_labels:
        target.cmb_section.Items.Add(name)
    if row_config.section_name and row_config.section_name in section_plane_labels:
        target.cmb_section.SelectedItem = row_config.section_name
    else:
        target.cmb_section.SelectedIndex = 0
    root.Children.Add(target.cmb_section)

    root.Children.Add(_make_field_label("Opacite du contexte (autres pieces) :"))

    opacity_row = StackPanel()
    opacity_row.Orientation = Orientation.Horizontal
    opacity_row.Margin = Thickness(0, 4, 0, 12)

    target.slider_opacity = Slider()
    target.slider_opacity.Minimum = 0
    target.slider_opacity.Maximum = 100
    target.slider_opacity.TickFrequency = 10
    target.slider_opacity.Width = 260
    target.slider_opacity.Value = row_config.context_opacity_percent
    opacity_row.Children.Add(target.slider_opacity)

    target.lbl_opacity_value = TextBlock()
    target.lbl_opacity_value.Text = "{} %".format(row_config.context_opacity_percent)
    target.lbl_opacity_value.Margin = Thickness(10, 0, 0, 0)
    target.lbl_opacity_value.VerticalAlignment = VerticalAlignment.Center
    opacity_row.Children.Add(target.lbl_opacity_value)

    def _on_opacity_changed(sender, e):
        target.lbl_opacity_value.Text = "{} %".format(int(target.slider_opacity.Value))
    target.slider_opacity.ValueChanged += _on_opacity_changed

    root.Children.Add(opacity_row)


def _apply_geometry_part_fields(target, row_config):
    """
    Fait : lit les champs vue/coupe/opacite depuis target et les applique a row_config.
    Depend de : target.cmb_view/cmb_section/slider_opacity.
    Retourne : rien (effet de bord sur row_config).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view

    selected_section = unicode(target.cmb_section.SelectedItem)
    row_config.section_name = None if selected_section == NO_SECTION_LABEL else selected_section

    row_config.context_opacity_percent = int(target.slider_opacity.Value)


# --- SECTION 5bis - Champs partages : vue (geometrie par piece isolee, mesh) ---

def _build_mesh_part_fields(target, root, row_config, views):
    """
    Fait : construit le champ vue (View Manager) uniquement et l'ajoute a root - pas de coupe ni
    d'opacite (contrairement a _build_geometry_part_fields) : l'isolation se fait par masquage
    complet des autres corps (show_only_body), une coupe/opacite de contexte n'aurait pas de sens
    ici. Reutilise telle quelle pour Geometrie/Maillage/Contexte d'analyse (vue seule egalement).
    Depend de : _make_field_label.
    Retourne : rien (pose sur target : cmb_view).
    """
    root.Children.Add(_make_field_label("Vue (View Manager) :"))
    target.cmb_view = ComboBox()
    target.cmb_view.Margin = Thickness(0, 4, 0, 12)
    target.cmb_view.Items.Add(NO_VIEW_LABEL)
    for name in sorted(views.keys()):
        target.cmb_view.Items.Add(name)
    if row_config.view_name and row_config.view_name in views:
        target.cmb_view.SelectedItem = row_config.view_name
    else:
        target.cmb_view.SelectedIndex = 0
    root.Children.Add(target.cmb_view)


def _apply_mesh_part_fields(target, row_config):
    """
    Fait : lit le champ vue depuis target et l'applique a row_config.
    Depend de : target.cmb_view.
    Retourne : rien (effet de bord sur row_config).
    """
    selected_view = unicode(target.cmb_view.SelectedItem)
    row_config.view_name = None if selected_view == NO_VIEW_LABEL else selected_view


# --- SECTION 6 - Champs partages : titre / axes / couleur (Solution Information) ---

def _build_solution_info_fields(target, root, row_config):
    """
    Fait : construit les champs titre/axes/couleur de courbe pour un tracker Solution Information et les ajoute a root.
    Depend de : _make_field_label, _themed_textbox, CURVE_COLOR_OPTIONS, curve_color_label.
    Retourne : rien (pose sur target : txt_title/txt_x_label/txt_y_label/cmb_color).
    """
    root.Children.Add(_make_field_label("Titre du graphique (vide = nom du tracker) :"))
    target.txt_title = _themed_textbox()
    target.txt_title.Margin = Thickness(0, 4, 0, 12)
    target.txt_title.Text = row_config.chart_title or ""
    root.Children.Add(target.txt_title)

    root.Children.Add(_make_field_label("Nom de l'axe X (vide = deduit du CSV) :"))
    target.txt_x_label = _themed_textbox()
    target.txt_x_label.Margin = Thickness(0, 4, 0, 12)
    target.txt_x_label.Text = row_config.x_axis_label or ""
    root.Children.Add(target.txt_x_label)

    root.Children.Add(_make_field_label("Nom de l'axe Y (vide = deduit du CSV, courbe unique) :"))
    target.txt_y_label = _themed_textbox()
    target.txt_y_label.Margin = Thickness(0, 4, 0, 12)
    target.txt_y_label.Text = row_config.y_axis_label or ""
    root.Children.Add(target.txt_y_label)

    root.Children.Add(_make_field_label("Couleur de la courbe :"))
    target.cmb_color = ComboBox()
    target.cmb_color.Margin = Thickness(0, 4, 0, 12)
    for color_label, color_value in CURVE_COLOR_OPTIONS:
        target.cmb_color.Items.Add(color_label)
    target.cmb_color.SelectedItem = curve_color_label(row_config.curve_color)
    root.Children.Add(target.cmb_color)


def _apply_solution_info_fields(target, row_config):
    """
    Fait : lit les champs titre/axes/couleur depuis target et les applique a row_config.
    Depend de : target.txt_title/txt_x_label/txt_y_label/cmb_color, curve_color_from_label.
    Retourne : rien (effet de bord sur row_config).
    """
    row_config.chart_title = target.txt_title.Text.strip() or None
    row_config.x_axis_label = target.txt_x_label.Text.strip() or None
    row_config.y_axis_label = target.txt_y_label.Text.strip() or None
    row_config.curve_color = curve_color_from_label(unicode(target.cmb_color.SelectedItem))


# --- SECTION 6bis - "Slide combinee (differents resultats)" : etat et constantes ---
# Le flux "template puis grille puis case" vivait auparavant dans 3 boites de dialogue modales
# (MultiResultTemplatePickerWindow / MultiResultGridWindow / ResultPickerWindow). Il est desormais
# integre directement dans l'onglet "04   Slide combinee" de la fenetre principale (voir
# ReportGeneratorApp._build_multi_result_tab et les methodes _multi_result_*/_on_multi_result_*) :
# plus aucune fenetre separee, le choix du template, la grille et la configuration d'une case vivent
# tous dans ce meme onglet (template en haut, grille a gauche, panneau de case a droite).

GRID_CELL_UNCONFIGURED_BRUSH = ROW_STATUS_NOT_SELECTED_BRUSH
GRID_CELL_CONFIGURED_BRUSH = ROW_STATUS_CONFIGURED_BRUSH
GRID_CELL_DISABLED_BRUSH = SolidColorBrush(WpfColor.FromRgb(0xD8, 0xD8, 0xD8))
GRID_CELL_SELECTED_BORDER_BRUSH = SolidColorBrush(WpfColor.FromRgb(0x00, 0x8D, 0xD5))

MULTI_RESULT_CELL_TOTAL = 8  # nombre d'emplacements de la grille (2 lignes x 4 colonnes, voir gridMultiResultCells)


class MultiResultSlideConfig(object):
    """
    Une slide combinee "differents resultats" configuree depuis l'onglet "Slide combinee", en attente
    de generation : elle n'est pas construite immediatement mais stockee dans
    ReportGeneratorApp._multi_result_slides et apparait comme une carte a part entiere dans l'onglet
    Apercu, generee seulement lors du clic sur "Generer le rapport" (meme session PowerPoint que tout
    le reste, meme ordre glisser-depose).
    """

    def __init__(self, template_count, cell_configs):
        """
        Fait : stocke le nombre d'emplacements et la configuration graphique de chaque case.
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.template_count = template_count
        self.cell_configs = cell_configs  # liste de SlideRowConfig, une par case, dans l'ordre des emplacements


class _ConfigFieldsHolder(object):
    """
    Conteneur generique pour les controles d'un panneau de configuration cree en code (Border/
    ComboBox/TextBox... sans fenetre dediee) : sert de `target` a _build_row_config_fields /
    _build_steps_section_fields / _build_geometry_part_fields / _build_mesh_part_fields /
    _build_solution_info_fields (et leurs _apply_* correspondants), aussi bien pour le panneau
    lateral global ("..." de n'importe quelle ligne, voir ReportGeneratorApp._open_config_panel)
    que pour le panneau de case de l'onglet "Slide combinee" (_show_multi_result_editor).
    """
    pass


# --- SECTION 7 - Fenetre principale (chargee depuis AnsysReportGenerator_WPF.xaml) ---

class ReportGeneratorApp(object):
    """
    Point d'entree WPF de l'application : meme logique metier (fonctions de 04_slides.py /
    05_interactive_slides.py), presentation chargee depuis AnsysReportGenerator_WPF.xaml et
    organisee en 3 onglets :

    - "Slides generales" : Geometrie / Maillage (cases a cocher simples) +
      "Pieces a isoler (geometrie)", "Piece a isoler mesh" et "Contexte
      d'analyse" (une slide "Analysis Parameters" par analyse cochee - voir
      collect_analyses de 05_interactive_slides.py) (grilles de selection).
    - "Conditions et contacts" : Boundary Conditions, Bolt Pretension,
      Contacts a afficher, Connexion : Contact Tool (Contact Tool sans step,
      branche Connections), Solution Information.
    - "Categories de resultats" : Contact Tool Results (Contact Tool avec
      steps, branche Solution), Resultats, Bolt Tool - une slide par ligne
      cochee.

    Un 4eme onglet "Apercu du rapport" affiche une carte par slide selectionnee
    (nom + parametres detailles), reorganisable par glisser-deposer : l'ordre
    choisi est respecte a la generation (voir self._preview_order).
    """

    def __init__(self, xaml_path):
        """
        Fait : initialise l'application (collecte des donnees Mechanical, chargement du XAML, cablage des controles).
        Depend de : collect_* (05_interactive_slides.py), ExtAPI.DataModel, self._load_window/_find_controls/_build_sections/_wire_*.
        Retourne : rien (construit self.window pret a etre affiche par SECTION 8).
        """
        remove_stale_figures()

        self._analysis = ExtAPI.DataModel.Project.Model.Analyses[0]

        # Resultats / Contact Tool Results / Bolt Tool / Bolt Pretension / Solution Information
        # sont compiles depuis TOUTES les analyses du projet des qu'il y en a plus d'une, et
        # taguees avec leur analyse d'origine (tuples (obj, analysis)) pour differencier les noms
        # identiques venant de deux analyses differentes (voir analysis_suffix). Sur un projet
        # mono-analyse, comportement identique a avant (listes taguees (obj, None), jamais suffixees).
        self._analyses = collect_analyses()
        self._multi_analysis = len(self._analyses) > 1
        print "Analyses trouvees : {}".format(len(self._analyses))
        for _i, _a in enumerate(self._analyses, start=1):
            print "  {} : {}".format(_i, _a.Name)

        self._contact_tool_connections_results = collect_connection_contact_tool_results(self._analysis)
        self._bodies = collect_bodies()
        self._contact_regions = collect_contact_regions()

        if self._multi_analysis:
            self._bcs = collect_boundary_conditions_multi(self._analyses)
            self._results = collect_all_results_multi(self._analyses)
            self._contact_tool_results = collect_contact_tool_results_multi(self._analyses)
            self._bolt_tool_results = collect_bolt_tool_results_multi(self._analyses)
            self._bolt_pretensions = collect_bolt_pretensions_multi(self._analyses)
            self._solution_info_trackers = collect_solution_information_trackers_multi(self._analyses)
        else:
            self._bcs = [(obj, None) for obj in collect_boundary_conditions(self._analysis)]
            self._results = [(obj, None) for obj in collect_all_results(self._analysis)]
            self._contact_tool_results = [(obj, None) for obj in collect_contact_tool_results(self._analysis)]
            self._bolt_tool_results = [(obj, None) for obj in collect_bolt_tool_results(self._analysis)]
            self._bolt_pretensions = [(obj, None) for obj in collect_bolt_pretensions(self._analysis)]
            self._solution_info_trackers = [(obj, None) for obj in collect_solution_information_trackers(self._analysis)]

        self._views = collect_views()
        self._section_planes = collect_section_planes()
        self._section_plane_labels = [
            section_plane_label(sp, i) for i, sp in enumerate(self._section_planes)
        ]
        self._step_count = get_step_count(self._analysis)
        self._legend_names = collect_legend_files()

        # Vue (View Manager) choisie pour les slides Geometrie/Maillage (cases a cocher simples,
        # pas de liste) : reutilise MeshPartRowConfig (vue uniquement, deja utilise pour le mesh
        # par piece isolee) plutot que d'introduire une classe dediee, ses attributs (obj/view_name/
        # configured) suffisant tels quels.
        self._geometry_view_config = MeshPartRowConfig(ExtAPI.DataModel.Project.Model.Geometry)
        self._mesh_view_config = MeshPartRowConfig(ExtAPI.DataModel.Project.Model.Mesh)

        # Choix du tableau de maillage (par defaut / complet) pour la slide Maillage : deplace du
        # ComboBox autrefois dans la carte "Slide maillage" (onglet 01) vers le panneau "PARAMETRES"
        # de cette meme case (voir _open_config_panel, kind="mesh_part" + row_config is self._mesh_view_config).
        self._mesh_table_full = False  # False = tableau par defaut, True = tableau complet

        self._sections = {}       # nom de section -> {rows, group_key, label, search_box}
        self._section_order = []  # ordre d'affichage stable pour l'apercu
        self._multi_result_slides = []  # liste de MultiResultSlideConfig, une par slide "differents resultats" ajoutee au rapport (voir onglet "Slide combinee", _on_multi_result_add_to_report)

        # --- Etat du panneau lateral global de configuration ("...", voir _open_config_panel) ---
        self._config_panel_kind = None         # "result"/"geometry_part"/"mesh_part"/"solution_info"/None
        self._config_panel_row_config = None   # row_config en cours d'edition dans le panneau
        self._config_panel_fields = None       # _ConfigFieldsHolder courant (cmb_view/cmb_section/...)
        self._config_panel_refresh = None      # callable() invoque apres "Appliquer" (rafraichit la ligne/l'apercu)

        # --- Etat de l'onglet "Slide combinee" (grille en cours de construction, voir _build_multi_result_tab) ---
        self._mr_template_count = None    # nombre de cases actives du template choisi (2/3/4/6/8)
        self._mr_cell_configs = [None] * MULTI_RESULT_CELL_TOTAL  # index -> SlideRowConfig ou None
        self._mr_cell_borders = []        # Border de chaque case de gridMultiResultCells, dans l'ordre
        self._mr_cell_labels = []         # TextBlock de chaque case, dans l'ordre
        self._mr_template_buttons = {}    # nombre de resultats -> Button (pour l'etat visuel "selectionne")
        self._mr_selected_cell_index = None  # case actuellement affichee dans panelMultiResultSidePanel (None = aucune)
        self._mr_editing = None           # (index, cfg) en cours d'edition dans le panneau de droite (etat "editeur")
        self._mr_editor_fields = None     # _ConfigFieldsHolder courant (cmb_view/cmb_section/... de l'editeur affiche)
        self._mr_picker_rows = []         # [(Border, TextBlock, resultat)] de la liste "choisir un resultat" affichee

        self._preview_order = []  # liste ordonnee de tuples (kind, payload), reorganisable par glisser-deposer
        self._entry_to_card = {}  # (kind, payload) -> Border actuellement affiche dans panelPreview
        self._entry_to_badge = {}  # (kind, payload) -> TextBlock du badge numerote de la carte

        # Etat du glisser-deposer en cours (voir _begin_potential_drag / _start_drag / _end_drag).
        self._drag_pending_card = None
        self._drag_pending_entry = None
        self._drag_start_point = None
        self._drag_active = False
        self._drag_entry = None
        self._drag_source_card = None
        self._drag_popup = None

        self._last_report_path = None  # chemin du dernier rapport PPTX genere (onglet Fichiers)

        self.window = self._load_window(xaml_path)

        # Rend les ressources (brushes + styles) de la fenetre principale accessibles aux panneaux
        # de configuration "..." construits en code (voir _shared_resources, section 2) : une seule
        # instance de ReportGeneratorApp par execution (SECTION 8), donc cette assignation unique
        # suffit pour toute la duree du script.
        global _shared_resources
        _shared_resources = self.window.Resources

        self._find_controls()
        self._refresh_general_slide_status()
        self._build_sections()
        self._build_multi_result_tab()
        self._wire_contacts_filter()
        self._wire_zone_select_buttons()
        self._wire_file_paths()
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()
        self._wire_events()
        self._update_preview()

    # --- Chargement du XAML ---

    def _load_window(self, xaml_path):
        """
        Fait : charge AnsysReportGenerator_WPF.xaml et construit la Window WPF correspondante.
        Depend de : StreamReader, XmlReader, XamlReader.Load (System.Windows.Markup).
        Retourne : Window, la fenetre chargee (pas encore affichee).
        """
        reader = StreamReader(xaml_path)
        xml_reader = XmlReader.Create(reader)
        return XamlReader.Load(xml_reader)

    def _find_controls(self):
        """
        Fait : recupere les references des controles nommes (x:Name) definis dans le XAML.
        Depend de : self.window.FindName, les x:Name declares dans AnsysReportGenerator_WPF.xaml.
        Retourne : rien (initialise les attributs self.btn_.../chk_.../panel_... etc.).
        """
        w = self.window
        self._load_logo()

        self.btn_delete_figures = w.FindName("btnDeleteFigures")
        self.btn_reset_legends = w.FindName("btnResetLegends")
        self.btn_create_views = w.FindName("btnCreateViews")
        self.btn_export_3d = w.FindName("btnExport3D")

        self.border_config_panel = w.FindName("borderConfigPanel")
        self.panel_config_panel = w.FindName("panelConfigPanel")

        self.panel_multi_result_template_buttons = w.FindName("panelMultiResultTemplateButtons")
        self.lbl_multi_result_fill_count = w.FindName("lblMultiResultFillCount")
        self.btn_multi_result_add_to_report = w.FindName("btnMultiResultAddToReport")
        self.grid_multi_result_cells = w.FindName("gridMultiResultCells")
        self.lbl_multi_result_hint = w.FindName("lblMultiResultHint")
        self.panel_multi_result_side = w.FindName("panelMultiResultSidePanel")

        self.chk_geometry = w.FindName("chkGeometry")
        self.btn_geometry_view = w.FindName("btnGeometryView")
        self.lbl_geometry_status = w.FindName("lblGeometryStatus")
        self.chk_mesh = w.FindName("chkMesh")
        self.btn_mesh_view = w.FindName("btnMeshView")
        self.lbl_mesh_status = w.FindName("lblMeshStatus")
        self.cmb_contacts_filter = w.FindName("cmbContactsFilter")

        self.btn_check_all_general = w.FindName("btnCheckAllGeneral")
        self.btn_uncheck_all_general = w.FindName("btnUncheckAllGeneral")
        self.btn_check_all_conditions = w.FindName("btnCheckAllConditions")
        self.btn_uncheck_all_conditions = w.FindName("btnUncheckAllConditions")
        self.btn_check_all_results = w.FindName("btnCheckAllResults")
        self.btn_uncheck_all_results = w.FindName("btnUncheckAllResults")

        self.panel_preview = w.FindName("panelPreview")
        self.btn_generate = w.FindName("btnGenerate")
        self.btn_close = w.FindName("btnClose")

        self.panel_csv_files = w.FindName("panelCsvFiles")
        self.panel_data_cleanup = w.FindName("panelDataCleanup")
        self.btn_delete_all_data = w.FindName("btnDeleteAllData")
        self.btn_reset_file_paths = w.FindName("btnResetFilePaths")
        self.border_report_status = w.FindName("borderReportStatus")
        self.border_report_tile = w.FindName("borderReportTile")
        self.lbl_progress_status = w.FindName("lblProgressStatus")
        self.progress_track = w.FindName("progressTrack")
        self.progress_fill = w.FindName("progressFill")
        self.lbl_report_name = w.FindName("lblReportName")
        self.btn_report_view = w.FindName("btnReportView")
        self.btn_report_show_in_folder = w.FindName("btnReportShowInFolder")

    def _load_logo(self):
        """
        Fait : charge le logo entreprise (LOGO_PATH) dans la ressource "SidebarLogoBitmap" utilisee
        par la carte credit en bas de la colonne d'onglets (voir le ControlTemplate de TabControl
        dans le XAML, Image avec Source="{DynamicResource SidebarLogoBitmap}") - DynamicResource
        (pas x:Name/FindName) car cette Image vit dans le NameScope prive du ControlTemplate,
        inaccessible depuis self.window.FindName().
        Depend de : self.window.Resources, LOGO_PATH (00_constants.py), BitmapImage/BitmapCacheOption/Uri.
        Retourne : rien (effet de bord sur self.window.Resources["SidebarLogoBitmap"] ; ne fait rien si LOGO_PATH est absent).
        """
        if not os.path.isfile(LOGO_PATH):
            return
        try:
            bitmap = BitmapImage()
            bitmap.BeginInit()
            # CacheOption.OnLoad : charge le fichier entierement en memoire puis libere le handle
            # immediatement (sinon le PNG reste ouvert/verrouille par le processus Mechanical pour
            # toute la duree de la session).
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.UriSource = Uri(LOGO_PATH, UriKind.Absolute)
            bitmap.EndInit()
            self.window.Resources["SidebarLogoBitmap"] = bitmap
        except Exception as e:
            print "Chargement du logo impossible : " + str(e)

    def _refresh_general_slide_status(self):
        """
        Fait : rafraichit le texte de statut des cartes "Slide geometrie"/"Slide maillage" (onglet 01).
        Depend de : self.lbl_geometry_status/lbl_mesh_status, self._geometry_view_config/_mesh_view_config, _general_slide_status_text.
        Retourne : rien (effet de bord sur lblGeometryStatus/lblMeshStatus).
        """
        self.lbl_geometry_status.Text = _general_slide_status_text(self._geometry_view_config)
        self.lbl_mesh_status.Text = _general_slide_status_text(self._mesh_view_config)

    def _wire_contacts_filter(self):
        """
        Fait : peuple et cable le ComboBox de filtre par type de contact (section "Contacts a afficher").
        Depend de : self.cmb_contacts_filter, CONTACTS_FILTER_OPTIONS, self._on_contacts_filter_changed.
        Retourne : rien (effet de bord sur self.cmb_contacts_filter).
        """
        for label in CONTACTS_FILTER_OPTIONS:
            self.cmb_contacts_filter.Items.Add(label)
        self.cmb_contacts_filter.SelectedIndex = 0
        self.cmb_contacts_filter.SelectionChanged += self._on_contacts_filter_changed

    def _on_contacts_filter_changed(self, sender, e):
        """
        Fait : reordonne VISUELLEMENT panelContacts pour faire remonter en haut les contacts du type
        choisi (Frictional/Bonded/Autres), ou restaure l'ordre d'origine ("Tous"). Meme principe que
        _perform_search (voir plus bas) : seul l'ordre d'AFFICHAGE change, self._sections["Contacts"]["rows"]
        garde son ordre d'origine, qui reste celui utilise pour la generation du rapport.
        Depend de : self._sections["Contacts"]["rows"/"panel"], _classify_contact_name, CONTACTS_FILTER_OPTIONS.
        Retourne : rien (effet de bord sur l'ordre visuel de panelContacts).
        """
        selected = unicode(self.cmb_contacts_filter.SelectedItem) if self.cmb_contacts_filter.SelectedItem else "Tous"
        section = self._sections["Contacts"]
        rows = section["rows"]
        panel = section["panel"]

        if selected not in CONTACTS_FILTER_OPTIONS or selected == "Tous":
            ordered_rows = rows
        else:
            matching = [row for row in rows if _classify_contact_name(row.row_config.obj.Name) == selected]
            rest = [row for row in rows if row not in matching]
            ordered_rows = matching + rest

        panel.Children.Clear()
        for row in ordered_rows:
            panel.Children.Add(row.border)

    # --- Onglet "Fichiers" : chemins modifiables ---
    # Voir FILE_PATH_SETTINGS / _DEFAULT_FILE_PATHS (SECTION 1) : chaque ligne modifie
    # directement le global correspondant de 00_constants.py, lu par tout Report
    # Generator/*.py au moment de l'appel - aucun autre changement necessaire ailleurs.

    def _wire_file_paths(self):
        """
        Fait : initialise les 5 TextBox de chemins avec leur valeur courante et cable boutons "..."/reset.
        Depend de : FILE_PATH_SETTINGS, self._make_path_edit_handler/_make_path_browse_handler/_on_reset_file_paths.
        Retourne : rien (effet de bord sur les controles de l'onglet Fichiers).
        """
        self._path_textboxes = {}
        for name, textbox_id, browse_id, kind in FILE_PATH_SETTINGS:
            textbox = self.window.FindName(textbox_id)
            browse_button = self.window.FindName(browse_id)
            textbox.Text = globals()[name]
            self._path_textboxes[name] = textbox
            textbox.LostFocus += self._make_path_edit_handler(name, kind)
            browse_button.Click += self._make_path_browse_handler(name, kind)
        self.btn_reset_file_paths.Click += self._on_reset_file_paths
        self.btn_delete_all_data.Click += self._on_delete_all_data

    def _make_path_edit_handler(self, name, kind):
        """
        Fait : ferme name/kind par valeur pour produire le handler LostFocus d'une TextBox de chemin.
        Depend de : self._apply_file_path_edit.
        Retourne : function, le handler(sender, e) a cabler sur textbox.LostFocus.
        """
        def handler(sender, e):
            """
            Fait : valide le chemin tape des que la TextBox perd le focus.
            Depend de : self._apply_file_path_edit, name/kind captures par la fermeture.
            Retourne : rien (effet de bord sur le global correspondant).
            """
            self._apply_file_path_edit(name, kind)
        return handler

    def _make_path_browse_handler(self, name, kind):
        """
        Fait : ferme name/kind par valeur pour produire le handler Click du bouton "..." d'une ligne de chemin.
        Depend de : self._browse_file_path.
        Retourne : function, le handler(sender, e) a cabler sur browse_button.Click.
        """
        def handler(sender, e):
            """
            Fait : ouvre le dialogue de parcours pour choisir un chemin personnalise.
            Depend de : self._browse_file_path, name/kind captures par la fermeture.
            Retourne : rien (effet de bord sur le global correspondant).
            """
            self._browse_file_path(name, kind)
        return handler

    def _apply_file_path_edit(self, name, kind):
        """
        Fait : valide le texte tape dans la TextBox du chemin 'name' et reaffecte le global correspondant si valide.
        Depend de : self._path_textboxes, os.path, ensure_folder_exists, globals().
        Retourne : rien (effet de bord : met a jour globals()[name] ou restaure le texte precedent).
        """
        textbox = self._path_textboxes[name]
        new_value = textbox.Text.strip()
        current_value = globals()[name]
        if new_value == current_value:
            return

        if kind == "file":
            if not os.path.isfile(new_value) or not new_value.lower().endswith(".pptx"):
                MessageBox.Show("Chemin invalide : le template doit etre un fichier .pptx existant.",
                                 "Chemin invalide", MessageBoxButton.OK, MessageBoxImage.Warning)
                textbox.Text = current_value
                return
        else:
            try:
                ensure_folder_exists(new_value)
            except Exception as ex:
                MessageBox.Show("Impossible d'utiliser ce dossier :\n" + str(ex),
                                 "Chemin invalide", MessageBoxButton.OK, MessageBoxImage.Warning)
                textbox.Text = current_value
                return

        globals()[name] = new_value
        print "Chemin '{}' mis a jour : {}".format(name, new_value)
        if name == "CSV_EXPORT_FOLDER":
            self._refresh_csv_files()

    def _browse_file_path(self, name, kind):
        """
        Fait : ouvre un OpenFileDialog (template) ou FolderBrowserDialog (dossiers) pour choisir un chemin.
        Depend de : self._path_textboxes, System.Windows.Forms (SWF.OpenFileDialog/FolderBrowserDialog), self._apply_file_path_edit.
        Retourne : rien (effet de bord sur la TextBox et le global correspondant si un chemin est choisi).
        """
        textbox = self._path_textboxes[name]
        current_value = globals()[name]
        if kind == "file":
            dialog = SWF.OpenFileDialog()
            dialog.Filter = "PowerPoint (*.pptx)|*.pptx"
            if os.path.isfile(current_value):
                dialog.InitialDirectory = os.path.dirname(current_value)
            if dialog.ShowDialog() == SWF.DialogResult.OK:
                textbox.Text = dialog.FileName
                self._apply_file_path_edit(name, kind)
        else:
            dialog = SWF.FolderBrowserDialog()
            if os.path.isdir(current_value):
                dialog.SelectedPath = current_value
            if dialog.ShowDialog() == SWF.DialogResult.OK:
                textbox.Text = dialog.SelectedPath
                self._apply_file_path_edit(name, kind)

    def _on_reset_file_paths(self, sender, e):
        """
        Fait : bouton "Reinitialiser les chemins" - revient aux valeurs d'origine de 00_constants.py.
        Depend de : FILE_PATH_SETTINGS, _DEFAULT_FILE_PATHS, self._refresh_csv_files.
        Retourne : rien (effet de bord : reaffecte globals() et les TextBox de chemins).
        """
        for name, textbox_id, browse_id, kind in FILE_PATH_SETTINGS:
            default_value = _DEFAULT_FILE_PATHS[name]
            globals()[name] = default_value
            self._path_textboxes[name].Text = default_value
        self._refresh_csv_files()
        print "Chemins de fichiers reinitialises aux valeurs d'origine."

    # --- Onglet "Fichiers" : liste des CSV disponibles ---

    def _refresh_csv_files(self):
        """
        Fait : reconstruit panelCsvFiles (liste, pas des tuiles) a partir du contenu courant de CSV_EXPORT_FOLDER.
        Depend de : CSV_EXPORT_FOLDER, os.listdir, self._build_csv_row.
        Retourne : rien (effet de bord sur self.panel_csv_files).
        """
        self.panel_csv_files.Children.Clear()
        try:
            names = sorted(f for f in os.listdir(CSV_EXPORT_FOLDER) if f.lower().endswith(".csv"))
        except Exception:
            names = []

        if not names:
            placeholder = TextBlock()
            placeholder.Text = "(Aucun fichier CSV pour le moment)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_csv_files.Children.Add(placeholder)
            return

        for name in names:
            self.panel_csv_files.Children.Add(self._build_csv_row(name, os.path.join(CSV_EXPORT_FOLDER, name)))

    def _build_csv_row(self, name, path):
        """
        Fait : construit une ligne de la liste des CSV (nom a gauche, boutons Ouvrir/Afficher dans le dossier a droite).
        Depend de : self._make_view_handler/_make_show_in_folder_handler, _shared_resources.
        Retourne : Border, la ligne prete a etre ajoutee a panelCsvFiles (bordure basse = separateur de tableau).
        """
        grid = Grid()
        col_name = ColumnDefinition()
        col_name.Width = GridLength(1, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_name)
        col_actions = ColumnDefinition()
        col_actions.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(col_actions)

        text_block = TextBlock()
        text_block.Text = name
        text_block.VerticalAlignment = VerticalAlignment.Center
        text_block.TextTrimming = TextTrimming.CharacterEllipsis
        text_block.Margin = Thickness(0, 0, 10, 0)
        Grid.SetColumn(text_block, 0)
        grid.Children.Add(text_block)

        actions = StackPanel()
        actions.Orientation = Orientation.Horizontal
        actions.VerticalAlignment = VerticalAlignment.Center

        btn_open = _themed_button()
        btn_open.Content = "Ouvrir"
        btn_open.Padding = Thickness(8, 3, 8, 3)
        btn_open.FontSize = 11
        btn_open.Margin = Thickness(0, 0, 6, 0)
        btn_open.Click += self._make_view_handler(path)
        actions.Children.Add(btn_open)

        btn_show = _themed_button()
        btn_show.Content = "Afficher dans le dossier"
        btn_show.Padding = Thickness(8, 3, 8, 3)
        btn_show.FontSize = 11
        btn_show.Click += self._make_show_in_folder_handler(path)
        actions.Children.Add(btn_show)

        Grid.SetColumn(actions, 1)
        grid.Children.Add(actions)

        row = Border()
        row.BorderBrush = _shared_resources["CardBorderBrush"]
        row.BorderThickness = Thickness(0, 0, 0, 1)
        row.Padding = Thickness(4, 6, 4, 6)
        row.Child = grid
        return row

    def _make_view_handler(self, path):
        """
        Fait : ferme path par valeur pour produire le handler Click du bouton "Ouvrir" d'une ligne CSV.
        Depend de : self._on_view_file.
        Retourne : function, le handler(sender, e) a cabler sur btn_open.Click.
        """
        def handler(sender, e):
            """
            Fait : ouvre le fichier CSV avec son application associee.
            Depend de : self._on_view_file, path capture par la fermeture.
            Retourne : rien (effet de bord : lance l'application associee).
            """
            self._on_view_file(path)
        return handler

    def _make_show_in_folder_handler(self, path):
        """
        Fait : ferme path par valeur pour produire le handler Click du bouton "Afficher dans le dossier".
        Depend de : self._on_show_in_folder.
        Retourne : function, le handler(sender, e) a cabler sur btn_show.Click.
        """
        def handler(sender, e):
            """
            Fait : ouvre l'explorateur Windows avec le fichier surligne.
            Depend de : self._on_show_in_folder, path capture par la fermeture.
            Retourne : rien (effet de bord : lance explorer.exe).
            """
            self._on_show_in_folder(path)
        return handler

    def _on_view_file(self, path):
        """
        Fait : ouvre un fichier avec son application associee (equivalent d'un double-clic dans l'explorateur).
        Depend de : System.Diagnostics.Process.Start.
        Retourne : rien (effet de bord : lance l'application associee, affiche une MessageBox si echec).
        """
        try:
            Process.Start(path)
        except Exception as ex:
            MessageBox.Show("Impossible d'ouvrir le fichier :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_show_in_folder(self, path):
        """
        Fait : ouvre l'explorateur Windows avec le fichier deja selectionne/surligne ("Afficher dans le dossier").
        Depend de : System.Diagnostics.Process.Start("explorer.exe", "/select,...") (API .NET).
        Retourne : rien (effet de bord : lance explorer.exe, affiche une MessageBox si echec).
        """
        try:
            Process.Start("explorer.exe", '/select,"{}"'.format(path))
        except Exception as ex:
            MessageBox.Show("Impossible d'afficher le fichier dans le dossier :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    # --- Onglet "Fichiers" : nettoyage des dossiers de donnees ---
    # Une tuile par sous-dossier de DATA_ROOT (hors legendes, voir list_data_cleanup_folders dans
    # 00_constants.py) : taille + nombre de fichiers, bouton "Vider" individuel (rouge clair), et un
    # bouton global "Tout supprimer" (rouge voyant, btnDeleteAllData) qui vide tous ces dossiers d'un
    # coup. Les legendes ne sont jamais concernees : ce sont des fichiers de configuration reutilises
    # d'une generation a l'autre, pas des exports jetables.

    def _refresh_data_cleanup_tiles(self):
        """
        Fait : reconstruit panelDataCleanup a partir du contenu courant de DATA_ROOT (hors legendes).
        Depend de : list_data_cleanup_folders (00_constants.py), self._build_data_cleanup_tile.
        Retourne : rien (effet de bord sur self.panel_data_cleanup).
        """
        self.panel_data_cleanup.Children.Clear()
        folders = list_data_cleanup_folders()

        if not folders:
            placeholder = TextBlock()
            placeholder.Text = "(Aucun dossier de donnees pour le moment)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_data_cleanup.Children.Add(placeholder)
            return

        for name, path in folders:
            self.panel_data_cleanup.Children.Add(self._build_data_cleanup_tile(name, path))

    def _build_data_cleanup_tile(self, name, path):
        """
        Fait : construit une tuile de nettoyage (nom du dossier, taille, nombre de fichiers, bouton "Vider").
        Depend de : get_folder_stats/format_folder_size (00_constants.py), self._make_clear_folder_handler, _shared_resources, CARD_NORMAL_BACKGROUND.
        Retourne : Border, la tuile prete a etre ajoutee a panelDataCleanup (etirable, panelDataCleanup est un UniformGrid 2x2).
        """
        size_bytes, file_count = get_folder_stats(path)

        content = StackPanel()

        title = TextBlock()
        title.Text = name
        title.FontWeight = FontWeights.SemiBold
        title.TextTrimming = TextTrimming.CharacterEllipsis
        content.Children.Add(title)

        detail = TextBlock()
        detail.Text = "{} - {} fichier(s)".format(format_folder_size(size_bytes), file_count)
        detail.FontSize = 10
        detail.Foreground = SEARCH_PLACEHOLDER_BRUSH
        detail.Margin = Thickness(0, 2, 0, 6)
        content.Children.Add(detail)

        btn_clear = Button()
        btn_clear.Content = "Vider"
        btn_clear.Style = _shared_resources["DangerButtonLight"]
        btn_clear.Padding = Thickness(8, 3, 8, 3)
        btn_clear.FontSize = 11
        btn_clear.HorizontalAlignment = HorizontalAlignment.Left
        btn_clear.Click += self._make_clear_folder_handler(name, path)
        content.Children.Add(btn_clear)

        # Pas de Width fixe (contrairement aux autres tuiles de l'app) : panelDataCleanup est un
        # UniformGrid 2x2 (voir XAML), chaque case doit s'etirer pour occuper l'espace alloue.
        tile = Border()
        tile.Background = CARD_NORMAL_BACKGROUND
        tile.BorderBrush = _shared_resources["CardBorderBrush"]
        tile.BorderThickness = Thickness(1)
        tile.CornerRadius = CornerRadius(0)
        tile.Padding = Thickness(10)
        tile.Margin = Thickness(4)
        tile.Child = content

        return tile

    def _make_clear_folder_handler(self, name, path):
        """
        Fait : ferme name/path par valeur pour produire le handler Click du bouton "Vider" d'une tuile de nettoyage.
        Depend de : self._on_clear_folder.
        Retourne : function, le handler(sender, e) a cabler sur btn_clear.Click.
        """
        def handler(sender, e):
            self._on_clear_folder(name, path)
        return handler

    def _on_clear_folder(self, name, path):
        """
        Fait : demande confirmation puis vide un dossier de donnees (bouton "Vider" d'une tuile).
        Depend de : clear_folder_contents (00_constants.py), REPORT_OUTPUT_FOLDER, self._reset_report_status_tile, self._refresh_csv_files/_refresh_data_cleanup_tiles, MessageBox.
        Retourne : rien (effet de bord sur le systeme de fichiers et l'UI, si confirme).
        """
        answer = MessageBox.Show(
            "Supprimer tout le contenu de \"{}\" ? Cette action est irreversible.".format(name),
            "Confirmer la suppression", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if answer != MessageBoxResult.Yes:
            return

        clear_folder_contents(path)
        print "Dossier vide : " + path

        # Le dernier rapport genere n'existe peut-etre plus si c'est justement ce dossier qui vient
        # d'etre vide : la tuile "resultat de rapport" (Ouvrir/Afficher dans le dossier) doit repasser a l'etat neutre.
        if os.path.normcase(os.path.normpath(path)) == os.path.normcase(os.path.normpath(REPORT_OUTPUT_FOLDER)):
            self._reset_report_status_tile()

        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()

    def _reset_report_status_tile(self):
        """
        Fait : remet la tuile "resultat de rapport" a l'etat neutre (ex : apres suppression du dossier des rapports).
        Depend de : _shared_resources, self.border_report_tile/lbl_report_name/btn_report_view/btn_report_show_in_folder.
        Retourne : rien (effet de bord sur self._last_report_path et les controles de la tuile de rapport).
        """
        self._last_report_path = None
        self.border_report_tile.Background = _shared_resources["SecondaryBackgroundBrush"]
        self.lbl_report_name.Text = "Aucun rapport genere"
        self.btn_report_view.IsEnabled = False
        self.btn_report_show_in_folder.IsEnabled = False

    def _on_delete_all_data(self, sender, e):
        """
        Fait : demande confirmation puis vide TOUS les dossiers de donnees (hors legendes) - bouton "Tout supprimer".
        Depend de : list_data_cleanup_folders/clear_folder_contents (00_constants.py), self._reset_report_status_tile, self._refresh_csv_files/_refresh_data_cleanup_tiles, MessageBox.
        Retourne : rien (effet de bord sur le systeme de fichiers et l'UI, si confirme).
        """
        folders = list_data_cleanup_folders()
        if not folders:
            return

        answer = MessageBox.Show(
            "Supprimer tout le contenu de {} dossier(s) de donnees (images, CSV, exports 3D, rapports...) ? "
            "Les legendes ne sont pas concernees. Cette action est irreversible.".format(len(folders)),
            "Confirmer la suppression globale", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if answer != MessageBoxResult.Yes:
            return

        for _name, path in folders:
            clear_folder_contents(path)
        print "{} dossier(s) de donnees vide(s) (legendes conservees).".format(len(folders))

        self._reset_report_status_tile()
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()

    # --- Onglet "Fichiers" : progression + rapport genere ---

    def _reset_generation_ui(self, total):
        """
        Fait : remet la tuile de statut a l'etat neutre en debut de generation.
        Depend de : _shared_resources, les controles progress_fill/lbl_progress_status/border_*/btn_report_*.
        Retourne : rien (effet de bord sur les controles de la tuile de statut).
        """
        self.progress_fill.Width = 0
        self.lbl_progress_status.Text = "Generation en cours... (0/{})".format(total)
        self.border_report_status.Background = _shared_resources["CardBackgroundBrush"]
        self.border_report_tile.Background = _shared_resources["SecondaryBackgroundBrush"]
        self.lbl_report_name.Text = "Generation en cours..."
        self.btn_report_view.IsEnabled = False
        self.btn_report_show_in_folder.IsEnabled = False

    def _update_generation_progress(self, done, total):
        """
        Fait : avance la barre de progression, rafraichit la grille CSV et les tuiles de nettoyage, repompe la boucle de messages Win32.
        Depend de : self.progress_fill/progress_track/lbl_progress_status, self._refresh_csv_files/_refresh_data_cleanup_tiles, SWF.Application.DoEvents().
        Retourne : rien (effet de bord sur l'UI ; DoEvents() garde la fenetre reactive pendant la generation).
        """
        # SWF.Application.DoEvents() (meme technique que _set_result_display_time de
        # 05_interactive_slides.py) : sans cet appel, la fenetre gele jusqu'a la fin de la generation.
        fraction = float(done) / total if total else 1.0
        self.progress_fill.Width = fraction * self.progress_track.ActualWidth
        self.lbl_progress_status.Text = "Generation en cours... ({}/{})".format(done, total)
        self._refresh_csv_files()
        self._refresh_data_cleanup_tiles()
        SWF.Application.DoEvents()

    def _mark_report_ready(self, path):
        """
        Fait : active la tuile PPTX (Ouvrir/Afficher dans le dossier) une fois le rapport genere.
        Depend de : _shared_resources["ReportReadyBackgroundBrush"], self.border_report_tile/lbl_report_name/btn_report_*.
        Retourne : rien (effet de bord sur les controles de la tuile de rapport).
        """
        # Seule la sous-tuile "resultat" (borderReportTile) passe en vert - la tuile englobante et
        # la sous-tuile de progression restent neutres, pour ne mettre en avant que l'element qui
        # donne effectivement acces au rapport.
        self.lbl_progress_status.Text = "Rapport termine"
        self.progress_fill.Width = self.progress_track.ActualWidth
        self.border_report_tile.Background = _shared_resources["ReportReadyBackgroundBrush"]
        self.lbl_report_name.Text = os.path.basename(path)
        self.btn_report_view.IsEnabled = True
        self.btn_report_show_in_folder.IsEnabled = True

    def _on_view_report(self, sender, e):
        """
        Fait : ouvre le dernier rapport PPTX genere (bouton "Ouvrir" de la tuile de rapport).
        Depend de : self._last_report_path, self._on_view_file.
        Retourne : rien (effet de bord : lance PowerPoint sur le fichier si un rapport existe).
        """
        if self._last_report_path:
            self._on_view_file(self._last_report_path)

    def _on_show_report_in_folder(self, sender, e):
        """
        Fait : ouvre l'explorateur Windows sur le dernier rapport PPTX genere (bouton "Afficher dans le dossier" de la tuile de rapport).
        Depend de : self._last_report_path, self._on_show_in_folder.
        Retourne : rien (effet de bord : lance explorer.exe si un rapport existe).
        """
        if self._last_report_path:
            self._on_show_in_folder(self._last_report_path)

    # --- Construction des 9 sections de selection ---

    def _build_sections(self):
        """
        Fait : peuple les StackPanel de selection (x:Name panelXxx du XAML) avec une ligne par objet Mechanical.
        Depend de : self._bodies/_bcs/_bolt_pretensions/etc., self._build_section, self._wire_search_box.
        Retourne : rien (effet de bord : remplit self._sections/_section_order et les panneaux WPF).
        """
        # Le dernier element de chaque tuple (`tagged`) indique si `objects` est une liste simple
        # (row_config_factory(obj)) ou une liste de tuples (obj, analysis) (row_config_factory(obj,
        # analysis), categories multi-analyses - voir ReportGeneratorApp.__init__).
        # panel_kind identifie l'etat du panneau lateral global a afficher pour le "..." de cette
        # section (voir ReportGeneratorApp._open_config_panel) - None si la categorie n'a rien a
        # configurer (Contacts a afficher : juste une case a cocher, pas de vue/coupe/etc).
        section_defs = [
            ("GeometryParts", "panelGeometryParts", "searchGeometryParts", "general",
             "Pieces a isoler (geometrie)", self._bodies,
             GeometryPartRowConfig, build_geometry_row_display_name, "geometry_part", False),
            ("MeshParts", "panelMeshParts", "searchMeshParts", "general",
             "Piece a isoler mesh", self._bodies,
             MeshPartRowConfig, build_mesh_part_row_display_name, "mesh_part", False),
            ("AnalysisContext", "panelAnalysisContext", "searchAnalysisContext", "general",
             "Contexte d'analyse (steps, parametres)", self._analyses,
             AnalysisContextRowConfig, build_analysis_context_row_display_name, "mesh_part", False),
            ("BoundaryConditions", "panelBoundaryConditions", "searchBoundaryConditions", "conditions",
             "Boundary Conditions", self._bcs,
             SlideRowConfig, build_row_display_name, "result", True),
            ("BoltPretension", "panelBoltPretension", "searchBoltPretension", "conditions",
             "Bolt Pretension", self._bolt_pretensions,
             SlideRowConfig, build_row_display_name, "result", True),
            ("Contacts", "panelContacts", "searchContacts", "conditions",
             "Contacts a afficher", self._contact_regions,
             ContactRowConfig, build_contact_row_display_name, None, False),
            ("ContactToolConnections", "panelContactToolConnections", "searchContactToolConnections", "conditions",
             "Connexion : Contact Tool", self._contact_tool_connections_results,
             SlideRowConfig, build_row_display_name, "result", False),
            ("SolutionInfo", "panelSolutionInfo", "searchSolutionInfo", "conditions",
             "Solution Information", self._solution_info_trackers,
             SolutionInfoRowConfig, build_solution_info_row_display_name, "solution_info", True),
            ("ContactTool", "panelContactTool", "searchContactTool", "results",
             "Contact Tool Results", self._contact_tool_results,
             SlideRowConfig, build_row_display_name, "result", True),
            ("Results", "panelResults", "searchResults", "results",
             "Resultats", self._results,
             SlideRowConfig, build_row_display_name, "result", True),
            ("BoltTool", "panelBoltTool", "searchBoltTool", "results",
             "Bolt Tool", self._bolt_tool_results,
             SlideRowConfig, build_row_display_name, "result", True),
        ]

        for (name, panel_name, search_name, group_key, label_text, objects,
             row_config_factory, display_name_func, panel_kind, tagged) in section_defs:
            panel = self.window.FindName(panel_name)
            search_box = self.window.FindName(search_name)
            self._init_search_placeholder(search_box)
            rows = self._build_section(panel, objects, row_config_factory, display_name_func,
                                        panel_kind, tagged)

            self._sections[name] = {
                "rows": rows,
                "group_key": group_key,
                "label": label_text,
                "search_box": search_box,
                "panel": panel,
            }
            self._section_order.append(name)
            self._wire_search_box(search_box, rows, panel)
            self._attach_list_fade(panel.Parent)

    def _attach_list_fade(self, scroll):
        """
        Fait : ajoute un fondu blanc en bas de scroll (OpacityMask sur le ScrollViewer lui-meme,
        pas de Border de recouvrement separe - inutile ici puisque ces listes reposent toujours sur
        un fond CardBorder blanc uni), visible uniquement quand le contenu deborde reellement de la
        hauteur visible. Meme intention que _build_preview_list_container, mecanisme adapte puisque
        ces ScrollViewer sont directement definis dans le XAML (hauteur variable/MaxHeight="210" ou
        etiree, contrairement a la hauteur fixe des cartes d'apercu).
        Depend de : ITEM_LIST_FADE_HEIGHT, scroll.ScrollChanged/ActualHeight/ScrollableHeight.
        Retourne : rien (effet de bord : scroll.OpacityMask recalcule a chaque ScrollChanged).
        """
        # Espace vide ajoute sous la derniere ligne, de la meme hauteur que le fondu : sans lui, la
        # derniere ligne se retrouve exactement sous la zone de fondu (voire tronquee par le bas du
        # ScrollViewer) et devient illisible une fois tout en bas de la liste.
        content = scroll.Content
        if content is not None:
            content.Margin = Thickness(0, 0, 0, ITEM_LIST_FADE_HEIGHT)

        def update_fade(sender, e):
            """
            Fait : recalcule le masque d'opacite de scroll a chaque changement de defilement/contenu.
            Depend de : scroll (capture par la fermeture), ITEM_LIST_FADE_HEIGHT.
            Retourne : rien (effet de bord sur scroll.OpacityMask).
            """
            height = scroll.ActualHeight
            if height <= 0 or scroll.ScrollableHeight <= 0:
                scroll.OpacityMask = None
                return
            fade_start = 1.0 - min(ITEM_LIST_FADE_HEIGHT / height, 0.5)
            fade_mid = fade_start + 0.6 * (1.0 - fade_start)
            brush = LinearGradientBrush()
            brush.StartPoint = Point(0, 0)
            brush.EndPoint = Point(0, 1)
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(255, 255, 255, 255), 0))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(255, 255, 255, 255), fade_start))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(200, 255, 255, 255), fade_mid))
            brush.GradientStops.Add(GradientStop(WpfColor.FromArgb(0, 255, 255, 255), 1))
            scroll.OpacityMask = brush
        scroll.ScrollChanged += update_fade

    def _build_section(self, panel, objects, row_config_factory, display_name_func, panel_kind, tagged=False):
        """
        Fait : construit toutes les lignes WPF d'une section a partir des objets Mechanical fournis.
        Depend de : row_config_factory, self._build_row.
        Retourne : list de SectionRow, une par objet, dans l'ordre de `objects`.
        """
        rows = []
        for entry in objects:
            if tagged:
                obj, analysis = entry
                row_config = row_config_factory(obj, analysis)
            else:
                row_config = row_config_factory(entry)
            row = self._build_row(row_config, display_name_func, panel_kind)
            panel.Children.Add(row.border)
            rows.append(row)
        return rows

    def _build_row(self, row_config, display_name_func, panel_kind):
        """
        Fait : construit une ligne WPF (Border > Grid[CheckBox | TextBlock | Button?]) pour un row_config.
        Depend de : SectionRow, _row_status_brush, self._make_toggle_handler/_make_config_click_handler.
        Retourne : SectionRow, la ligne construite et cablee (evenements Checked/Unchecked/Click).
        """
        # Colonne du nom en largeur "Star" avec TextTrimming.CharacterEllipsis : le texte ne peut
        # jamais deborder sur le bouton "..." ni sur un controle voisin (contrairement a des
        # colonnes en pixels fixes).
        grid = Grid()

        col_check = ColumnDefinition()
        col_check.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(col_check)

        col_name = ColumnDefinition()
        col_name.Width = GridLength(1, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_name)

        checkbox = CheckBox()
        checkbox.Margin = Thickness(0)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        text_block = TextBlock()
        text_block.Text = display_name_func(row_config)
        text_block.VerticalAlignment = VerticalAlignment.Center
        text_block.TextTrimming = TextTrimming.CharacterEllipsis
        text_block.Margin = Thickness(6, 0, 6, 0)
        Grid.SetColumn(text_block, 1)
        grid.Children.Add(text_block)

        config_button = None
        if panel_kind:
            col_button = ColumnDefinition()
            col_button.Width = GridLength.Auto
            grid.ColumnDefinitions.Add(col_button)

            config_button = Button()
            # SecondaryButton est une ressource nommee (pas un style par defaut par TargetType) :
            # sans cette ligne, ce bouton garde le chrome Windows clair par defaut.
            config_button.Style = self.window.FindResource("SecondaryButton")
            config_button.Content = "..."
            config_button.Padding = Thickness(6, 2, 6, 2)
            config_button.MinWidth = 32
            Grid.SetColumn(config_button, 2)
            grid.Children.Add(config_button)

        border = Border()
        border.Padding = Thickness(4, 2, 4, 2)
        border.Margin = Thickness(0, 1, 0, 1)
        border.CornerRadius = CornerRadius(0)
        border.BorderThickness = Thickness(0)
        border.Child = grid

        row = SectionRow(border, checkbox, text_block, config_button, row_config,
                          display_name_func, panel_kind)
        row.border.Background = _row_status_brush(row)

        toggle_handler = self._make_toggle_handler(row)
        checkbox.Checked += toggle_handler
        checkbox.Unchecked += toggle_handler
        if config_button:
            config_button.Click += self._make_config_click_handler(row)

        return row

    def _make_toggle_handler(self, row):
        """
        Fait : ferme row par valeur pour produire le handler Checked/Unchecked d'une CheckBox de ligne.
        Depend de : _row_status_brush, self._update_preview.
        Retourne : function, le handler(sender, e) a cabler sur checkbox.Checked/Unchecked.
        """
        def handler(sender, e):
            """
            Fait : recalcule la couleur de statut (3 etats) de la ligne et rafraichit l'apercu.
            Depend de : _row_status_brush, self._update_preview, row capture par la fermeture.
            Retourne : rien (effet de bord sur row.border.Background et l'apercu).
            """
            row.border.Background = _row_status_brush(row)
            self._update_preview()
        return handler

    def _make_config_click_handler(self, row):
        """
        Fait : ferme row par valeur pour produire le handler Click du bouton "..." d'une ligne.
        Depend de : self._on_row_config_click.
        Retourne : function, le handler(sender, e) a cabler sur config_button.Click.
        """
        # Fermeture necessaire : la boucle appelante (_build_section) reutilise sa variable de
        # boucle, une reference directe a row capturerait toujours la derniere iteration.
        def handler(sender, e):
            """
            Fait : ouvre la boite de dialogue "..." de la ligne.
            Depend de : self._on_row_config_click, row capture par la fermeture.
            Retourne : rien (effet de bord : peut modifier row.row_config).
            """
            self._on_row_config_click(row)
        return handler

    def _on_row_config_click(self, row):
        """
        Fait : ouvre le panneau lateral global de configuration pour la ligne cliquee.
        Depend de : row.panel_kind/row_config/display_name_func, self._open_config_panel.
        Retourne : rien (effet de bord : affiche borderConfigPanel).
        """
        def refresh():
            """
            Fait : rafraichit le texte/la couleur de statut de la ligne et l'apercu apres "Appliquer".
            Depend de : row (capture par la fermeture), _row_status_brush, self._update_preview.
            Retourne : rien (effet de bord sur row.text_block/row.border et l'apercu).
            """
            row.text_block.Text = row.display_name_func(row.row_config)
            row.border.Background = _row_status_brush(row)
            self._update_preview()
        self._open_config_panel(row.panel_kind, row.row_config, refresh)

    # --- Panneau lateral global de configuration ("...") ---
    # Remplace les 4 boites de dialogue modales d'origine (RowConfigWindow, GeometryPartConfigWindow,
    # MeshPartConfigWindow, SolutionInfoConfigWindow - voir SECTION 4/5/5bis/6) : un seul panneau,
    # cache par defaut (borderConfigPanel.Visibility = Collapsed), qui affiche l'un de 4 "kinds" de
    # champs selon la ligne cliquee ("result"/"geometry_part"/"mesh_part"/"solution_info").
    # "Appliquer" valide et ferme (comme l'ancien bouton OK) ; "Annuler"/le bouton "x" ferment sans
    # valider (comme l'ancien bouton Annuler/la fermeture de la fenetre).

    def _open_config_panel(self, kind, row_config, refresh_callback):
        """
        Fait : ouvre le panneau lateral global pour row_config, dans l'etat kind, et l'affiche.
        Depend de : self.panel_config_panel/border_config_panel, _build_row_config_fields/_build_steps_section_fields/
            _build_geometry_part_fields/_build_mesh_part_fields/_build_solution_info_fields, _ConfigFieldsHolder.
        Retourne : rien (effet de bord : peuple panelConfigPanel, rend borderConfigPanel visible).
        """
        self._config_panel_kind = kind
        self._config_panel_row_config = row_config
        self._config_panel_refresh = refresh_callback
        self._config_panel_fields = _ConfigFieldsHolder()

        panel = self.panel_config_panel
        panel.Children.Clear()

        lbl_kicker = TextBlock()
        lbl_kicker.Text = "PARAMETRES"
        lbl_kicker.FontSize = 11
        lbl_kicker.FontWeight = FontWeights.Bold
        lbl_kicker.Foreground = _shared_resources["TextMutedBrush"]
        lbl_kicker.Margin = Thickness(0, 0, 0, 2)
        panel.Children.Add(lbl_kicker)

        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        lbl_title = TextBlock()
        lbl_title.Text = row_config.obj.Name
        lbl_title.Style = _shared_resources["CardTitle"]
        lbl_title.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(lbl_title, 0)
        header.Children.Add(lbl_title)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_config_panel_close
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        panel.Children.Add(header)

        badge = Border()
        badge.Background = ROW_STATUS_CONFIGURED_BRUSH if row_config.configured else ROW_STATUS_SELECTED_BRUSH
        badge.CornerRadius = CornerRadius(0)
        badge.Padding = Thickness(6, 2, 6, 2)
        badge.Margin = Thickness(0, 4, 0, 10)
        badge.HorizontalAlignment = HorizontalAlignment.Left
        badge_text = TextBlock()
        badge_text.Text = "configure" if row_config.configured else "a configurer"
        badge_text.FontSize = 10
        badge_text.FontWeight = FontWeights.SemiBold
        badge.Child = badge_text
        panel.Children.Add(badge)

        fields = self._config_panel_fields
        if kind == "result":
            step_count = (get_step_count(row_config.analysis) if row_config.analysis is not None
                          else self._step_count)
            _build_row_config_fields(fields, panel, row_config, self._views, self._section_plane_labels,
                                      self._legend_names)
            _build_steps_section_fields(fields, panel, row_config, step_count)
        elif kind == "geometry_part":
            _build_geometry_part_fields(fields, panel, row_config, self._views, self._section_plane_labels)
        elif kind == "mesh_part":
            _build_mesh_part_fields(fields, panel, row_config, self._views)
            # Le choix du tableau de maillage n'a de sens que pour la ligne Maillage elle-meme
            # (les autres lignes "mesh_part" - pieces isolees, contexte d'analyse - n'en ont pas).
            if row_config is self._mesh_view_config:
                panel.Children.Add(_make_field_label("Tableau de maillage :"))
                fields.cmb_mesh_table = ComboBox()
                fields.cmb_mesh_table.Margin = Thickness(0, 4, 0, 12)
                fields.cmb_mesh_table.Items.Add("Tableau par defaut (ElementSize, Nodes, Elements)")
                fields.cmb_mesh_table.Items.Add("Tableau complet (toutes les proprietes)")
                fields.cmb_mesh_table.SelectedIndex = 1 if self._mesh_table_full else 0
                panel.Children.Add(fields.cmb_mesh_table)
        elif kind == "solution_info":
            _build_solution_info_fields(fields, panel, row_config)

        buttons = StackPanel()
        buttons.Orientation = Orientation.Horizontal
        buttons.Margin = Thickness(0, 14, 0, 0)

        btn_apply = _themed_button(primary=True)
        btn_apply.Content = "Appliquer"
        btn_apply.Width = 110
        btn_apply.Margin = Thickness(0, 0, 10, 0)
        btn_apply.Click += self._on_config_panel_apply
        buttons.Children.Add(btn_apply)

        btn_cancel = _themed_button()
        btn_cancel.Content = "Annuler"
        btn_cancel.Width = 100
        btn_cancel.Click += self._on_config_panel_close
        buttons.Children.Add(btn_cancel)

        panel.Children.Add(buttons)

        self.border_config_panel.Visibility = Visibility.Visible

    def _on_config_panel_apply(self, sender, e):
        """
        Fait : valide la configuration en cours dans le panneau lateral (bouton "Appliquer") et le ferme.
        Depend de : self._config_panel_kind/_row_config/_fields/_refresh, _apply_row_config_fields/_apply_steps_section_fields/
            _apply_geometry_part_fields/_apply_mesh_part_fields/_apply_solution_info_fields.
        Retourne : rien (effet de bord : met a jour row_config.configured et rafraichit l'appelant, ferme le panneau).
        """
        kind = self._config_panel_kind
        row_config = self._config_panel_row_config
        fields = self._config_panel_fields
        refresh_callback = self._config_panel_refresh

        if kind == "result":
            _apply_row_config_fields(fields, row_config)
            _apply_steps_section_fields(fields, row_config)
        elif kind == "geometry_part":
            _apply_geometry_part_fields(fields, row_config)
        elif kind == "mesh_part":
            _apply_mesh_part_fields(fields, row_config)
            if row_config is self._mesh_view_config and hasattr(fields, "cmb_mesh_table"):
                self._mesh_table_full = (fields.cmb_mesh_table.SelectedIndex == 1)
        elif kind == "solution_info":
            _apply_solution_info_fields(fields, row_config)

        row_config.configured = True
        self._close_config_panel()
        if refresh_callback:
            refresh_callback()

    def _on_config_panel_close(self, sender, e):
        """
        Fait : ferme le panneau lateral sans valider (bouton "Annuler" ou "x").
        Depend de : self._close_config_panel.
        Retourne : rien (effet de bord : cache borderConfigPanel).
        """
        self._close_config_panel()

    def _close_config_panel(self):
        """
        Fait : vide et cache le panneau lateral global de configuration.
        Depend de : self.panel_config_panel/border_config_panel.
        Retourne : rien (effet de bord : reinitialise l'etat _config_panel_*).
        """
        self._config_panel_kind = None
        self._config_panel_row_config = None
        self._config_panel_fields = None
        self._config_panel_refresh = None
        self.panel_config_panel.Children.Clear()
        self.border_config_panel.Visibility = Visibility.Collapsed

    # --- Onglet "Slide combinee (differents resultats)" ---
    # Plus aucune fenetre separee (voir SECTION 6bis) : le choix du template peuple
    # panelMultiResultTemplateButtons, la grille vit dans gridMultiResultCells (8 cases, seules les
    # N premieres du template choisi sont actives), et panelMultiResultSidePanel affiche l'un de 3
    # etats pour la case selectionnee - aucune case (_show_multi_result_placeholder), choix du
    # resultat (_show_multi_result_picker) ou configuration graphique complete (_show_multi_result_editor,
    # memes champs qu'une slide normale via _build_row_config_fields, sans notion de step).

    def _build_multi_result_tab(self):
        """
        Fait : peuple les boutons de template et initialise la grille/le panneau lateral au chargement de la fenetre.
        Depend de : self.panel_multi_result_template_buttons, MULTI_STEP_SLIDE_TEMPLATES, self._set_multi_result_template.
        Retourne : rien (effet de bord sur l'onglet "Slide combinee").
        """
        self.panel_multi_result_template_buttons.Children.Clear()
        self._mr_template_buttons = {}
        template_counts = sorted(MULTI_STEP_SLIDE_TEMPLATES.keys())
        for count in template_counts:
            btn = _themed_button()
            btn.Content = "{} resultats".format(count)
            btn.Padding = Thickness(10, 5, 10, 5)
            btn.FontSize = 11
            btn.Margin = Thickness(0, 0, 6, 4)
            btn.Tag = count
            btn.Click += self._on_pick_multi_result_template
            self.panel_multi_result_template_buttons.Children.Add(btn)
            self._mr_template_buttons[count] = btn

        if template_counts:
            self._set_multi_result_template(template_counts[0])
        else:
            self._show_multi_result_placeholder()
            self._update_multi_result_fill_count()

    def _on_pick_multi_result_template(self, sender, e):
        """
        Fait : reagit au clic sur un bouton de template (2/3/4/6/8 resultats).
        Depend de : sender.Tag, self._set_multi_result_template.
        Retourne : rien (effet de bord : change de template, reinitialise la grille en cours).
        """
        self._set_multi_result_template(sender.Tag)

    def _set_multi_result_template(self, count):
        """
        Fait : selectionne un template (nombre de cases actives) et reinitialise entierement la grille en cours de construction.
        Depend de : self._mr_cell_configs/_refresh_multi_result_template_buttons/_rebuild_multi_result_grid/_show_multi_result_placeholder/_update_multi_result_fill_count.
        Retourne : rien (effet de bord sur l'etat de l'onglet "Slide combinee").
        """
        self._mr_template_count = count
        self._mr_cell_configs = [None] * MULTI_RESULT_CELL_TOTAL
        self._refresh_multi_result_template_buttons()
        self._rebuild_multi_result_grid()
        self._show_multi_result_placeholder()
        self._update_multi_result_fill_count()

    def _refresh_multi_result_template_buttons(self):
        """
        Fait : met en evidence (PrimaryButton) le bouton du template actuellement selectionne, les autres restant SecondaryButton.
        Depend de : self._mr_template_buttons/_mr_template_count, _shared_resources.
        Retourne : rien (effet de bord sur le Style des boutons de template).
        """
        for count, btn in self._mr_template_buttons.items():
            btn.Style = _shared_resources["PrimaryButton" if count == self._mr_template_count else "SecondaryButton"]

    def _rebuild_multi_result_grid(self):
        """
        Fait : reconstruit les 8 cases de la grille (seules les N premieres du template choisi sont actives et cliquables).
        Depend de : self.grid_multi_result_cells, self._mr_template_count, self._make_multi_result_cell_click_handler.
        Retourne : rien (effet de bord : repeuple self._mr_cell_borders/_mr_cell_labels et gridMultiResultCells).
        """
        self.grid_multi_result_cells.Children.Clear()
        self._mr_cell_borders = []
        self._mr_cell_labels = []

        for index in range(MULTI_RESULT_CELL_TOTAL):
            active = index < (self._mr_template_count or 0)

            cell = Border()
            cell.BorderBrush = _shared_resources["CardBorderBrush"]
            cell.BorderThickness = Thickness(1)
            cell.CornerRadius = CornerRadius(0)
            cell.Margin = Thickness(3)
            cell.Background = GRID_CELL_UNCONFIGURED_BRUSH if active else GRID_CELL_DISABLED_BRUSH

            label = TextBlock()
            label.TextWrapping = TextWrapping.Wrap
            label.TextTrimming = TextTrimming.CharacterEllipsis
            label.TextAlignment = TextAlignment.Center
            label.HorizontalAlignment = HorizontalAlignment.Center
            label.VerticalAlignment = VerticalAlignment.Center
            label.Foreground = _shared_resources["TextPrimaryBrush"]
            label.FontSize = 11
            label.Margin = Thickness(4)
            cell.Child = label

            if active:
                cell.Cursor = Cursors.Hand
                cell.MouseLeftButtonUp += self._make_multi_result_cell_click_handler(index)

            self.grid_multi_result_cells.Children.Add(cell)
            self._mr_cell_borders.append(cell)
            self._mr_cell_labels.append(label)
            if active:
                self._update_multi_result_cell_visual(index)

        if self._mr_template_count:
            self.lbl_multi_result_hint.Text = (
                "La grille suit le template choisi : seules les {} premieres cases sont "
                "configurables. Cliquez sur une case pour la configurer dans le panneau de "
                "droite - plus aucune fenetre separee.".format(self._mr_template_count))
        else:
            self.lbl_multi_result_hint.Text = "Choisissez un template ci-dessus pour commencer."

    def _make_multi_result_cell_click_handler(self, index):
        """
        Fait : ferme index par valeur pour produire le handler de clic d'une case active de la grille.
        Depend de : self._on_multi_result_cell_click.
        Retourne : function, le handler(sender, e) a cabler sur cell.MouseLeftButtonUp.
        """
        def handler(sender, e):
            self._on_multi_result_cell_click(index)
        return handler

    def _on_multi_result_cell_click(self, index):
        """
        Fait : selectionne la case cliquee et affiche le panneau approprie a droite (choix du resultat si vide, edition directe si deja configuree).
        Depend de : self._mr_cell_configs/_mr_selected_cell_index, self._show_multi_result_picker/_show_multi_result_editor, self._update_multi_result_cell_visual.
        Retourne : rien (effet de bord sur l'etat de selection et le panneau lateral).
        """
        self._mr_selected_cell_index = index
        for i in range(len(self._mr_cell_borders)):
            if i < (self._mr_template_count or 0):
                self._update_multi_result_cell_visual(i)

        cfg = self._mr_cell_configs[index]
        if cfg is not None:
            self._show_multi_result_editor(index, cfg)
        else:
            self._show_multi_result_picker(index)

    def _update_multi_result_cell_visual(self, index):
        """
        Fait : rafraichit le fond/texte/bordure d'une case active selon son etat (configuree, et/ou actuellement selectionnee).
        Depend de : self._mr_cell_configs/_mr_selected_cell_index/_mr_cell_borders/_mr_cell_labels, GRID_CELL_*_BRUSH.
        Retourne : rien (effet de bord sur les controles WPF de la case).
        """
        cfg = self._mr_cell_configs[index]
        border = self._mr_cell_borders[index]
        label = self._mr_cell_labels[index]

        if cfg is not None:
            border.Background = GRID_CELL_CONFIGURED_BRUSH
            label.Text = "Case {}\n{}\n(etat actuel)".format(index + 1, cfg.obj.Name)
        else:
            border.Background = GRID_CELL_UNCONFIGURED_BRUSH
            label.Text = "Case {}\ncliquer pour choisir un resultat\n+".format(index + 1)

        if index == self._mr_selected_cell_index:
            border.BorderThickness = Thickness(2)
            border.BorderBrush = GRID_CELL_SELECTED_BORDER_BRUSH
        else:
            border.BorderThickness = Thickness(1)
            border.BorderBrush = _shared_resources["CardBorderBrush"]

    def _show_multi_result_placeholder(self):
        """
        Fait : affiche le panneau lateral par defaut (aucune case selectionnee) et deselectionne visuellement la grille.
        Depend de : self.panel_multi_result_side, self._mr_selected_cell_index, self._update_multi_result_cell_visual.
        Retourne : rien (effet de bord sur panelMultiResultSidePanel et la grille).
        """
        self._mr_editing = None
        previous_index = self._mr_selected_cell_index
        self._mr_selected_cell_index = None
        if previous_index is not None and previous_index < len(self._mr_cell_borders):
            self._update_multi_result_cell_visual(previous_index)

        self.panel_multi_result_side.Children.Clear()
        txt = TextBlock()
        txt.Text = ("Choisissez un template puis cliquez sur une case de la grille pour la "
                     "configurer." if self._mr_template_count else
                     "Choisissez un template de slide combinee pour commencer.")
        txt.TextWrapping = TextWrapping.Wrap
        txt.Foreground = _shared_resources["TextMutedBrush"]
        txt.Margin = Thickness(4)
        self.panel_multi_result_side.Children.Add(txt)

    def _show_multi_result_picker(self, index, current_result=None):
        """
        Fait : affiche dans le panneau lateral la liste (filtrable) des resultats disponibles pour la case index.
        Depend de : self.panel_multi_result_side, self._results, self._init_search_placeholder, self._make_multi_result_pick_handler.
        Retourne : rien (effet de bord sur panelMultiResultSidePanel ; repeuple self._mr_picker_rows).
        """
        self._mr_editing = None
        self.panel_multi_result_side.Children.Clear()

        # Meme en-tete (kicker + titre + "x" en haut a droite) que _show_multi_result_editor et
        # _open_config_panel : les 3 etats du panneau lateral doivent se fermer de la meme facon,
        # pas de bouton "Fermer" texte en bas uniquement pour cet etat.
        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        title_panel = StackPanel()
        lbl_case = TextBlock()
        lbl_case.Text = "CASE {}".format(index + 1)
        lbl_case.FontWeight = FontWeights.Bold
        lbl_case.FontSize = 11
        lbl_case.Foreground = _shared_resources["TextMutedBrush"]
        title_panel.Children.Add(lbl_case)
        lbl_title = TextBlock()
        lbl_title.Text = "Choisir un resultat"
        lbl_title.Style = _shared_resources["CardTitle"]
        title_panel.Children.Add(lbl_title)
        Grid.SetColumn(title_panel, 0)
        header.Children.Add(title_panel)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_multi_result_close_editor
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        self.panel_multi_result_side.Children.Add(header)

        search_box = TextBox()
        search_box.Style = _shared_resources["SearchBox"]
        self._init_search_placeholder(search_box)
        search_box.TextChanged += self._on_multi_result_search_changed
        self.panel_multi_result_side.Children.Add(search_box)

        list_panel = StackPanel()
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Height = 380
        scroll.Content = list_panel
        self.panel_multi_result_side.Children.Add(scroll)

        self._mr_picker_rows = []
        result_objects = [obj for obj, _analysis in self._results]
        for result in result_objects:
            row_border = Border()
            row_border.BorderThickness = Thickness(0, 0, 0, 1)
            row_border.BorderBrush = _shared_resources["CardBorderBrush"]
            row_border.Padding = Thickness(6)
            row_border.Cursor = Cursors.Hand
            row_border.Background = GRID_CELL_CONFIGURED_BRUSH if result == current_result else Brushes.Transparent

            row_text = TextBlock()
            row_text.Text = result.Name
            row_text.FontSize = 11
            row_text.TextWrapping = TextWrapping.Wrap
            row_border.Child = row_text

            row_border.MouseLeftButtonUp += self._make_multi_result_pick_handler(index, result)
            list_panel.Children.Add(row_border)
            self._mr_picker_rows.append((row_border, row_text, result))

    def _on_multi_result_search_changed(self, sender, e):
        """
        Fait : filtre en direct la liste "Choisir un resultat" selon le texte tape (sous-chaine, insensible a la casse).
        Depend de : sender (le TextBox de recherche), self._mr_picker_rows, SEARCH_PLACEHOLDER.
        Retourne : rien (effet de bord : Visibility des lignes de self._mr_picker_rows).
        """
        text = sender.Text
        if text == SEARCH_PLACEHOLDER:
            text = ""
        query = text.strip().lower()
        for row_border, row_text, _result in self._mr_picker_rows:
            visible = (not query) or (query in row_text.Text.lower())
            row_border.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def _make_multi_result_pick_handler(self, index, result):
        """
        Fait : ferme index/result par valeur pour produire le handler de clic d'une ligne de la liste "Choisir un resultat".
        Depend de : self._on_multi_result_pick.
        Retourne : function, le handler(sender, e) a cabler sur row_border.MouseLeftButtonUp.
        """
        def handler(sender, e):
            self._on_multi_result_pick(index, result)
        return handler

    def _on_multi_result_pick(self, index, result):
        """
        Fait : reagit au choix d'un resultat pour la case index et bascule le panneau lateral en mode edition.
        Depend de : self._mr_cell_configs, SlideRowConfig, self._show_multi_result_editor.
        Retourne : rien (effet de bord sur panelMultiResultSidePanel).
        """
        existing_cfg = self._mr_cell_configs[index]
        # Reutilise la config existante (garde vue/coupe/legende/etc deja choisis) si le meme
        # resultat est reselectionne ; repart d'une config vierge si l'utilisateur change de resultat.
        if existing_cfg is not None and existing_cfg.obj == result:
            cfg = existing_cfg
        else:
            cfg = SlideRowConfig(result)
        self._show_multi_result_editor(index, cfg)

    def _show_multi_result_editor(self, index, cfg):
        """
        Fait : affiche dans le panneau lateral la configuration graphique complete (vue/coupe/legende/apparence/scale factor, sans steps) de la case index pour le resultat cfg.obj.
        Depend de : self.panel_multi_result_side, _build_row_config_fields, _ConfigFieldsHolder, ROW_STATUS_*_BRUSH.
        Retourne : rien (effet de bord sur panelMultiResultSidePanel ; initialise self._mr_editing/_mr_editor_fields).
        """
        self._mr_editing = (index, cfg)
        self.panel_multi_result_side.Children.Clear()

        header = Grid()
        col_title = ColumnDefinition()
        col_title.Width = GridLength(1, GridUnitType.Star)
        header.ColumnDefinitions.Add(col_title)
        col_close = ColumnDefinition()
        col_close.Width = GridLength.Auto
        header.ColumnDefinitions.Add(col_close)

        title_panel = StackPanel()
        lbl_case = TextBlock()
        lbl_case.Text = "CASE {}".format(index + 1)
        lbl_case.FontWeight = FontWeights.Bold
        lbl_case.FontSize = 11
        lbl_case.Foreground = _shared_resources["TextMutedBrush"]
        title_panel.Children.Add(lbl_case)
        lbl_result = TextBlock()
        lbl_result.Text = cfg.obj.Name
        lbl_result.Style = _shared_resources["CardTitle"]
        lbl_result.TextWrapping = TextWrapping.Wrap
        title_panel.Children.Add(lbl_result)
        Grid.SetColumn(title_panel, 0)
        header.Children.Add(title_panel)

        btn_close = Button()
        btn_close.Content = _build_close_icon()
        btn_close.Width = 24
        btn_close.Height = 24
        btn_close.Padding = Thickness(0)
        btn_close.Style = _shared_resources["SecondaryButton"]
        btn_close.VerticalAlignment = VerticalAlignment.Top
        btn_close.Click += self._on_multi_result_close_editor
        Grid.SetColumn(btn_close, 1)
        header.Children.Add(btn_close)

        self.panel_multi_result_side.Children.Add(header)

        status_row = StackPanel()
        status_row.Orientation = Orientation.Horizontal
        status_row.Margin = Thickness(0, 4, 0, 10)

        badge = Border()
        badge.Background = ROW_STATUS_CONFIGURED_BRUSH if cfg.configured else ROW_STATUS_SELECTED_BRUSH
        badge.CornerRadius = CornerRadius(0)
        badge.Padding = Thickness(6, 2, 6, 2)
        badge_text = TextBlock()
        badge_text.Text = "configure" if cfg.configured else "a configurer"
        badge_text.FontSize = 10
        badge_text.FontWeight = FontWeights.SemiBold
        badge.Child = badge_text
        status_row.Children.Add(badge)

        btn_change = _themed_button()
        btn_change.Content = "Changer de resultat"
        btn_change.FontSize = 11
        btn_change.Padding = Thickness(8, 2, 8, 2)
        btn_change.Margin = Thickness(8, 0, 0, 0)
        btn_change.Click += self._make_multi_result_change_handler(index, cfg)
        status_row.Children.Add(btn_change)

        self.panel_multi_result_side.Children.Add(status_row)

        self._mr_editor_fields = _ConfigFieldsHolder()
        _build_row_config_fields(self._mr_editor_fields, self.panel_multi_result_side, cfg,
                                  self._views, self._section_plane_labels, self._legend_names)

        btn_apply = _themed_button(primary=True)
        btn_apply.Content = "Appliquer"
        btn_apply.Margin = Thickness(0, 10, 0, 0)
        btn_apply.Click += self._on_multi_result_apply
        self.panel_multi_result_side.Children.Add(btn_apply)

    def _make_multi_result_change_handler(self, index, cfg):
        """
        Fait : ferme index/cfg par valeur pour produire le handler du bouton "Changer de resultat" de l'editeur de case.
        Depend de : self._show_multi_result_picker.
        Retourne : function, le handler(sender, e) a cabler sur btn_change.Click.
        """
        def handler(sender, e):
            self._show_multi_result_picker(index, cfg.obj)
        return handler

    def _on_multi_result_close_editor(self, sender, e):
        """
        Fait : ferme le panneau de case (picker ou editeur) sans valider, retour a l'etat "aucune case selectionnee".
        Depend de : self._show_multi_result_placeholder.
        Retourne : rien (effet de bord sur panelMultiResultSidePanel).
        """
        self._show_multi_result_placeholder()

    def _on_multi_result_apply(self, sender, e):
        """
        Fait : valide la configuration graphique de la case en cours d'edition (bouton "Appliquer").
        Depend de : self._mr_editing/_mr_editor_fields, _apply_row_config_fields, self._update_multi_result_cell_visual/_update_multi_result_fill_count.
        Retourne : rien (effet de bord : met a jour self._mr_cell_configs et rafraichit la case/le panneau).
        """
        if self._mr_editing is None:
            return
        index, cfg = self._mr_editing
        _apply_row_config_fields(self._mr_editor_fields, cfg)
        cfg.configured = True
        self._mr_cell_configs[index] = cfg
        self._update_multi_result_cell_visual(index)
        self._update_multi_result_fill_count()
        self._show_multi_result_editor(index, cfg)

    def _update_multi_result_fill_count(self):
        """
        Fait : rafraichit le compteur "X / N cases remplies" au-dessus de la grille.
        Depend de : self.lbl_multi_result_fill_count, self._mr_template_count/_mr_cell_configs.
        Retourne : rien (effet de bord sur lblMultiResultFillCount).
        """
        if not self._mr_template_count:
            self.lbl_multi_result_fill_count.Text = ""
            return
        filled = sum(1 for cfg in self._mr_cell_configs[:self._mr_template_count] if cfg is not None)
        self.lbl_multi_result_fill_count.Text = "{} / {} cases remplies".format(filled, self._mr_template_count)

    def _on_multi_result_add_to_report(self, sender, e):
        """
        Fait : valide que toutes les cases actives sont configurees, ajoute la slide combinee a l'apercu du rapport (generation differee) et reinitialise la grille pour en construire une autre.
        Depend de : self._mr_template_count/_mr_cell_configs, MultiResultSlideConfig, self._multi_result_slides, self._update_preview, self._set_multi_result_template.
        Retourne : rien (effet de bord : peut ajouter une entree a self._multi_result_slides et rafraichir l'apercu).
        """
        if not self._mr_template_count:
            MessageBox.Show("Choisissez d'abord un template (nombre de resultats a combiner).",
                             "Aucun template", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        active_configs = self._mr_cell_configs[:self._mr_template_count]
        missing = active_configs.count(None)
        if missing:
            MessageBox.Show(
                "Toutes les cases doivent etre configurees avant d'ajouter la slide au rapport "
                "({} case(s) sur {} manquante(s)).".format(missing, self._mr_template_count),
                "Configuration incomplete", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        self._multi_result_slides.append(MultiResultSlideConfig(self._mr_template_count, active_configs))
        self._update_preview()
        print "Slide combinee ajoutee a l'apercu du rapport ({} resultats).".format(self._mr_template_count)

        # Repart d'une grille vierge sur le meme template, pour en construire une autre a la suite.
        self._set_multi_result_template(self._mr_template_count)

    # --- Champs de recherche : texte indicatif grise ---

    def _init_search_placeholder(self, search_box):
        """
        Fait : initialise un champ de recherche avec son texte indicatif ("Rechercher...", grise).
        Depend de : SEARCH_PLACEHOLDER, SEARCH_PLACEHOLDER_BRUSH, SEARCH_TEXT_BRUSH.
        Retourne : rien (effet de bord : configure search_box et cable GotFocus/LostFocus).
        """
        search_box.Text = SEARCH_PLACEHOLDER
        search_box.Foreground = SEARCH_PLACEHOLDER_BRUSH

        def on_got_focus(sender, e):
            """
            Fait : efface le texte indicatif quand l'utilisateur clique dans le champ.
            Depend de : SEARCH_PLACEHOLDER, SEARCH_TEXT_BRUSH, search_box capture par la fermeture.
            Retourne : rien (effet de bord sur search_box).
            """
            if search_box.Text == SEARCH_PLACEHOLDER:
                search_box.Text = ""
                search_box.Foreground = SEARCH_TEXT_BRUSH

        def on_lost_focus(sender, e):
            """
            Fait : restaure le texte indicatif si le champ est laisse vide.
            Depend de : SEARCH_PLACEHOLDER, SEARCH_PLACEHOLDER_BRUSH, search_box capture par la fermeture.
            Retourne : rien (effet de bord sur search_box).
            """
            if not search_box.Text.strip():
                search_box.Text = SEARCH_PLACEHOLDER
                search_box.Foreground = SEARCH_PLACEHOLDER_BRUSH

        search_box.GotFocus += on_got_focus
        search_box.LostFocus += on_lost_focus

    # --- Recherche dans une section ---

    def _wire_search_box(self, search_box, rows, panel):
        """
        Fait : cable la touche Entree d'un champ de recherche pour declencher la recherche.
        Depend de : self._perform_search.
        Retourne : rien (effet de bord : cable search_box.KeyDown).
        """
        def on_key_down(sender, e):
            """
            Fait : declenche la recherche quand l'utilisateur appuie sur Entree.
            Depend de : self._perform_search, search_box/rows/panel captures par la fermeture.
            Retourne : rien (effet de bord : marque e.Handled et lance la recherche).
            """
            if e.Key == Key.Enter:
                e.Handled = True
                self._perform_search(search_box, rows, panel)
        search_box.KeyDown += on_key_down

    def _perform_search(self, search_box, rows, panel):
        """
        Fait : cherche et selectionne la prochaine ligne de rows dont le nom contient le texte tape.
        Depend de : search_box.Text/Tag, rows, panel, SEARCH_HIGHLIGHT_BRUSH, SEARCH_BOX_*_BACKGROUND.
        Retourne : rien (effet de bord : coche/surligne la ligne trouvee ou colore le champ en rose si aucune).
        """
        # La ligne trouvee est remontee tout en haut de panel (ordre VISUEL uniquement - "rows"
        # garde son ordre d'origine, qui reste l'ordre de generation du rapport). Une recherche
        # relancee avec le meme texte reprend apres la derniere occurrence trouvee (search_box.Tag
        # stocke (texte, index), base sur l'ordre d'origine de rows, inchange par le deplacement).
        text = search_box.Text
        if text == SEARCH_PLACEHOLDER:
            text = ""
        query = text.strip().lower()
        row_count = len(rows)
        if not query or row_count == 0:
            search_box.Background = SEARCH_BOX_DEFAULT_BACKGROUND
            return

        last_query, last_index = search_box.Tag if search_box.Tag else (None, -1)
        start_index = (last_index + 1) % row_count if last_query == query else 0

        for offset in range(row_count):
            index = (start_index + offset) % row_count
            row = rows[index]
            if query in row.text_block.Text.lower():
                for r in rows:
                    r.border.BorderThickness = Thickness(0)
                row.border.BorderThickness = Thickness(2)
                row.border.BorderBrush = SEARCH_HIGHLIGHT_BRUSH
                row.checkbox.IsChecked = True
                panel.Children.Remove(row.border)
                panel.Children.Insert(0, row.border)
                row.border.BringIntoView()
                search_box.Tag = (query, index)
                search_box.Background = SEARCH_BOX_DEFAULT_BACKGROUND
                return

        search_box.Tag = (query, -1)
        search_box.Background = SEARCH_BOX_NO_MATCH_BACKGROUND

    # --- Apercu live (onglet "Apercu du rapport") ---
    # Une carte par THEME (pas par ligne) : chaque tuple (kind, payload) de self._preview_order
    # est soit ("general", libelle) pour Geometrie/Maillage, soit (nom_de_section, None) pour une
    # section des qu'AU MOINS UNE de ses lignes est cochee - cette carte unique regroupe alors
    # tous les elements coches de la section (voir _build_preview_card). L'ordre de cette liste,
    # modifiable par glisser-deposer, est celui respecte a la generation du rapport (_on_generate).

    # Sections dont la carte garde le detail complet par element (vue, coupe, steps, ...) ; les
    # autres sections n'affichent que le nom brut de chaque element selectionne.
    FULL_DETAIL_SECTIONS = ("ContactTool", "ContactToolConnections", "Results", "BoltTool")

    def _collect_desired_preview_entries(self):
        """
        Fait : calcule la liste des entrees (kind, payload) qui devraient apparaitre dans l'apercu.
        Depend de : self.chk_geometry/chk_mesh, self._section_order, self._sections[...]["rows"], self._multi_result_slides.
        Retourne : list de tuples (kind, payload), dans un ordre naturel (pas encore l'ordre d'apercu).
        """
        entries = []
        if self.chk_geometry.IsChecked:
            entries.append(("general", "Geometrie"))
        if self.chk_mesh.IsChecked:
            entries.append(("general", "Maillage"))
        for name in self._section_order:
            if any(row.checkbox.IsChecked for row in self._sections[name]["rows"]):
                entries.append((name, None))
        # Pas de case a cocher pour ces entrees (ajoutees depuis l'onglet "Slide combinee", voir
        # _on_multi_result_add_to_report) : chacune est toujours "desiree" tant qu'elle n'a pas ete
        # explicitement supprimee depuis sa carte (voir _on_delete_multi_result_slide).
        for cfg in self._multi_result_slides:
            entries.append(("MultiResultSlide", cfg))
        return entries

    def _update_preview(self):
        """
        Fait : met a jour self._preview_order selon les cases/lignes cochees, sans perdre l'ordre du glisser-deposer.
        Depend de : self._collect_desired_preview_entries, self._render_preview.
        Retourne : rien (effet de bord sur self._preview_order et l'affichage de l'apercu).
        """
        # Entrees decochees retirees, nouvelles ajoutees a la fin, le reste garde sa position actuelle.
        desired = self._collect_desired_preview_entries()
        desired_set = set(desired)

        self._preview_order = [entry for entry in self._preview_order if entry in desired_set]
        kept_set = set(self._preview_order)
        for entry in desired:
            if entry not in kept_set:
                self._preview_order.append(entry)
                kept_set.add(entry)

        self._render_preview()

    def _build_preview_list_row(self, primary_text, secondary_lines=None):
        """
        Fait : construit UNE ligne de la liste verticale d'une carte d'apercu (nom + details eventuels).
        Depend de : _shared_resources["CardBorderBrush"], SEARCH_PLACEHOLDER_BRUSH.
        Retourne : Border, la ligne prete a etre ajoutee au conteneur de liste (self._build_preview_list_container).
        """
        # Utilisee de maniere identique pour TOUTES les categories : une ligne par parametre/element,
        # au lieu d'un bloc de texte partage illisible des que plusieurs elements ont chacun
        # plusieurs parametres (vue, coupe, steps, ...).
        inner = StackPanel()

        primary_block = TextBlock()
        primary_block.Text = primary_text
        primary_block.FontSize = 11
        primary_block.TextWrapping = TextWrapping.Wrap
        inner.Children.Add(primary_block)

        for line in (secondary_lines or []):
            detail_block = TextBlock()
            detail_block.Text = line
            detail_block.FontSize = 9
            detail_block.Foreground = SEARCH_PLACEHOLDER_BRUSH
            detail_block.TextWrapping = TextWrapping.Wrap
            detail_block.Margin = Thickness(0, 1, 0, 0)
            inner.Children.Add(detail_block)

        row = Border()
        row.BorderBrush = _shared_resources["CardBorderBrush"]
        row.BorderThickness = Thickness(0, 0, 0, 1)
        row.Padding = Thickness(4, 4, 4, 4)
        row.Child = inner
        return row

    def _build_preview_list_container(self, rows):
        """
        Fait : construit le conteneur de liste verticale d'une carte d'apercu (fond legerement gris,
        colle directement sous le bandeau titre) : hauteur FIXE (PREVIEW_LIST_DEFAULT_HEIGHT, memeS
        pour toutes les cartes) et defilable, avec un fondu en bas (PREVIEW_LIST_FADE_HEIGHT) visible
        UNIQUEMENT si la liste deborde reellement de la hauteur visible (verifie apres layout, voir on_list_loaded).
        Depend de : PREVIEW_LIST_DEFAULT_HEIGHT/FADE_HEIGHT/BACKGROUND(_COLOR), rows deja construites par self._build_preview_list_row.
        Retourne : Grid, le conteneur (liste defilable + fondu superpose) pret a etre ajoute a la carte.
        """
        list_panel = StackPanel()
        # Meme raison que dans _attach_list_fade : sans cet espace, la derniere ligne se retrouve
        # exactement sous la zone de fondu une fois tout en bas de la liste et devient illisible.
        list_panel.Margin = Thickness(0, 0, 0, PREVIEW_LIST_FADE_HEIGHT)
        for row in rows:
            list_panel.Children.Add(row)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Height = PREVIEW_LIST_DEFAULT_HEIGHT
        scroll.Content = list_panel

        list_border = Border()
        list_border.Background = PREVIEW_LIST_BACKGROUND
        list_border.BorderBrush = _shared_resources["CardBorderBrush"]
        list_border.BorderThickness = Thickness(1)
        list_border.Child = scroll

        container = Grid()
        container.Children.Add(list_border)

        fade = Border()
        fade.Height = PREVIEW_LIST_FADE_HEIGHT
        fade.VerticalAlignment = VerticalAlignment.Bottom
        fade.IsHitTestVisible = False
        fade.Visibility = Visibility.Collapsed

        # 3 arrets plutot que 2 (transparent -> oppaque des la moitie du fondu -> oppaque) : le
        # degrade "tient" son opacite plus tot/plus fort au lieu de s'estomper lineairement sur
        # toute la hauteur - rendu plus marque, moins delave, qu'un simple degrade lineaire.
        fade_brush = LinearGradientBrush()
        fade_brush.StartPoint = Point(0, 0)
        fade_brush.EndPoint = Point(0, 1)
        transparent_color = WpfColor.FromArgb(0, PREVIEW_LIST_BACKGROUND_COLOR.R,
                                               PREVIEW_LIST_BACKGROUND_COLOR.G, PREVIEW_LIST_BACKGROUND_COLOR.B)
        mid_color = WpfColor.FromArgb(200, PREVIEW_LIST_BACKGROUND_COLOR.R,
                                       PREVIEW_LIST_BACKGROUND_COLOR.G, PREVIEW_LIST_BACKGROUND_COLOR.B)
        fade_brush.GradientStops.Add(GradientStop(transparent_color, 0))
        fade_brush.GradientStops.Add(GradientStop(mid_color, 0.45))
        fade_brush.GradientStops.Add(GradientStop(PREVIEW_LIST_BACKGROUND_COLOR, 1))
        fade.Background = fade_brush

        def on_list_loaded(sender, e):
            """
            Fait : n'affiche le fondu qu'une fois la hauteur reelle du contenu connue (apres layout).
            Depend de : scroll.ScrollableHeight, fade capture par la fermeture.
            Retourne : rien (effet de bord sur fade.Visibility).
            """
            fade.Visibility = Visibility.Visible if scroll.ScrollableHeight > 0 else Visibility.Collapsed
        scroll.Loaded += on_list_loaded

        container.Children.Add(fade)
        return container

    def _build_preview_card_content(self, title, chips, order_number):
        """
        Fait : construit le contenu d'une carte d'apercu - bandeau titre+badge (fond blanc, herite
        de la carte), colle directement au-dessus du conteneur de liste verticale des elements
        selectionnes (self._build_preview_list_container).
        Depend de : _shared_resources["AccentBrush"], self._build_preview_list_container (via chips deja construits par self._build_preview_list_row).
        Retourne : tuple (StackPanel contenu, TextBlock badge) - le badge est renvoye a part pour etre renumerote sans reconstruire la carte.
        """
        content = StackPanel()

        title_row = StackPanel()
        title_row.Orientation = Orientation.Horizontal
        title_row.Margin = Thickness(0, 0, 0, 6)

        badge = Border()
        badge.Width = 20
        badge.Height = 20
        badge.CornerRadius = CornerRadius(0)
        badge.Background = _shared_resources["AccentBrush"]
        badge.Margin = Thickness(0, 0, 8, 0)

        badge_text = TextBlock()
        badge_text.Text = str(order_number)
        badge_text.Foreground = Brushes.White
        badge_text.FontSize = 11
        badge_text.FontWeight = FontWeights.Bold
        badge_text.HorizontalAlignment = HorizontalAlignment.Center
        badge_text.VerticalAlignment = VerticalAlignment.Center
        badge.Child = badge_text
        title_row.Children.Add(badge)

        title_block = TextBlock()
        title_block.Text = title
        title_block.FontWeight = FontWeights.Bold
        title_block.TextWrapping = TextWrapping.Wrap
        title_block.VerticalAlignment = VerticalAlignment.Center
        title_row.Children.Add(title_block)

        content.Children.Add(title_row)

        # Toujours ajoute, meme si chips est vide (ex : Geometrie/Maillage sans vue configuree) :
        # garde une hauteur de carte uniforme dans tous les cas (voir PREVIEW_LIST_DEFAULT_HEIGHT),
        # plutot qu'une carte anormalement courte des qu'il n'y a rien a lister.
        content.Children.Add(self._build_preview_list_container(chips))

        return content, badge_text

    def _build_preview_card(self, entry, index):
        """
        Fait : construit une carte d'apercu complete (fond, bordure, ombre) pour une entree de self._preview_order.
        Depend de : self._sections, self.FULL_DETAIL_SECTIONS, self._geometry_view_config/_mesh_view_config, self._build_preview_card_content, self._begin_potential_drag, build_row_display_name (05_interactive_slides.py, pour MultiResultSlide).
        Retourne : Border, la carte prete a etre ajoutee a panelPreview.
        """
        # Seule LA CARTE DE CATEGORIE elle-meme est glissable/receptrice (voir _begin_potential_drag) ;
        # les chips internes ne le sont pas.
        kind, payload = entry

        delete_handler = None

        if kind == "general":
            chips = []
            if payload == "Geometrie":
                if self._geometry_view_config.view_name:
                    chips.append(self._build_preview_list_row("vue=" + self._geometry_view_config.view_name))
            elif payload == "Maillage":
                table_mode = "Tableau complet" if self._mesh_table_full else "Tableau par defaut"
                chips.append(self._build_preview_list_row(table_mode))
                if self._mesh_view_config.view_name:
                    chips.append(self._build_preview_list_row("vue=" + self._mesh_view_config.view_name))
            content, badge = self._build_preview_card_content(payload, chips, index + 1)
        elif kind == "MultiResultSlide":
            # Pas de case a cocher pour ce type d'entree (voir _on_multi_result_add_to_report) : la
            # carte porte elle-meme un bouton "Supprimer" pour la retirer de l'apercu/de la generation.
            chips = []
            for cell_cfg in payload.cell_configs:
                full_text = build_row_display_name(cell_cfg)
                parts = full_text.split(" | ")
                chips.append(self._build_preview_list_row(parts[0], parts[1:]))
            title = "Slide combinee ({} resultats)".format(payload.template_count)
            content, badge = self._build_preview_card_content(title, chips, index + 1)
            delete_handler = self._make_multi_result_delete_handler(payload)
        else:
            section = self._sections[kind]
            checked_rows = [row for row in section["rows"] if row.checkbox.IsChecked]
            if kind in self.FULL_DETAIL_SECTIONS:
                # Categories "resultats" : un chip par resultat, avec son detail complet (vue, coupe, steps, ...).
                chips = []
                for row in checked_rows:
                    full_text = row.display_name_func(row.row_config)
                    parts = full_text.split(" | ")
                    chips.append(self._build_preview_list_row(parts[0], parts[1:]))
            else:
                # Categories "contexte" (pieces, BC, BP, contacts, solution info, analyses) : un
                # chip par nom, sans detail - via display_name_func()[0] (et non obj.Name brut)
                # pour que le suffixe d'analyse (voir analysis_suffix) apparaisse aussi ici pour
                # Bolt Pretension / Solution Information sur un projet multi-analyses.
                chips = [self._build_preview_list_row(row.display_name_func(row.row_config).split(" | ")[0])
                         for row in checked_rows]
            content, badge = self._build_preview_card_content(section["label"], chips, index + 1)

        if delete_handler is not None:
            btn_delete = _themed_button()
            btn_delete.Content = "Supprimer"
            btn_delete.FontSize = 10
            btn_delete.Padding = Thickness(6, 1, 6, 1)
            btn_delete.Margin = Thickness(10, 0, 0, 0)
            btn_delete.VerticalAlignment = VerticalAlignment.Center
            btn_delete.Click += delete_handler
            content.Children[0].Children.Add(btn_delete)  # content.Children[0] = title_row (StackPanel horizontal, voir _build_preview_card_content)

        self._entry_to_badge[entry] = badge

        card = Border()
        card.Background = CARD_NORMAL_BACKGROUND
        card.BorderBrush = _shared_resources["CardBorderBrush"]
        card.BorderThickness = Thickness(1)
        card.CornerRadius = CornerRadius(0)
        card.Padding = Thickness(10)
        card.Margin = Thickness(4)
        card.Width = CARD_WIDTH
        card.Cursor = Cursors.SizeAll
        card.Tag = entry
        card.Child = content

        def on_mouse_enter(sender, e):
            """
            Fait : passe la carte en bleu tres clair au survol, sauf pendant un glisser en cours.
            Depend de : self._drag_active, CARD_HOVER_BACKGROUND, card capture par la fermeture.
            Retourne : rien (effet de bord sur card.Background).
            """
            if not self._drag_active:
                card.Background = CARD_HOVER_BACKGROUND

        def on_mouse_leave(sender, e):
            """
            Fait : restaure le fond normal de la carte quand la souris la quitte.
            Depend de : CARD_NORMAL_BACKGROUND, card capture par la fermeture.
            Retourne : rien (effet de bord sur card.Background).
            """
            card.Background = CARD_NORMAL_BACKGROUND

        def on_preview_mouse_down(sender, e):
            """
            Fait : enregistre le point de depart d'un glisser potentiel sur cette carte, sauf si le
            clic provient d'un bouton imbrique (ex : "Supprimer" d'une slide combinee) - sinon
            CaptureMouse() sur panelPreview empeche le Click du bouton de se declencher normalement.
            Depend de : self._begin_potential_drag, self._is_button_descendant, card/entry captures par la fermeture.
            Retourne : rien (effet de bord : initialise l'etat de glisser).
            """
            if self._is_button_descendant(e.OriginalSource):
                return
            self._begin_potential_drag(card, entry, e)

        card.MouseEnter += on_mouse_enter
        card.MouseLeave += on_mouse_leave
        card.PreviewMouseLeftButtonDown += on_preview_mouse_down

        return card

    def _render_preview(self):
        """
        Fait : reconstruit entierement les cartes du WrapPanel panelPreview a partir de self._preview_order.
        Depend de : self._build_preview_card, self._preview_order.
        Retourne : rien (effet de bord sur self.panel_preview et self._entry_to_card/_entry_to_badge).
        """
        self.panel_preview.Children.Clear()
        self._entry_to_card = {}
        self._entry_to_badge = {}

        for index, entry in enumerate(self._preview_order):
            card = self._build_preview_card(entry, index)
            self._entry_to_card[entry] = card
            self.panel_preview.Children.Add(card)

        if not self._preview_order:
            placeholder = TextBlock()
            placeholder.Text = "(Aucune slide selectionnee)"
            placeholder.Foreground = SEARCH_PLACEHOLDER_BRUSH
            placeholder.Margin = Thickness(6)
            self.panel_preview.Children.Add(placeholder)

    # --- Glisser-deposer des cartes d'apercu ---
    # La souris est capturee par panelPreview (pas par la carte elle-meme) : ainsi, reordonner
    # les enfants du WrapPanel pendant le glisser ne fait jamais perdre la capture, meme si la
    # carte source est brievement retiree/reinseree. Un Popup sans decoration ("fantome", copie
    # visuelle via VisualBrush) suit la souris ; les autres cartes se decalent en direct des que
    # le curseur survole une autre carte.

    def _begin_potential_drag(self, card, entry, e):
        """
        Fait : enregistre le point de depart d'un glisser potentiel et capture la souris sur panelPreview.
        Depend de : self.panel_preview.CaptureMouse().
        Retourne : rien (effet de bord : initialise self._drag_pending_card/_drag_pending_entry/_drag_start_point).
        """
        self._drag_pending_card = card
        self._drag_pending_entry = entry
        self._drag_start_point = e.GetPosition(self.panel_preview)
        self.panel_preview.CaptureMouse()

    def _on_preview_panel_mouse_move(self, sender, e):
        """
        Fait : demarre le glisser au-dela d'un seuil de mouvement, puis fait suivre le fantome et reordonne au survol.
        Depend de : self._drag_pending_card/_drag_pending_entry/_drag_start_point, self._start_drag/_update_drag_ghost_position/_update_drag_hover.
        Retourne : rien (effet de bord : declenche le glisser ou met a jour sa position).
        """
        if e.LeftButton != MouseButtonState.Pressed or self._drag_pending_entry is None:
            return

        current_point = e.GetPosition(self.panel_preview)

        if not self._drag_active:
            delta_x = abs(current_point.X - self._drag_start_point.X)
            delta_y = abs(current_point.Y - self._drag_start_point.Y)
            if delta_x < 6 and delta_y < 6:
                return
            self._start_drag(self._drag_pending_card, self._drag_pending_entry)

        self._update_drag_ghost_position(e)
        self._update_drag_hover(current_point)

    def _on_preview_panel_mouse_up(self, sender, e):
        """
        Fait : relache la capture souris et termine proprement le glisser en cours (s'il y en a un).
        Depend de : self.panel_preview.ReleaseMouseCapture(), self._drag_active, self._end_drag.
        Retourne : rien (effet de bord : remet a zero l'etat de glisser en attente).
        """
        self.panel_preview.ReleaseMouseCapture()
        if self._drag_active:
            self._end_drag()
        self._drag_pending_card = None
        self._drag_pending_entry = None
        self._drag_start_point = None

    def _start_drag(self, card, entry):
        """
        Fait : demarre effectivement le glisser (estompe la carte source, ouvre le Popup fantome).
        Depend de : Border/VisualBrush/Popup (WPF), card.ActualWidth/ActualHeight.
        Retourne : rien (effet de bord : initialise self._drag_active/_drag_entry/_drag_source_card/_drag_popup).
        """
        self._drag_active = True
        self._drag_entry = entry
        self._drag_source_card = card
        card.Opacity = 0.25

        ghost = Border()
        ghost.Width = card.ActualWidth
        ghost.Height = card.ActualHeight
        ghost.Background = VisualBrush(card)
        ghost.Opacity = 0.85

        popup = Popup()
        popup.AllowsTransparency = True
        popup.Placement = PlacementMode.Absolute
        popup.IsHitTestVisible = False
        popup.Focusable = False
        popup.Child = ghost
        popup.IsOpen = True
        self._drag_popup = popup

    def _update_drag_ghost_position(self, e):
        """
        Fait : deplace le Popup fantome pour qu'il reste centre sur le curseur.
        Depend de : self._drag_popup, self._drag_source_card, self.panel_preview.PointToScreen.
        Retourne : rien (effet de bord sur self._drag_popup.HorizontalOffset/VerticalOffset).
        """
        screen_point = self.panel_preview.PointToScreen(e.GetPosition(self.panel_preview))
        card = self._drag_source_card
        self._drag_popup.HorizontalOffset = screen_point.X - card.ActualWidth / 2.0
        self._drag_popup.VerticalOffset = screen_point.Y - card.ActualHeight / 2.0

    def _is_button_descendant(self, element):
        """
        Fait : determine si element est un Button ou est contenu dans un Button (ex : le TextBlock
        auto-genere pour Content="Supprimer"), en remontant l'arbre visuel.
        Depend de : VisualTreeHelper.GetParent (WPF).
        Retourne : bool, True des qu'un Button est trouve sur le chemin.
        """
        node = element
        while node is not None:
            if isinstance(node, Button):
                return True
            node = VisualTreeHelper.GetParent(node)
        return False

    def _find_ancestor_card(self, element):
        """
        Fait : remonte l'arbre visuel depuis un resultat de hit-test jusqu'a la Border d'une carte.
        Depend de : VisualTreeHelper.GetParent (WPF), la convention Tag = entry sur les cartes (voir _build_preview_card).
        Retourne : Border ou None, la carte trouvee (Tag non None) ou None si aucune sur le chemin.
        """
        node = element
        while node is not None:
            if isinstance(node, Border) and node.Tag is not None:
                return node
            node = VisualTreeHelper.GetParent(node)
        return None

    def _update_drag_hover(self, position):
        """
        Fait : deplace l'entree glissee dans self._preview_order si le curseur survole une AUTRE carte.
        Depend de : self._find_ancestor_card, self._preview_order, self._reorder_children_to_match_preview_order.
        Retourne : rien (effet de bord sur self._preview_order et l'affichage si un deplacement a lieu).
        """
        hit = self.panel_preview.InputHitTest(position)
        target_card = self._find_ancestor_card(hit) if hit else None
        if target_card is None or target_card is self._drag_source_card:
            return

        target_entry = target_card.Tag
        if self._drag_entry not in self._preview_order or target_entry not in self._preview_order:
            return

        source_index = self._preview_order.index(self._drag_entry)
        target_index = self._preview_order.index(target_entry)
        if source_index == target_index:
            return

        moved = self._preview_order.pop(source_index)
        self._preview_order.insert(target_index, moved)
        self._reorder_children_to_match_preview_order()

    def _reorder_children_to_match_preview_order(self):
        """
        Fait : reordonne panelPreview.Children pour refleter self._preview_order sans recreer les cartes.
        Depend de : self._entry_to_card/_entry_to_badge, self._preview_order.
        Retourne : rien (effet de bord sur panelPreview.Children et les badges numerotes).
        """
        # Contrairement a _render_preview, ne recree pas les cartes : essentiel pendant un glisser
        # en cours, pour ne pas perdre les gestionnaires d'evenements ni la capture souris (capturee
        # sur le panel, pas sur la carte). Renumerote aussi les badges pour qu'ils restent justes
        # pendant le glisser, pas seulement une fois relache.
        children = self.panel_preview.Children
        for target_index, entry in enumerate(self._preview_order):
            card = self._entry_to_card.get(entry)
            if card is None:
                continue
            current_index = children.IndexOf(card)
            if current_index != target_index:
                children.RemoveAt(current_index)
                children.Insert(target_index, card)

            badge = self._entry_to_badge.get(entry)
            if badge is not None:
                badge.Text = str(target_index + 1)

    def _end_drag(self):
        """
        Fait : termine le glisser en cours (ferme le fantome, restaure l'opacite de la carte source).
        Depend de : self._drag_popup, self._drag_source_card.
        Retourne : rien (effet de bord : remet a zero l'etat de glisser).
        """
        self._drag_active = False
        if self._drag_popup is not None:
            self._drag_popup.IsOpen = False
            self._drag_popup = None
        if self._drag_source_card is not None:
            self._drag_source_card.Opacity = 1.0
        self._drag_source_card = None
        self._drag_entry = None

    def _get_checked_row_configs(self, name):
        """
        Fait : recupere les row_config des lignes cochees de la section 'name'.
        Depend de : self._sections[name]["rows"].
        Retourne : list de row_config, ceux dont la CheckBox est cochee.
        """
        return [row.row_config for row in self._sections[name]["rows"] if row.checkbox.IsChecked]

    # --- Handlers de cases a cocher simples ---

    def _on_simple_toggle(self, sender, e):
        """
        Fait : rafraichit l'apercu quand une case simple (Geometrie/Maillage) change d'etat.
        Depend de : self._update_preview.
        Retourne : rien (effet de bord sur l'apercu).
        """
        self._update_preview()

    def _on_geometry_view_click(self, sender, e):
        """
        Fait : ouvre le panneau lateral global de selection de vue pour la slide Geometrie (bouton "Parametres" de la carte).
        Depend de : self._open_config_panel, self._geometry_view_config, self._refresh_general_slide_status, self._update_preview.
        Retourne : rien (effet de bord : affiche borderConfigPanel ; met a jour self._geometry_view_config.view_name, le statut de la carte et l'apercu si "Appliquer").
        """
        def refresh():
            self._refresh_general_slide_status()
            self._update_preview()
        self._open_config_panel("mesh_part", self._geometry_view_config, refresh)

    def _on_mesh_view_click(self, sender, e):
        """
        Fait : ouvre le panneau lateral global de selection de vue pour la slide Maillage (bouton "Parametres" de la carte).
        Depend de : self._open_config_panel, self._mesh_view_config, self._refresh_general_slide_status, self._update_preview.
        Retourne : rien (effet de bord : affiche borderConfigPanel ; met a jour self._mesh_view_config.view_name, le statut de la carte et l'apercu si "Appliquer").
        """
        def refresh():
            self._refresh_general_slide_status()
            self._update_preview()
        self._open_config_panel("mesh_part", self._mesh_view_config, refresh)

    # --- Tout (de)selectionner, par onglet ---

    def _set_group_checked(self, group_key, checked):
        """
        Fait : (de)coche toutes les lignes de toutes les sections d'un onglet (group_key).
        Depend de : self._section_order, self._sections, self._update_preview.
        Retourne : rien (effet de bord sur les CheckBox des sections concernees et sur l'apercu).
        """
        for name in self._section_order:
            section = self._sections[name]
            if section["group_key"] != group_key:
                continue
            for row in section["rows"]:
                row.checkbox.IsChecked = checked
        self._update_preview()

    def _set_section_checked(self, name, checked):
        """
        Fait : (de)coche toutes les lignes d'UNE SEULE section (une zone de selection precise).
        Depend de : self._sections[name]["rows"], self._update_preview.
        Retourne : rien (effet de bord sur les CheckBox de la section et sur l'apercu).
        """
        for row in self._sections[name]["rows"]:
            row.checkbox.IsChecked = checked
        self._update_preview()

    def _make_zone_toggle_handler(self, name, checked):
        """
        Fait : ferme name/checked par valeur pour produire le handler Click du bouton 'Tout'/'Aucun' d'une zone.
        Depend de : self._set_section_checked.
        Retourne : function, le handler(sender, e) a cabler sur le bouton de zone.
        """
        def handler(sender, e):
            """
            Fait : (de)coche toutes les lignes de la zone associee au bouton.
            Depend de : self._set_section_checked, name/checked captures par la fermeture.
            Retourne : rien (effet de bord sur les CheckBox de la zone).
            """
            self._set_section_checked(name, checked)
        return handler

    def _wire_zone_select_buttons(self):
        """
        Fait : cable les boutons "Tout"/"Aucun" de chaque zone de selection (en-tete de carte XAML).
        Depend de : self._section_order, self.window.FindName, self._make_zone_toggle_handler.
        Retourne : rien (effet de bord : cable les Click des boutons btnZoneCheck{name}/btnZoneUncheck{name}).
        """
        # Raccourci par zone, en plus des boutons "Tout (de)selectionner" existants qui agissent
        # sur tout un onglet a la fois (voir _set_group_checked).
        for name in self._section_order:
            check_btn = self.window.FindName("btnZoneCheck" + name)
            uncheck_btn = self.window.FindName("btnZoneUncheck" + name)
            if check_btn is not None:
                check_btn.Click += self._make_zone_toggle_handler(name, True)
            if uncheck_btn is not None:
                uncheck_btn.Click += self._make_zone_toggle_handler(name, False)

    def _on_check_all_general(self, sender, e):
        """
        Fait : coche toutes les lignes de l'onglet "Slides generales" (y compris Geometrie/Maillage).
        Depend de : self.chk_geometry/chk_mesh, self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self.chk_geometry.IsChecked = True
        self.chk_mesh.IsChecked = True
        self._set_group_checked("general", True)

    def _on_uncheck_all_general(self, sender, e):
        """
        Fait : decoche toutes les lignes de l'onglet "Slides generales" (y compris Geometrie/Maillage).
        Depend de : self.chk_geometry/chk_mesh, self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self.chk_geometry.IsChecked = False
        self.chk_mesh.IsChecked = False
        self._set_group_checked("general", False)

    def _on_check_all_conditions(self, sender, e):
        """
        Fait : coche toutes les lignes de l'onglet "Conditions et contacts".
        Depend de : self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self._set_group_checked("conditions", True)

    def _on_uncheck_all_conditions(self, sender, e):
        """
        Fait : decoche toutes les lignes de l'onglet "Conditions et contacts".
        Depend de : self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self._set_group_checked("conditions", False)

    def _on_check_all_results(self, sender, e):
        """
        Fait : coche toutes les lignes de l'onglet "Categories de resultats".
        Depend de : self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self._set_group_checked("results", True)

    def _on_uncheck_all_results(self, sender, e):
        """
        Fait : decoche toutes les lignes de l'onglet "Categories de resultats".
        Depend de : self._set_group_checked.
        Retourne : rien (effet de bord sur les CheckBox de l'onglet).
        """
        self._set_group_checked("results", False)

    # --- Utilitaires : suppression des figures / creation des vues de base ---

    def _on_delete_figures(self, sender, e):
        """
        Fait : supprime les figures obsoletes generees par les exports precedents (bouton dedie).
        Depend de : remove_stale_figures (05_interactive_slides.py).
        Retourne : rien (effet de bord : supprime des fichiers, affiche une MessageBox si echec).
        """
        try:
            remove_stale_figures()
            print "Figures supprimees."
        except Exception as ex:
            print "ERREUR pendant la suppression des figures : " + str(ex)
            MessageBox.Show("Erreur pendant la suppression des figures :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_reset_legends(self, sender, e):
        """
        Fait : reinitialise les legendes de resultats appliquees dans Mechanical (bouton dedie).
        Depend de : reset_legend (05_interactive_slides.py).
        Retourne : rien (effet de bord : modifie l'etat des legendes, affiche une MessageBox si echec).
        """
        try:
            reset_legend()
            print "Legendes reinitialisees."
        except Exception as ex:
            print "ERREUR pendant la reinitialisation des legendes : " + str(ex)
            MessageBox.Show("Erreur pendant la reinitialisation des legendes :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_create_basic_views(self, sender, e):
        """
        Fait : cree les vues de base dans le View Manager et rafraichit les listes de vues/coupes.
        Depend de : create_basic_views/collect_views/collect_section_planes/section_plane_label (05_interactive_slides.py).
        Retourne : rien (effet de bord : cree des vues Mechanical, met a jour self._views/_section_planes/_section_plane_labels).
        """
        try:
            created = create_basic_views()
            self._views = collect_views()
            self._section_planes = collect_section_planes()
            self._section_plane_labels = [
                section_plane_label(sp, i) for i, sp in enumerate(self._section_planes)
            ]
            if created:
                print "{} vue(s) de base creee(s) : {}.".format(len(created), ", ".join(created))
            else:
                print "Aucune vue de base n'a pu etre creee (voir la console Mechanical)."
        except Exception as ex:
            print "ERREUR pendant la creation des vues de base : " + str(ex)
            MessageBox.Show("Erreur pendant la creation des vues de base :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_export_3d(self, sender, e):
        """
        Fait : exporte en .avz la vue 3D de tous les resultats et Contact/Bolt Tool (branche Solution) du projet.
        Depend de : export_all_3d_views, EXPORT_3D_FOLDER (00_constants.py), _print_console_banner.
        Retourne : rien (effet de bord : cree des fichiers .avz dans EXPORT_3D_FOLDER, affiche une MessageBox si echec).
        """
        try:
            _print_console_banner("EXPORT 3D (.avz) EN COURS...")
            exported_count = export_all_3d_views(EXPORT_3D_FOLDER)
            if exported_count:
                _print_console_banner("{} VUE(S) 3D EXPORTEE(S)".format(exported_count))
                print "Fichiers .avz disponibles dans : " + EXPORT_3D_FOLDER
            else:
                _print_console_banner("AUCUNE VUE 3D EXPORTEE")
                print "Aucun resultat / Contact Tool / Bolt Tool (branche Solution) trouve a exporter."
        except Exception as ex:
            print "ERREUR pendant l'export 3D : " + str(ex)
            MessageBox.Show("Erreur pendant l'export 3D :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    def _make_multi_result_delete_handler(self, cfg):
        """
        Fait : ferme cfg par valeur pour produire le handler du bouton "Supprimer" d'une carte MultiResultSlide.
        Depend de : self._on_delete_multi_result_slide.
        Retourne : function, le handler(sender, e) a cabler sur btn_delete.Click.
        """
        def handler(sender, e):
            self._on_delete_multi_result_slide(cfg)
        return handler

    def _on_delete_multi_result_slide(self, cfg):
        """
        Fait : retire une slide combinee "differents resultats" de l'apercu et de la generation.
        Depend de : self._multi_result_slides, self._update_preview.
        Retourne : rien (effet de bord sur self._multi_result_slides et l'apercu).
        """
        if cfg in self._multi_result_slides:
            self._multi_result_slides.remove(cfg)
        self._update_preview()

    def _on_close(self, sender, e):
        """
        Fait : ferme la fenetre principale de l'application (bouton Fermer).
        Depend de : self.window.
        Retourne : rien (effet de bord : ferme self.window).
        """
        self.window.Close()

    # --- Generation du rapport ---

    def _on_generate(self, sender, e):
        """
        Fait : genere le rapport PowerPoint en respectant l'ordre de self._preview_order.
        Depend de : PPTReportBuilder, build_*_slides/create_*_slide (04_slides.py/05_interactive_slides.py), apply_view_if_exists/self._geometry_view_config/_mesh_view_config, self._update_generation_progress.
        Retourne : rien (effet de bord : cree le fichier PPTX, met a jour l'UI de statut, affiche une MessageBox si echec).
        """
        # Chaque carte est traitee une a une, dans l'ordre : une carte "general" ajoute sa slide
        # unique, une carte de section ajoute TOUTES les slides de sa categorie d'un coup (fonction
        # par lot de 05_interactive_slides.py). La granularite de reordonnancement (et de la barre
        # de progression) est donc la carte/categorie, pas la slide individuelle.
        # PowerPoint est visible pendant toute cette methode (voir PPTReportBuilder.__init__ - le
        # garder invisible s'est revele instable sur un rapport avec beaucoup de slides) ; la
        # fenetre WPF reste reactive grace a SWF.Application.DoEvents() (_update_generation_progress).
        if not self._preview_order:
            MessageBox.Show("Aucune slide selectionnee : cochez au moins une option avant de generer le rapport.",
                             "Rien a generer", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        total = len(self._preview_order)
        report = None
        self._reset_generation_ui(total)

        try:
            remove_stale_figures()

            _print_console_banner("GENERATION DU RAPPORT EN COURS...")
            print "Ouverture du template PowerPoint..."
            report = PPTReportBuilder(TEMPLATE_PATH)

            for index, (kind, payload) in enumerate(self._preview_order):
                if kind == "general":
                    if payload == "Geometrie":
                        apply_view_if_exists(self._geometry_view_config.view_name, self._views)
                        create_geometry_slide(report)
                    elif payload == "Maillage":
                        apply_view_if_exists(self._mesh_view_config.view_name, self._views)
                        build_mesh_slide(report, self._mesh_table_full)
                    print "Slide ajoutee : " + payload
                    self._update_generation_progress(index + 1, total)
                    continue

                if kind == "MultiResultSlide":
                    template = get_multi_step_template(payload.template_count)
                    build_multi_result_slide(report, template, payload.cell_configs, self._views,
                                              self._section_planes, self._section_plane_labels)
                    print "Slide combinee multi-resultats ajoutee ({} resultats).".format(len(payload.cell_configs))
                    self._update_generation_progress(index + 1, total)
                    continue

                selected = self._get_checked_row_configs(kind)
                if not selected:
                    self._update_generation_progress(index + 1, total)
                    continue

                if kind == "AnalysisContext":
                    build_analysis_context_slides(report, selected, self._views)
                elif kind == "GeometryParts":
                    build_geometry_part_slides(report, selected, self._bodies,
                                                self._views, self._section_planes, self._section_plane_labels)
                elif kind == "MeshParts":
                    build_mesh_part_slides(report, selected, self._bodies, self._views)
                elif kind == "BoundaryConditions":
                    build_bc_slides(report, selected, self._views, self._section_planes, self._section_plane_labels)
                elif kind == "BoltPretension":
                    build_bp_slides(report, selected, self._views, self._section_planes, self._section_plane_labels)
                elif kind == "Contacts":
                    build_contact_summary_slide(report, selected)
                elif kind == "SolutionInfo":
                    build_solution_info_slides(report, selected)
                elif kind == "ContactTool":
                    build_result_slides(report, selected, "-- Contact Tool Results --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "ContactToolConnections":
                    build_result_slides(report, selected, "-- Connexion : Contact Tool --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "BoltTool":
                    build_result_slides(report, selected, "-- Bolt Tool --",
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)
                elif kind == "Results":
                    # Sous-titre generique en multi-analyses : self._analysis.Name (Analyses[0])
                    # serait trompeur pour des resultats venant d'une autre analyse - celle-ci
                    # est deja indiquee dans le TITRE de chaque slide (voir analysis_suffix).
                    results_subtitle = "-- Resultats --" if self._multi_analysis else self._analysis.Name
                    build_result_slides(report, selected, results_subtitle,
                                         self._views, self._section_planes, self._section_plane_labels,
                                         self._analysis)

                print "{} slide(s) {} ajoutee(s).".format(len(selected), self._sections[kind]["label"])
                self._update_generation_progress(index + 1, total)

            self.btn_generate.IsEnabled = False
            report.keep_open()  # ni Save(), ni Close()/Quit() - le rapport reste ouvert et non enregistre dans PowerPoint
            self._last_report_path = report.working_copy_path
            self._mark_report_ready(report.working_copy_path)
            _print_console_banner("RAPPORT GENERE AVEC SUCCES")
            print "Rapport disponible dans l'onglet Fichiers : " + report.working_copy_path
        except Exception as ex:
            _print_console_banner("ERREUR PENDANT LA GENERATION DU RAPPORT")
            print str(ex)
            if report is not None:
                try:
                    report.close()
                except Exception as close_ex:
                    print "Fermeture de PowerPoint impossible : " + str(close_ex)
            MessageBox.Show("Erreur pendant la generation du rapport :\n" + str(ex),
                             "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)

    # --- Branchement des evenements ---

    def _wire_events(self):
        """
        Fait : cable tous les evenements de la fenetre principale (boutons, cases, glisser-depose).
        Depend de : tous les controles trouves par self._find_controls, les handlers self._on_*.
        Retourne : rien (effet de bord : abonne les handlers aux evenements WPF).
        """
        self.btn_delete_figures.Click += self._on_delete_figures
        self.btn_reset_legends.Click += self._on_reset_legends
        self.btn_create_views.Click += self._on_create_basic_views
        self.btn_export_3d.Click += self._on_export_3d
        self.btn_multi_result_add_to_report.Click += self._on_multi_result_add_to_report

        # La souris est capturee sur panelPreview (pas sur chaque carte) pendant un glisser-depose :
        # ces deux gestionnaires doivent donc vivre ici, une seule fois (voir _begin_potential_drag).
        self.panel_preview.MouseMove += self._on_preview_panel_mouse_move
        self.panel_preview.PreviewMouseLeftButtonUp += self._on_preview_panel_mouse_up

        self.chk_geometry.Checked += self._on_simple_toggle
        self.chk_geometry.Unchecked += self._on_simple_toggle
        self.btn_geometry_view.Click += self._on_geometry_view_click
        self.chk_mesh.Checked += self._on_simple_toggle
        self.chk_mesh.Unchecked += self._on_simple_toggle
        self.btn_mesh_view.Click += self._on_mesh_view_click

        self.btn_check_all_general.Click += self._on_check_all_general
        self.btn_uncheck_all_general.Click += self._on_uncheck_all_general
        self.btn_check_all_conditions.Click += self._on_check_all_conditions
        self.btn_uncheck_all_conditions.Click += self._on_uncheck_all_conditions
        self.btn_check_all_results.Click += self._on_check_all_results
        self.btn_uncheck_all_results.Click += self._on_uncheck_all_results

        self.btn_generate.Click += self._on_generate
        self.btn_close.Click += self._on_close

        self.btn_report_view.Click += self._on_view_report
        self.btn_report_show_in_folder.Click += self._on_show_report_in_folder


# --- SECTION 8 - Point d'entree ---

_xaml_path = os.path.join(PROJECT_DIR, "AnsysReportGenerator_WPF.xaml")
_app = ReportGeneratorApp(_xaml_path)
_app.window.ShowDialog()
