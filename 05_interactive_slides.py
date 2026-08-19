# 05_interactive_slides.py : logique de support de la selection interactive (AnsysReportGenerator_WPF.py) - classes de configuration par ligne, collecte des vues/coupes/analyses, et constructeurs de slides limites a la selection de l'utilisateur (au lieu de traiter toujours toute une categorie comme 04_slides.py). Depend de 00_constants.py, 01_data_export.py, 02_image_export.py, 03_ppt_utils.py, 04_slides.py (doivent etre executes avant ce fichier).

import csv


def remove_stale_figures():
    """
    Fait : supprime les objets Figure residuels d'une precedente generation du rapport.
    Depend de : DataModel.GetObjectsByType/Remove, Transaction (Ansys.ACT.Mechanical).
    Retourne : rien (effet de bord : nettoie l'arbre avant une nouvelle generation).
    """
    try:
        # Transaction(True) differe le rafraichissement de l'arbre/viewport jusqu'a la fin de la suppression en masse.
        with Transaction(True):
            DataModel.Remove(DataModel.GetObjectsByType(DataModelObjectCategory.Figure))
    except Exception as e:
        print "Suppression des figures existantes impossible : " + str(e)


# Colonne 2 (ViewOrientationType) corrigee suite a une verification visuelle dans Mechanical :
# l'orientation de la piece fait que le resultat reel de chaque enum ne correspond pas a son nom
# naturel dans l'API .NET - ex: ViewOrientationType.Back est ce qui produit reellement la vue X+.
BASIC_VIEW_ORIENTATIONS = [
    ("X+", "Back"),
    ("X-", "Front"),
    ("Y+", "Left"),
    ("Y-", "Right"),
    ("Z+", "Top"),
    ("Z-", "Bottom"),
    ("ISO", "Iso"),
]


def create_basic_views():
    """
    Fait : cree 7 vues de base (X+, X-, Y+, Y-, Z+, Z-, ISO) dans le View Manager.
    Depend de : Ansys.Mechanical.DataModel.Enums.ViewOrientationType, ExtAPI.Graphics.Camera/ModelViewManager.
    Retourne : list, les noms des vues effectivement creees.
    """
    from Ansys.Mechanical.DataModel.Enums import ViewOrientationType

    cam = ExtAPI.Graphics.Camera
    mvm = ExtAPI.Graphics.ModelViewManager

    # Transaction(True) differe le rafraichissement d'interface jusqu'a la fin des 7 creations : CreateView() ne depend que de l'etat de la camera, pas d'un rendu deja affiche.
    created = []
    with Transaction(True):
        for name, orientation_attr in BASIC_VIEW_ORIENTATIONS:
            try:
                orientation = getattr(ViewOrientationType, orientation_attr)
                cam.SetSpecificViewOrientation(orientation)
                cam.SetFit()
                mvm.CreateView(name)
                created.append(name)
                print "Vue creee : " + name
            except Exception as e:
                print "Impossible de creer la vue {} : {}".format(name, str(e))

    return created


def export_object_3d_view(obj, directory):
    """
    Fait : active un objet et exporte sa vue 3D interactive (.avz) via le View Manager.
    Depend de : obj.Activate, ExtAPI.Graphics.ModelViewManager.Capture3DImage, get_unique_file_path/safe_file_name (00_constants.py).
    Retourne : str, le chemin du fichier .avz genere, ou None en cas d'erreur.
    """
    try:
        obj.Activate()
        avz_path = get_unique_file_path(directory, safe_file_name(obj.Name), ".avz")
        ExtAPI.Graphics.ModelViewManager.Capture3DImage(avz_path)
        print "Vue 3D exportee : " + avz_path
        return avz_path
    except Exception as e:
        print "Export 3D impossible pour {} : {}".format(obj.Name, str(e))
        return None


def collect_3d_exportable_objects(analysis):
    """
    Fait : liste les objets a exporter en 3D pour une analyse (resultats simples + Contact Tool/Bolt Tool de la branche Solution).
    Depend de : collect_all_results, collect_contact_tool_results, collect_bolt_tool_results.
    Retourne : list, les objets Mechanical exportables en .avz pour cette analyse.
    """
    # Contact Tool / Bolt Tool de la branche Connections (definition, sans resultat 3D a proprement
    # parler) sont volontairement exclus : seuls ceux de la branche Solution ont un sens ici.
    objects = list(collect_all_results(analysis))
    objects.extend(collect_contact_tool_results(analysis))
    objects.extend(collect_bolt_tool_results(analysis))
    return objects


def export_all_3d_views(directory):
    """
    Fait : exporte en .avz la vue 3D de tous les resultats simples et Contact/Bolt Tool (branche Solution) de toutes les analyses du projet.
    Depend de : ensure_folder_exists (00_constants.py), collect_analyses, collect_3d_exportable_objects, export_object_3d_view.
    Retourne : int, le nombre de fichiers .avz effectivement exportes.
    """
    ensure_folder_exists(directory)
    exported_count = 0
    for analysis in collect_analyses():
        for obj in collect_3d_exportable_objects(analysis):
            if export_object_3d_view(obj, directory):
                exported_count += 1
    return exported_count


NO_VIEW_LABEL = "-- Vue courante --"
NO_SECTION_LABEL = "-- Aucune coupe --"


class SlideRowConfig(object):
    """
    Configuration d'affichage pour UNE ligne d'une liste de selection (un BC, un resultat, ...) : objet, vue/coupe a appliquer avant capture, steps eventuels.
    """

    def __init__(self, obj, analysis=None):
        """
        Fait : initialise la configuration d'une ligne (vue/coupe/steps/echelle/legende par defaut).
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.obj = obj
        # None si categorie hors multi-analyses (BC, Contacts...) ou projet mono-analyse : aucun suffixe affiche dans ce cas (voir analysis_suffix).
        self.analysis = analysis
        self.view_name = None
        self.section_name = None
        self.selected_steps = None       # None ou liste vide = pas de gestion par step
        self.step_display_mode = "individual"  # "individual" ou "combined"
        self.scale_factor = 1.0          # facteur d'echelle de deformation (mode "manual" uniquement, 1.0 = pas de mise a l'echelle)
        self.deformation_scale_mode = DEFAULT_DEFORMATION_SCALE_MODE  # "manual"/"auto_x1"/"auto_x2", voir DEFORMATION_SCALE_MODE_OPTIONS
        self.legend_name = None          # nom de legende (voir collect_legend_files), None = legende actuelle/automatique
        self.contour_view = DEFAULT_CONTOUR_VIEW            # mode d'affichage des couleurs (Isolines/SmoothContours/SolidFill/ContourBands)
        self.legend_orientation = DEFAULT_LEGEND_ORIENTATION  # orientation de la legende (Vertical/Horizontal)
        self.scoping_display = DEFAULT_SCOPING_DISPLAY       # affichage du scoping (ScopedBodies/ResultOnly/AllBodies)
        self.configured = False          # passe a True des que le bouton "..." a ete valide (OK)


def analysis_suffix(row_config):
    """
    Fait : construit le suffixe " (Nom Analyse)" a afficher pour differencier un meme resultat entre deux analyses.
    Depend de : row_config.analysis (voir collect_analyses et les collecteurs *_multi de ce fichier).
    Retourne : str, le suffixe formate, ou chaine vide si row_config.analysis est None.
    """
    if row_config.analysis is not None:
        return " ({})".format(row_config.analysis.Name)
    return ""


def build_row_display_name(row_config):
    """
    Fait : construit le texte affiche pour une ligne de selection dans la liste.
    Depend de : row_config (obj, view_name, section_name, selected_steps, deformation_scale_mode, scale_factor, legend_name, contour_view, legend_orientation, scoping_display), analysis_suffix.
    Retourne : str, le nom de l'objet suivi des reglages choisis separes par " | ".
    """
    parts = [row_config.obj.Name + analysis_suffix(row_config)]
    if row_config.view_name:
        parts.append("vue=" + row_config.view_name)
    if row_config.section_name:
        parts.append("coupe=" + row_config.section_name)
    if row_config.selected_steps:
        mode_label = "combine" if row_config.step_display_mode == "combined" else "individuel"
        steps_label = ",".join(str(step) for step in row_config.selected_steps)
        parts.append("steps={} ({})".format(steps_label, mode_label))
    if row_config.deformation_scale_mode == "auto_x1":
        parts.append("scale=Auto x1")
    elif row_config.deformation_scale_mode == "auto_x2":
        parts.append("scale=Auto x2")
    elif row_config.scale_factor and row_config.scale_factor != 1.0:
        parts.append("scale=x{}".format(row_config.scale_factor))
    if row_config.legend_name:
        parts.append("legende=" + row_config.legend_name)
    if row_config.contour_view and row_config.contour_view != DEFAULT_CONTOUR_VIEW:
        parts.append("affichage=" + contour_view_label(row_config.contour_view))
    if row_config.legend_orientation and row_config.legend_orientation != DEFAULT_LEGEND_ORIENTATION:
        parts.append("legende_orientation=" + legend_orientation_label(row_config.legend_orientation))
    if row_config.scoping_display and row_config.scoping_display != DEFAULT_SCOPING_DISPLAY:
        parts.append("scoping=" + scoping_display_label(row_config.scoping_display))
    return " | ".join(parts)


def collect_views():
    """
    Fait : liste les vues enregistrees dans le View Manager de Mechanical.
    Depend de : ExtAPI.Graphics.ModelViewManager.ExportModelViews, un export XML temporaire, xml.etree.ElementTree.
    Retourne : dict {nom (str): index (int)}, vide si aucune vue ou en cas d'erreur.
    """
    views = {}
    try:
        view_manager = ExtAPI.Graphics.ModelViewManager
        xml_path = os.path.join(CSV_EXPORT_FOLDER, "_model_views_tmp.xml")
        # Le View Manager ne s'inspecte pas directement via l'API scriptee : on passe par un export XML temporaire.
        view_manager.ExportModelViews(xml_path)
        tree = ET.parse(xml_path)
        for index, node in enumerate(list(tree.getroot())):
            if node.tag == "ModelView":
                views[node.attrib["Name"]] = index
    except Exception as e:
        print "Vues du View Manager indisponibles : " + str(e)
    return views


def collect_section_planes():
    """
    Fait : liste les plans de coupe (Section Planes) deja definis dans le modele.
    Depend de : ExtAPI.Graphics.SectionPlanes.
    Retourne : list, les objets Section Plane trouves (vide en cas d'erreur).
    """
    planes = []
    try:
        section_planes = ExtAPI.Graphics.SectionPlanes
        for i in range(section_planes.Count):
            planes.append(section_planes[i])
    except Exception as e:
        print "Plans de coupe indisponibles : " + str(e)
    return planes


def section_plane_label(section_plane, index):
    """
    Fait : construit un libelle lisible pour un plan de coupe.
    Depend de : section_plane.Name.
    Retourne : str, le nom du plan de coupe, ou un nom genere ("Section Plane N") s'il n'en a pas.
    """
    try:
        if section_plane.Name:
            return section_plane.Name
    except Exception:
        pass
    return "Section Plane {}".format(index + 1)


def apply_view_if_exists(view_name, views):
    """
    Fait : applique une vue du View Manager par son nom, si elle existe encore.
    Depend de : ExtAPI.Graphics.ModelViewManager.ApplyModelView, le dict views (voir collect_views).
    Retourne : rien (effet de bord : change la vue du viewport, ou ne fait rien si absente).
    """
    if not view_name or view_name not in views:
        return
    try:
        ExtAPI.Graphics.ModelViewManager.ApplyModelView(views[view_name])
    except Exception as e:
        print "Application de la vue '{}' impossible : {}".format(view_name, str(e))


def apply_section_plane(section_planes, section_labels, section_name):
    """
    Fait : active uniquement le plan de coupe designe par section_name, desactive les autres.
    Depend de : disable_all_section_planes, la correspondance d'index entre section_planes et section_labels.
    Retourne : rien (effet de bord : change l'etat Active des plans de coupe).
    """
    if not section_name:
        disable_all_section_planes(section_planes)
        return
    for i in range(len(section_planes)):
        try:
            section_planes[i].Active = (section_labels[i] == section_name)
        except Exception:
            pass


def disable_all_section_planes(section_planes):
    """
    Fait : desactive tous les plans de coupe fournis.
    Depend de : rien (parcourt la liste fournie).
    Retourne : rien (effet de bord : remet les plans de coupe a l'etat neutre avant/apres capture).
    """
    for section_plane in section_planes:
        try:
            section_plane.Active = False
        except Exception:
            pass


DEFORMATION_SCALE_MODE_OPTIONS = [
    ("Manuel (valeur ci-dessous)", "manual"),
    ("Auto Scale x1", "auto_x1"),
    ("Auto Scale x2", "auto_x2"),
]
DEFAULT_DEFORMATION_SCALE_MODE = "manual"


def deformation_scale_mode_label(value):
    """
    Fait : trouve le libelle affiche pour une valeur de deformation_scale_mode.
    Depend de : DEFORMATION_SCALE_MODE_OPTIONS.
    Retourne : str, le libelle correspondant (celui de DEFAULT_DEFORMATION_SCALE_MODE si value est inconnue).
    """
    for label, option_value in DEFORMATION_SCALE_MODE_OPTIONS:
        if option_value == value:
            return label
    return deformation_scale_mode_label(DEFAULT_DEFORMATION_SCALE_MODE)


def deformation_scale_mode_from_label(label):
    """
    Fait : trouve la valeur de deformation_scale_mode associee a un libelle de DEFORMATION_SCALE_MODE_OPTIONS.
    Depend de : DEFORMATION_SCALE_MODE_OPTIONS.
    Retourne : str, la valeur correspondante (DEFAULT_DEFORMATION_SCALE_MODE si le libelle est inconnu).
    """
    for option_label, value in DEFORMATION_SCALE_MODE_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_DEFORMATION_SCALE_MODE


def apply_scale_factor(deformation_scale_mode, scale_factor):
    """
    Fait : force l'echelle de deformation avant capture d'image - soit un facteur manuel
    (DeformationScaleMultiplier seul, comportement d'origine), soit un des deux presets natifs
    "Auto Scale" de Mechanical (DeformationScaling force sur Auto + multiplicateur fixe 1 ou 2).
    Depend de : ExtAPI.Graphics.ViewOptions.ResultPreference.DeformationScaling/DeformationScaleMultiplier,
        MechanicalEnums.Graphics.DeformationScaling (API Ansys, enum ambiant), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : change l'echelle affichee, ou ne fait rien en mode manuel avec scale_factor a 1.0).
    """
    try:
        vo = ExtAPI.Graphics.ViewOptions
        if deformation_scale_mode == "auto_x1":
            vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.Auto
            vo.ResultPreference.DeformationScaleMultiplier = 1
            ExtAPI.Graphics.Redraw()
        elif deformation_scale_mode == "auto_x2":
            vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.Auto
            vo.ResultPreference.DeformationScaleMultiplier = 2
            ExtAPI.Graphics.Redraw()
        elif scale_factor and scale_factor != 1.0:
            vo.ResultPreference.DeformationScaleMultiplier = float(scale_factor)
            ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Application du scale factor impossible : " + str(e)


def reset_scale_factor():
    """
    Fait : repasse l'echelle de deformation a l'etat neutre (mode Manuel, multiplicateur 1) apres
    une capture avec valeur personnalisee - y compris apres un preset "Auto Scale", pour ne pas
    laisser MechanicalEnums.Graphics.DeformationScaling sur Auto pour la capture suivante.
    Depend de : ExtAPI.Graphics.ViewOptions.ResultPreference.DeformationScaling/DeformationScaleMultiplier,
        MechanicalEnums.Graphics.DeformationScaling (API Ansys, enum ambiant), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : reinitialise l'echelle de deformation affichee).
    """
    try:
        vo = ExtAPI.Graphics.ViewOptions
        vo.ResultPreference.DeformationScaling = MechanicalEnums.Graphics.DeformationScaling.UserDefined
        vo.ResultPreference.DeformationScaleMultiplier = 1
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Reinitialisation du scale factor impossible : " + str(e)


CONTOUR_VIEW_OPTIONS = [
    ("ContourBands (default)", "ContourBands"),
    ("Isolines", "Isolines"),
    ("SmoothContours", "SmoothContours"),
    ("SolidFill", "SolidFill"),
]
DEFAULT_CONTOUR_VIEW = "ContourBands"


def contour_view_label(value):
    """
    Fait : trouve le libelle affiche pour une valeur de contour_view.
    Depend de : CONTOUR_VIEW_OPTIONS.
    Retourne : str, le libelle correspondant (celui de DEFAULT_CONTOUR_VIEW si value est inconnue).
    """
    for label, option_value in CONTOUR_VIEW_OPTIONS:
        if option_value == value:
            return label
    return contour_view_label(DEFAULT_CONTOUR_VIEW)


def contour_view_from_label(label):
    """
    Fait : trouve la valeur de contour_view associee a un libelle de CONTOUR_VIEW_OPTIONS.
    Depend de : CONTOUR_VIEW_OPTIONS.
    Retourne : str, la valeur correspondante (DEFAULT_CONTOUR_VIEW si le libelle est inconnu).
    """
    for option_label, value in CONTOUR_VIEW_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_CONTOUR_VIEW


def apply_contour_view(contour_view):
    """
    Fait : applique le mode d'affichage des couleurs de resultat (Isolignes/Contours lisses/Remplissage plein/Bandes de contour).
    Depend de : ExtAPI.Graphics.ViewOptions.ResultPreference.ContourView, ExtAPI.Graphics.Redraw (API Ansys).
    Retourne : rien (effet de bord sur le viewport, ou ne fait rien si contour_view est vide).
    """
    if not contour_view:
        return
    try:
        vo = ExtAPI.Graphics.ViewOptions
        # Les membres (Isolines/SmoothContours/SolidFill/ContourBands) sont lus depuis le TYPE de
        # la valeur courante plutot qu'importes explicitement : c'est un enum .NET ambiant, deja
        # utilise ainsi ailleurs dans le projet (ex: ModelColoring.ByMaterial).
        vo.ResultPreference.ContourView = getattr(vo.ResultPreference.ContourView, contour_view)
        # Redraw() indispensable : changer cette propriete par script ne rafraichit pas seul le
        # viewport (meme constat que pour la legende, voir reset_legend) - sans cet appel, l'image
        # exportee juste apres reste sur l'ancien mode d'affichage.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Application du mode d'affichage '{}' impossible : {}".format(contour_view, str(e))


def reset_contour_view():
    """
    Fait : repasse le mode d'affichage des couleurs de resultat a Bandes de contour (etat neutre) apres une capture personnalisee.
    Depend de : apply_contour_view, DEFAULT_CONTOUR_VIEW.
    Retourne : rien (effet de bord sur le viewport).
    """
    apply_contour_view(DEFAULT_CONTOUR_VIEW)


LEGEND_ORIENTATION_OPTIONS = [
    ("Verticale (defaut)", "Vertical"),
    ("Horizontale", "Horizontal"),
]
DEFAULT_LEGEND_ORIENTATION = "Vertical"


def legend_orientation_label(value):
    """
    Fait : trouve le libelle affiche pour une valeur de legend_orientation.
    Depend de : LEGEND_ORIENTATION_OPTIONS.
    Retourne : str, le libelle correspondant (celui de DEFAULT_LEGEND_ORIENTATION si value est inconnue).
    """
    for label, option_value in LEGEND_ORIENTATION_OPTIONS:
        if option_value == value:
            return label
    return legend_orientation_label(DEFAULT_LEGEND_ORIENTATION)


def legend_orientation_from_label(label):
    """
    Fait : trouve la valeur de legend_orientation associee a un libelle de LEGEND_ORIENTATION_OPTIONS.
    Depend de : LEGEND_ORIENTATION_OPTIONS.
    Retourne : str, la valeur correspondante (DEFAULT_LEGEND_ORIENTATION si le libelle est inconnu).
    """
    for option_label, value in LEGEND_ORIENTATION_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_LEGEND_ORIENTATION


def apply_legend_orientation(legend_orientation):
    """
    Fait : applique l'orientation de la legende du viewport (verticale/horizontale).
    Depend de : ExtAPI.Graphics.GlobalLegendSettings.LegendOrientation, LegendOrientationType (API Ansys, enum ambiant), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord sur le viewport, ou ne fait rien si legend_orientation est vide).
    """
    if not legend_orientation:
        return
    try:
        ExtAPI.Graphics.GlobalLegendSettings.LegendOrientation = getattr(LegendOrientationType, legend_orientation)
        # Redraw() indispensable : changer cette propriete par script ne rafraichit pas seul le
        # viewport (meme constat que pour la legende, voir reset_legend) - sans cet appel, l'image
        # exportee juste apres reste sur l'ancienne orientation.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Application de l'orientation de legende '{}' impossible : {}".format(legend_orientation, str(e))


def reset_legend_orientation():
    """
    Fait : repasse l'orientation de la legende a Verticale (etat neutre) apres une capture personnalisee.
    Depend de : apply_legend_orientation, DEFAULT_LEGEND_ORIENTATION.
    Retourne : rien (effet de bord sur le viewport).
    """
    apply_legend_orientation(DEFAULT_LEGEND_ORIENTATION)


SCOPING_DISPLAY_OPTIONS = [
    ("ScopedBodies (default)", "ScopedBodies"),
    ("ResultOnly", "ResultOnly"),
    ("AllBodies", "AllBodies"),
]
DEFAULT_SCOPING_DISPLAY = "ScopedBodies"


def scoping_display_label(value):
    """
    Fait : trouve le libelle affiche pour une valeur de scoping_display.
    Depend de : SCOPING_DISPLAY_OPTIONS.
    Retourne : str, le libelle correspondant (celui de DEFAULT_SCOPING_DISPLAY si value est inconnue).
    """
    for label, option_value in SCOPING_DISPLAY_OPTIONS:
        if option_value == value:
            return label
    return scoping_display_label(DEFAULT_SCOPING_DISPLAY)


def scoping_display_from_label(label):
    """
    Fait : trouve la valeur de scoping_display associee a un libelle de SCOPING_DISPLAY_OPTIONS.
    Depend de : SCOPING_DISPLAY_OPTIONS.
    Retourne : str, la valeur correspondante (DEFAULT_SCOPING_DISPLAY si le libelle est inconnu).
    """
    for option_label, value in SCOPING_DISPLAY_OPTIONS:
        if option_label == label:
            return value
    return DEFAULT_SCOPING_DISPLAY


def apply_scoping_display(scoping_display):
    """
    Fait : applique le mode d'affichage du scoping (corps scopes / resultat seul / tous les corps) avant capture.
    Depend de : ExtAPI.Graphics.ViewOptions.ResultPreference.ScopingDisplay, MechanicalEnums.Graphics.ScopingDisplay (API Ansys, enum ambiant), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord sur le viewport, ou ne fait rien si scoping_display est vide).
    """
    if not scoping_display:
        return
    try:
        vo = ExtAPI.Graphics.ViewOptions
        vo.ResultPreference.ScopingDisplay = getattr(MechanicalEnums.Graphics.ScopingDisplay, scoping_display)
        # Redraw() indispensable : changer cette propriete par script ne rafraichit pas seul le
        # viewport (meme constat que pour ContourView/LegendOrientation) - sans cet appel, l'image
        # exportee juste apres reste sur l'ancien mode d'affichage.
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Application du mode de scoping '{}' impossible : {}".format(scoping_display, str(e))


def reset_scoping_display():
    """
    Fait : repasse le mode d'affichage du scoping a Corps scopes (etat neutre) apres une capture personnalisee.
    Depend de : apply_scoping_display, DEFAULT_SCOPING_DISPLAY.
    Retourne : rien (effet de bord sur le viewport).
    """
    apply_scoping_display(DEFAULT_SCOPING_DISPLAY)


NO_LEGEND_LABEL = "-- Legende automatique --"


def collect_legend_files():
    """
    Fait : liste les legendes disponibles dans LEGEND_FOLDER (fichiers .xml).
    Depend de : os.path.isdir/os.listdir, LEGEND_FOLDER (00_constants.py).
    Retourne : list, noms de legende sans extension (str), tries alphabetiquement, vide si dossier absent.
    """
    if not os.path.isdir(LEGEND_FOLDER):
        return []
    return sorted(f[:-4] for f in os.listdir(LEGEND_FOLDER) if f.lower().endswith(".xml"))


def get_result_display_unit(result_obj, force_evaluate=True):
    """
    Fait : extrait l'unite du resultat affiche (ex: "MPa") depuis le texte de sa propriete Minimum/Maximum/Average.
    Depend de : result_obj.EvaluateAllResults, result_obj.VisibleProperties (le panneau Details).
    Retourne : str, l'unite detectee (ex: "MPa"), ou None si indisponible.
    """
    # result_obj.Maximum.Unit s'est avere peu fiable (unite absente/incorrecte selon le type de resultat) ; le texte de VisibleProperties contient toujours l'unite reellement utilisee.
    if force_evaluate:
        try:
            result_obj.EvaluateAllResults()
        except Exception:
            pass

    # "Minimum Occurs On"/"Maximum Occurs On" sont exclus expres : leur StringValue est un nom de corps, pas une valeur chiffree suivie d'une unite.
    candidate_captions = ("Minimum", "Maximum", "Average", "Minimum Value", "Maximum Value", "Average Value")

    # Log reserve aux appels reels (force_evaluate=True) : sinon la simple ouverture de la fenetre "..." inonderait la console sans rien generer.
    try:
        for prop in result_obj.VisibleProperties:
            try:
                caption = prop.Caption
                if caption not in candidate_captions:
                    continue
                tokens = prop.StringValue.split()
                if len(tokens) >= 2:
                    if force_evaluate:
                        print "Unite detectee pour {} via '{}' ({}) : {}".format(
                            result_obj.Name, caption, prop.StringValue, tokens[-1])
                    return tokens[-1]
            except Exception:
                pass
    except Exception as e:
        if force_evaluate:
            print "Unite indisponible pour {} : {}".format(result_obj.Name, str(e))
        return None

    if force_evaluate:
        print "Aucune unite detectee pour {} (aucune propriete Minimum/Maximum/Average exploitable).".format(result_obj.Name)
    return None


def apply_legend_if_exists(legend_name, result_obj):
    """
    Fait : importe une legende par son nom dans l'unite du resultat concerne et l'applique au viewport.
    Depend de : LEGEND_FOLDER, get_result_display_unit, ExtAPI.Graphics.ImportLegend, CurrentLegendSettings.
    Retourne : rien (effet de bord : change la legende du viewport, ou ne fait rien si legend_name est None).
    """
    if not legend_name:
        return
    xml_path = os.path.join(LEGEND_FOLDER, legend_name + ".xml")
    if not os.path.exists(xml_path):
        print "Legende introuvable : " + xml_path
        return

    # ImportLegend/Reset comparent l'unite demandee a celle de l'objet ACTUELLEMENT ACTIF dans le viewport, pas a result_obj : sans cet Activate() explicite, l'unite comparee restait celle de la ligne precedente.
    try:
        result_obj.Activate()
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Activation de {} impossible avant application de la legende : {}".format(result_obj.Name, str(e))

    # Tentative systematique meme si unit est None : laisse remonter l'erreur .NET reelle en console plutot que d'abandonner silencieusement.
    unit = get_result_display_unit(result_obj)
    print "Legende '{}' sur {} : unite utilisee pour ImportLegend = {}".format(legend_name, result_obj.Name, unit)

    reset_legend()

    try:
        legend = ExtAPI.Graphics.ImportLegend(xml_path, unit)
        legend.CopyTo(Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings())
        print "Legende '{}' appliquee sur {} (unite={}).".format(legend_name, result_obj.Name, unit)
    except Exception as e:
        print "Application de la legende '{}' impossible sur {} (unite={}) : {}".format(
            legend_name, result_obj.Name, unit, str(e))


def reset_legend():
    """
    Fait : reinitialise la legende courante du viewport a son etat automatique.
    Depend de : Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings, ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : remet la legende affichee a l'automatique).
    """
    # Redraw() indispensable : changer cette propriete par script ne rafraichit pas seul le viewport tant qu'aucun autre evenement ne force un redessin.
    try:
        Ansys.Mechanical.Graphics.Tools.CurrentLegendSettings().Reset()
        ExtAPI.Graphics.Redraw()
    except Exception as e:
        print "Reinitialisation de la legende impossible : " + str(e)


# Templates disponibles : 2, 3, 4, 6 et 8 steps ; 5, 7 steps (et au-dela de 8) retombent automatiquement sur le mode "slides individuelles".
MULTI_STEP_SLIDE_TEMPLATES = {
    2: {
        "layout_index": 3,
        "image_shape_indices": [3, 2],
    },
    3: {
        "layout_index": 4,
        "image_shape_indices": [3, 2, 8],
    },
    4: {
        "layout_index": 5,
        "image_shape_indices": [3, 2, 8, 9],
    },
    6: {
        "layout_index": 6,
        "image_shape_indices": [3, 2, 8, 9, 10, 11],
    },
    8: {
        "layout_index": 7,
        "image_shape_indices": [3, 2, 8, 9, 10, 11, 12, 13],
    },
}


def get_multi_step_template(step_count):
    """
    Fait : recupere le template de slide combinee correspondant a ce nombre exact de steps.
    Depend de : MULTI_STEP_SLIDE_TEMPLATES.
    Retourne : dict (layout_index/image_shape_indices), ou None si aucun template pour ce nombre de steps.
    """
    return MULTI_STEP_SLIDE_TEMPLATES.get(step_count)


def get_step_count(analysis):
    """
    Fait : lit le nombre de steps definis au niveau de l'analyse.
    Depend de : analysis.AnalysisSettings.NumberOfSteps.
    Retourne : int, le nombre de steps, ou 1 en cas d'erreur.
    """
    try:
        return int(analysis.AnalysisSettings.NumberOfSteps)
    except Exception as e:
        print "Nombre de steps indisponible : " + str(e)
        return 1


def _set_result_display_time(result_obj, display_time):
    """
    Fait : repositionne un resultat sur un DisplayTime precis et le reevalue.
    Depend de : result_obj.DisplayTime/EvaluateAllResults/Evaluate, ExtAPI.Graphics.Redraw, SWF.Application.DoEvents.
    Retourne : rien (effet de bord : restaure l'affichage d'origine du resultat).
    """
    # Uniquement pour restaurer l'etat d'origine apres coup : les captures par step passent par evaluate_result_for_step (SetNumber), pas par cette fonction.
    result_obj.DisplayTime = display_time
    for method_name in ("EvaluateAllResults", "Evaluate"):
        method = getattr(result_obj, method_name, None)
        if method is None:
            continue
        try:
            method()
            break
        except Exception as e:
            print "Evaluation ({}) impossible pour {} : {}".format(method_name, result_obj.Name, str(e))
    try:
        ExtAPI.Graphics.Redraw()
        SWF.Application.DoEvents()
    except Exception as e:
        print "Redraw impossible apres reevaluation de {} : {}".format(result_obj.Name, str(e))


def evaluate_result_for_step(result_obj, step_number):
    """
    Fait : positionne un resultat sur un set/step precis via SetNumber plutot que DisplayTime.
    Depend de : result_obj.Activate/By/SetNumber/EvaluateAllResults, SetDriverStyle.ResultSet, ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : le resultat affiche desormais ce step).
    """
    # SetNumber navigue vers un set deja calcule par le solveur sans reevaluation complete : plus fiable pour enchainer plusieurs steps que l'ancienne approche par DisplayTime.
    result_obj.Activate()
    result_obj.By = SetDriverStyle.ResultSet
    result_obj.SetNumber = step_number
    result_obj.EvaluateAllResults()
    ExtAPI.Graphics.Redraw()


def export_result_image_for_step(result_obj, step_number):
    """
    Fait : exporte l'image d'un resultat pour un set/step precis.
    Depend de : evaluate_result_for_step, export_current_view_image (02_image_export.py).
    Retourne : str, le chemin du PNG genere.
    """
    # Export direct de la vue apres Activate(), sans Figure snapshot (contrairement a export_object_image).
    evaluate_result_for_step(result_obj, step_number)
    return export_current_view_image("{}_step{}".format(result_obj.Name, step_number))


def add_multi_step_image_slide(presentation, template, title, image_paths):
    """
    Fait : ajoute une slide combinee (plusieurs images de steps) a partir d'un template de MULTI_STEP_SLIDE_TEMPLATES.
    Depend de : presentation.SlideMaster.CustomLayouts, presentation.Slides.AddSlide.
    Retourne : PPT.Slide, la slide creee.
    """
    layout = presentation.SlideMaster.CustomLayouts[template["layout_index"]]
    slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout)
    slide.Shapes[1].TextFrame.TextRange.Text = title

    # Tri par position reelle (haut->bas, gauche->droite), pas par indice de shape, pour respecter l'ordre chronologique des steps.
    placeholders = [slide.Shapes[idx] for idx in template["image_shape_indices"]]
    placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    for i in range(min(len(image_paths), len(placeholders))):
        ph = placeholders[i]
        slide.Shapes.AddPicture(image_paths[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                 ph.Left, ph.Top, ph.Width, ph.Height)
    return slide


def capture_multi_result_cell_image(cfg, views, section_planes, section_labels):
    """
    Fait : applique la configuration graphique d'UNE case (slide combinee multi-resultats : vue,
    coupe, legende, apparence, scale factor) et exporte une image unique du resultat choisi - pas de
    notion de step ici (contrairement a build_single_result_slide/build_step_based_result_slides),
    chaque case porte un resultat different affiche dans son etat courant.
    Depend de : apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_contour_view,
    apply_legend_orientation, apply_scoping_display, apply_legend_if_exists, export_solution_image
    (02_image_export.py), disable_all_section_planes, reset_scale_factor/reset_contour_view/
    reset_legend_orientation/reset_scoping_display.
    Retourne : str, le chemin de l'image exportee, ou None en cas d'erreur.
    """
    obj = cfg.obj
    image_path = None
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, obj)
        image_path = export_solution_image(obj)
    except Exception as e:
        print "Capture impossible pour {} : {}".format(obj.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
    return image_path


def build_multi_result_slide(report, template, cell_configs, views, section_planes, section_labels):
    """
    Fait : construit UNE slide combinee "differents resultats" (un resultat different par case
    configuree, chacun avec sa propre vue/coupe/legende/apparence), a partir d'un template de
    MULTI_STEP_SLIDE_TEMPLATES (meme famille de templates que les slides combinees multi-step, mais
    ici chaque emplacement recoit un resultat different plutot qu'un meme resultat a un step different).
    Depend de : capture_multi_result_cell_image, add_multi_step_image_slide.
    Retourne : rien (effet de bord : ajoute une slide a report.presentation, ou ne fait rien si aucune case n'est configuree).
    """
    image_paths = []
    titles = []
    for cfg in cell_configs:
        if cfg is None:
            continue
        image_path = capture_multi_result_cell_image(cfg, views, section_planes, section_labels)
        if image_path:
            image_paths.append(image_path)
            titles.append(cfg.obj.Name)

    if not image_paths:
        print "Aucune case configuree : slide combinee multi-resultats non generee."
        return

    title = "Resultats combines : " + ", ".join(titles)
    add_multi_step_image_slide(report.presentation, template, title, image_paths)
    print "Slide combinee multi-resultats ajoutee ({} resultats).".format(len(image_paths))


def build_step_based_result_slides(report, cfg, obj, subtitle, analysis):
    """
    Fait : construit la ou les slides d'un resultat avec une selection de steps (une combinee si possible, sinon une par step).
    Depend de : get_multi_step_template, export_result_image_for_step, export_result_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une ou plusieurs slides au rapport).
    """
    # analysis n'est plus utilise (captures par SetNumber, pas par un temps calcule depuis l'analyse) ; garde dans la signature pour ne pas casser les appelants.
    steps = cfg.selected_steps
    template = get_multi_step_template(len(steps)) if cfg.step_display_mode == "combined" else None
    original_display_time = obj.DisplayTime
    original_by = obj.By

    display_name = obj.Name + analysis_suffix(cfg)

    try:
        if template:
            image_paths = [export_result_image_for_step(obj, step) for step in steps]
            title = "{} - {} steps".format(display_name, len(steps))
            add_multi_step_image_slide(report.presentation, template, title, image_paths)
            print "Slide combinee ({} steps) ajoutee pour {}.".format(len(steps), display_name)
            return

        csv_path = None
        try:
            csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, obj)
        except Exception as e:
            print "Export CSV impossible pour {} : {}".format(obj.Name, str(e))

        for step in steps:
            img_path = None
            try:
                img_path = export_result_image_for_step(obj, step)
            except Exception as e:
                print "Export image impossible pour {} (step {}) : {}".format(obj.Name, step, str(e))
            title = "{} - Step {}".format(display_name, step)
            report.add_image_table_slide(title, subtitle, img_path=img_path, csv_path=csv_path)
    finally:
        # Restauration obligatoire meme en cas d'erreur : sinon l'objet reste fige sur le dernier step traite et deregle la legende des slides suivantes.
        obj.By = original_by
        _set_result_display_time(obj, original_display_time)


def flatten_results(objects):
    """
    Fait : deplie recursivement les dossiers de regroupement (ex: "Group Similar Children") en leurs objets terminaux.
    Depend de : rien (parcourt obj.Children recursivement).
    Retourne : list, les objets feuilles exportables en image/CSV.
    """
    leaves = []
    for obj in objects:
        try:
            children = obj.Children
        except Exception:
            children = None
        if children is not None and len(children) > 0:
            leaves.extend(flatten_results(list(children)))
        else:
            leaves.append(obj)
    return leaves


def collect_boundary_conditions(analysis=None):
    """
    Fait : liste les Boundary Conditions du modele, limitees a une analyse si fournie.
    Depend de : ExtAPI.DataModel.GetObjectsByType(GenericBoundaryCondition), _is_descendant_of.
    Retourne : list, les objets Boundary Condition (tout le projet si analysis est None).
    """
    all_bcs = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.GenericBoundaryCondition))
    if analysis is None:
        return all_bcs
    return [bc for bc in all_bcs if _is_descendant_of(bc, analysis)]


def collect_boundary_conditions_multi(analyses):
    """
    Fait : liste les Boundary Conditions de toutes les analyses fournies.
    Depend de : collect_boundary_conditions.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for bc in collect_boundary_conditions(analysis):
            pairs.append((bc, analysis))
    return pairs


def collect_analyses():
    """
    Fait : liste les analyses du projet exploitables par le generateur (Analysis Settings valides).
    Depend de : ExtAPI.DataModel.AnalysisList, analysis.AnalysisSettings.NumberOfSteps.
    Retourne : list, les objets analyse du projet - exclut les addins de post-traitement (ex : FEMFAT)
    dont AnalysisSettings est None : ils n'ont ni steps ni resultats de solution classiques, et
    faisaient planter toute la generation (settings.NumberOfSteps sur un objet None) des qu'ils
    etaient selectionnes dans une liste de la GUI.
    """
    analyses = []
    for analysis in ExtAPI.DataModel.AnalysisList:
        try:
            analysis.AnalysisSettings.NumberOfSteps
        except Exception:
            print "Analyse ignoree (Analysis Settings indisponible, ex: addin FEMFAT) : " + str(analysis.Name)
            continue
        analyses.append(analysis)
    return analyses


def collect_bolt_pretensions(analysis=None):
    """
    Fait : liste les Bolt Pretension du modele, limitees a une analyse si fournie.
    Depend de : ExtAPI.DataModel.GetObjectsByType(BoltPretension), _is_descendant_of.
    Retourne : list, les objets Bolt Pretension (tout le projet si analysis est None).
    """
    all_bolt_pretensions = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltPretension))
    if analysis is None:
        return all_bolt_pretensions
    return [bp for bp in all_bolt_pretensions if _is_descendant_of(bp, analysis)]


def collect_bolt_pretensions_multi(analyses):
    """
    Fait : liste les Bolt Pretension de toutes les analyses fournies.
    Depend de : collect_bolt_pretensions.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for bp in collect_bolt_pretensions(analysis):
            pairs.append((bp, analysis))
    return pairs


def _is_descendant_of(obj, ancestor):
    """
    Fait : verifie si obj est un descendant (direct ou indirect) de ancestor dans l'arbre Mechanical.
    Depend de : obj.Parent (remontee de l'arbre).
    Retourne : bool, True si ancestor est bien un parent de obj.
    """
    # Distingue un Contact Tool de la branche Connections (sans step) de son homonyme dans Solution (avec step) : meme categorie .NET, seule la position dans l'arbre differe.
    node = getattr(obj, "Parent", None)
    while node is not None:
        if node == ancestor:
            return True
        node = getattr(node, "Parent", None)
    return False


def collect_contact_tool_results(analysis):
    """
    Fait : liste les resultats des dossiers Contact Tool de la branche Solution (avec steps) pour une analyse.
    Depend de : ExtAPI.DataModel.GetObjectsByType(ContactTool), _is_descendant_of, flatten_results.
    Retourne : list, les objets resultat exportables propres a la branche Solution.
    """
    # ContactTool existe aussi dans Connections (sans step, memes noms d'enfants) : le filtre par branche evite de melanger les deux listes.
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool)
    children = []
    for tool in tools:
        if tool.Children and _is_descendant_of(tool, analysis.Solution):
            children.extend(list(tool.Children))
    return flatten_results(children)


def collect_contact_tool_results_multi(analyses):
    """
    Fait : liste les resultats Contact Tool (branche Solution) de toutes les analyses fournies.
    Depend de : collect_contact_tool_results.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_contact_tool_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def collect_connection_contact_tool_results(analysis):
    """
    Fait : liste les resultats des dossiers Contact Tool de la branche Connections (sans step) pour une analyse.
    Depend de : ExtAPI.DataModel.GetObjectsByType(ContactTool), _is_descendant_of, flatten_results.
    Retourne : list, les objets resultat exportables propres a la branche Connections.
    """
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool)
    children = []
    for tool in tools:
        if tool.Children and not _is_descendant_of(tool, analysis.Solution):
            children.extend(list(tool.Children))
    return flatten_results(children)


def collect_bolt_tool_results(analysis=None):
    """
    Fait : liste les resultats des dossiers Bolt Tool sous Solution, limites a une analyse si fournie.
    Depend de : ExtAPI.DataModel.GetObjectsByType(BoltTool), _is_descendant_of, flatten_results.
    Retourne : list, les objets resultat exportables (tout le projet si analysis est None).
    """
    tools = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool)
    children = []
    for tool in tools:
        if not tool.Children:
            continue
        if analysis is not None and not _is_descendant_of(tool, analysis.Solution):
            continue
        children.extend(list(tool.Children))
    return flatten_results(children)


def collect_bolt_tool_results_multi(analyses):
    """
    Fait : liste les resultats Bolt Tool de toutes les analyses fournies.
    Depend de : collect_bolt_tool_results.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_bolt_tool_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def collect_all_results(analysis):
    """
    Fait : liste les resultats "simples" de Solution (Deformation, Contrainte, Probe...), hors Solution Information/Contact Tool/Bolt Tool.
    Depend de : analysis.Solution.Children, ExtAPI.DataModel.GetObjectsByType(ContactTool/BoltTool), flatten_results.
    Retourne : list, les objets resultat exportables.
    """
    excluded_categories = [DataModelObjectCategory.ContactTool, DataModelObjectCategory.BoltTool]

    # Exclusion par identite en plus de la categorie : DataModelObjectCategory peut echouer silencieusement (category=None) sur le dossier Contact/Bolt Tool lui-meme, ce qui le ferait passer le filtre et dupliquerait ses enfants avec la liste separee dediee.
    already_handled = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool))
    already_handled += list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool))

    solution_children = analysis.Solution.Children

    candidates = []
    if solution_children:
        for i in range(1, len(solution_children)):  # index 0 = Solution Information
            child = solution_children[i]
            if child in already_handled:
                continue
            try:
                category = child.DataModelObjectCategory
            except Exception:
                category = None
            if category in excluded_categories:
                continue
            candidates.append(child)

    return flatten_results(candidates)


def collect_all_results_multi(analyses):
    """
    Fait : liste les resultats simples de toutes les analyses fournies.
    Depend de : collect_all_results.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for obj in collect_all_results(analysis):
            pairs.append((obj, analysis))
    return pairs


def build_single_bc_slide(report, cfg, views, section_planes, section_labels):
    """
    Fait : construit la slide d'UNE Boundary Condition avec sa vue/coupe/legende/apparence configuree.
    Depend de : apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, export_object_image, export_bc_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport).
    """
    bc = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, bc)
        img_path = export_object_image(bc, bc.Name)
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        csv_path = export_bc_tabular_data(CSV_EXPORT_FOLDER, bc)
        report.add_image_table_slide(bc.Name + analysis_suffix(cfg), "-- Boundary Conditions --",
                                      img_path=img_path, csv_path=csv_path)
    except Exception as e:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        print "Slide BC impossible pour {} : {}".format(bc.Name, str(e))


def build_bc_slides(report, row_configs, views, section_planes, section_labels):
    """
    Fait : ajoute une slide pour chaque Boundary Condition selectionnee.
    Depend de : build_single_bc_slide.
    Retourne : rien (effet de bord : ajoute une slide par ligne au rapport).
    """
    for cfg in row_configs:
        build_single_bc_slide(report, cfg, views, section_planes, section_labels)


def build_single_bp_slide(report, cfg, views, section_planes, section_labels):
    """
    Fait : construit la slide d'UNE Bolt Pretension avec sa vue/coupe/legende/apparence configuree.
    Depend de : apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, export_object_image, export_bp_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport).
    """
    bp = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, bp)
        img_path = export_object_image(bp, bp.Name)
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        csv_path = export_bp_tabular_data(CSV_EXPORT_FOLDER, bp)
        report.add_image_table_slide(bp.Name + analysis_suffix(cfg), "-- Bolt Pretension --",
                                      img_path=img_path, csv_path=csv_path)
    except Exception as e:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()
        print "Slide Bolt Pretension impossible pour {} : {}".format(bp.Name, str(e))


def build_bp_slides(report, row_configs, views, section_planes, section_labels):
    """
    Fait : ajoute une slide pour chaque Bolt Pretension selectionnee.
    Depend de : build_single_bp_slide.
    Retourne : rien (effet de bord : ajoute une slide par ligne au rapport).
    """
    for cfg in row_configs:
        build_single_bp_slide(report, cfg, views, section_planes, section_labels)


def build_single_result_slide(report, cfg, subtitle, views, section_planes, section_labels, analysis):
    """
    Fait : construit la slide d'UN objet resultat, avec sa vue/coupe/legende/apparence et sa selection de steps eventuelle.
    Depend de : apply_view_if_exists, apply_section_plane, apply_scale_factor, apply_legend_if_exists, apply_contour_view, apply_legend_orientation, apply_scoping_display, build_step_based_result_slides, export_object_image, export_result_tabular_data.
    Retourne : rien (effet de bord : ajoute une ou plusieurs slides au rapport).
    """
    obj = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        apply_scale_factor(cfg.deformation_scale_mode, cfg.scale_factor)
        apply_contour_view(cfg.contour_view)
        apply_legend_orientation(cfg.legend_orientation)
        apply_scoping_display(cfg.scoping_display)
        apply_legend_if_exists(cfg.legend_name, obj)

        if cfg.selected_steps:
            build_step_based_result_slides(report, cfg, obj, subtitle, analysis)
        else:
            img_path = None
            try:
                img_path = export_object_image(obj, obj.Name)
            except Exception as e:
                print "Export image impossible pour {} : {}".format(obj.Name, str(e))

            csv_path = None
            try:
                csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, obj)
            except Exception as e:
                print "Export CSV impossible pour {} : {}".format(obj.Name, str(e))

            if img_path or csv_path:
                report.add_image_table_slide(obj.Name + analysis_suffix(cfg), subtitle,
                                              img_path=img_path, csv_path=csv_path)
            else:
                print "Aucune donnee exportable pour " + obj.Name + " : slide ignoree."
    except Exception as e:
        print "Slide resultat impossible pour {} : {}".format(obj.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)
        reset_scale_factor()
        reset_contour_view()
        reset_legend_orientation()
        reset_scoping_display()


def build_result_slides(report, row_configs, subtitle, views, section_planes, section_labels, analysis):
    """
    Fait : ajoute une slide pour chaque objet resultat selectionne (Contact Tool, Bolt Tool ou resultats generaux).
    Depend de : build_single_result_slide.
    Retourne : rien (effet de bord : ajoute une slide par ligne au rapport).
    """
    for cfg in row_configs:
        build_single_result_slide(report, cfg, subtitle, views, section_planes, section_labels, analysis)


DEFAULT_CONTEXT_OPACITY_PERCENT = 50  


class GeometryPartRowConfig(object):
    """
    Configuration d'affichage pour une slide "geometrie simple" (une piece isolee, opaque, dans le contexte transparent de l'assemblage).
    """

    def __init__(self, body):
        """
        Fait : initialise la configuration d'une piece isolee (geometrie simple) avec ses valeurs par defaut.
        Depend de : DEFAULT_CONTEXT_OPACITY_PERCENT.
        Retourne : rien (constructeur).
        """
        self.obj = body
        self.view_name = None
        self.section_name = None
        self.context_opacity_percent = DEFAULT_CONTEXT_OPACITY_PERCENT
        self.configured = False


def build_geometry_row_display_name(row_config):
    """
    Fait : construit le texte affiche dans la liste pour une piece (geometrie simple).
    Depend de : rien (lit row_config.obj/view_name/section_name/context_opacity_percent).
    Retourne : str, le nom de la piece suivi des reglages choisis separes par " | ".
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("vue=" + row_config.view_name)
    if row_config.section_name:
        parts.append("coupe=" + row_config.section_name)
    parts.append("contexte={}%".format(row_config.context_opacity_percent))
    return " | ".join(parts)


def collect_bodies():
    """
    Fait : liste tous les corps (Body) du modele.
    Depend de : ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body).
    Retourne : list, les objets Body du modele.
    """
    return list(ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body, True))


def isolate_body_by_transparency(target_body, all_bodies, context_opacity_percent):
    """
    Fait : rend une piece totalement opaque et les autres semi-transparentes au pourcentage donne.
    Depend de : Body.Transparency, Transaction (Ansys.ACT.Mechanical), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : change la transparence de tous les corps du modele).
    """
    # Body.Transparency va de 0.0 (transparent) a 1.0 (opaque), malgre son nom qui suggere l'inverse.
    context_value = max(0.0, min(1.0, context_opacity_percent / 100.0))
    target_id = target_body.ObjectId
    # Transaction(True) differe le rafraichissement pendant la boucle ; le Redraw() explicite ensuite reste necessaire pour que le changement soit visible avant la capture d'image.
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Transparency = 1.0 if body.ObjectId == target_id else context_value
            except Exception:
                pass
    try:
        ExtAPI.Graphics.Redraw()
    except Exception:
        pass


def reset_all_bodies_transparency(all_bodies):
    """
    Fait : remet toutes les pieces a l'opacite normale (100%).
    Depend de : Body.Transparency, Transaction (Ansys.ACT.Mechanical), ExtAPI.Graphics.Redraw.
    Retourne : rien (effet de bord : restaure l'opacite de tous les corps du modele).
    """
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Transparency = 1.0
            except Exception:
                pass
    try:
        ExtAPI.Graphics.Redraw()
    except Exception:
        pass


def export_geometry_part_image(body, all_bodies, context_opacity_percent):
    """
    Fait : isole une piece (opaque) dans le contexte transparent de l'assemblage, puis exporte son image.
    Depend de : isolate_body_by_transparency, geometry.AddFigure, export_current_view_image, reset_all_bodies_transparency.
    Retourne : str, le chemin du PNG genere.
    """
    isolate_body_by_transparency(body, all_bodies, context_opacity_percent)
    geometry = ExtAPI.DataModel.Project.Model.Geometry
    figure = geometry.AddFigure()
    figure.Activate()
    # Pas de SetFit() : le cadrage de la camera (vue choisie via apply_view_if_exists, ou position
    # manuelle courante) est laisse tel quel, a la responsabilite de l'utilisateur.
    image_path = export_current_view_image("Geometry_" + safe_file_name(body.Name))
    reset_all_bodies_transparency(all_bodies)
    return image_path


def build_single_geometry_part_slide(report, cfg, all_bodies, views, section_planes, section_labels):
    """
    Fait : construit la slide "geometrie simple" (une image, pas de tableau) d'UNE piece isolee.
    Depend de : apply_view_if_exists, apply_section_plane, export_geometry_part_image, report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport).
    """
    body = cfg.obj
    try:
        apply_view_if_exists(cfg.view_name, views)
        apply_section_plane(section_planes, section_labels, cfg.section_name)
        img_path = export_geometry_part_image(body, all_bodies, cfg.context_opacity_percent)
        report.add_image_table_slide(body.Name, "-- Geometry --", img_path=img_path, csv_path=None)
    except Exception as e:
        print "Slide geometrie impossible pour {} : {}".format(body.Name, str(e))
    finally:
        disable_all_section_planes(section_planes)


def build_geometry_part_slides(report, row_configs, all_bodies, views, section_planes, section_labels):
    """
    Fait : ajoute une slide "geometrie simple" pour chaque piece selectionnee.
    Depend de : build_single_geometry_part_slide.
    Retourne : rien (effet de bord : ajoute une slide par piece au rapport).
    """
    for cfg in row_configs:
        build_single_geometry_part_slide(report, cfg, all_bodies, views, section_planes, section_labels)


class MeshPartRowConfig(object):
    """
    Ligne de selection pour le mesh par piece isolee : le corps et une vue eventuelle (pas de coupe/opacite, isolation par masquage total).
    """

    def __init__(self, body):
        """
        Fait : initialise la configuration d'une piece isolee (mesh) avec ses valeurs par defaut.
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.obj = body
        self.view_name = None
        self.configured = False  # passe a True des que le bouton "..." a ete valide (OK)


def build_mesh_part_row_display_name(row_config):
    """
    Fait : construit le texte affiche dans la liste pour une piece (mesh par piece isolee).
    Depend de : rien (lit row_config.obj/view_name).
    Retourne : str, le nom de la piece suivi de la vue choisie, separes par " | ".
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("vue=" + row_config.view_name)
    return " | ".join(parts)


def show_only_body(target_body, all_bodies):
    """
    Fait : masque toutes les pieces sauf target_body.
    Depend de : Body.Visible, Transaction (Ansys.ACT.Mechanical).
    Retourne : rien (effet de bord : change la visibilite de tous les corps du modele).
    """
    # Transaction(True) differe le rafraichissement pendant la boucle ; le rendu n'est de toute facon capture que plus tard, lors de l'export d'image.
    target_id = target_body.ObjectId
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Visible = (body.ObjectId == target_id)
            except Exception:
                pass


def show_all_bodies(all_bodies):
    """
    Fait : rend toutes les pieces visibles.
    Depend de : Body.Visible, Transaction (Ansys.ACT.Mechanical).
    Retourne : rien (effet de bord : restaure la visibilite de tous les corps du modele).
    """
    with Transaction(True):
        for body in all_bodies:
            try:
                body.Visible = True
            except Exception:
                pass


def get_body_mesh_counts(body):
    """
    Fait : recupere le nombre de noeuds/elements d'UN corps.
    Depend de : body.VisibleProperties (le panneau Details).
    Retourne : tuple (node_count, element_count), chacun None si indisponible.
    """
    # Lu via VisibleProperties (panneau Details) : MeshRegionById(body.ObjectId) s'est avere mal attribue (valeurs a 0 ou incoherentes entre pieces).
    node_count = None
    element_count = None
    try:
        for prop in body.VisibleProperties:
            if prop.Name == "Nodes":
                node_count = prop.StringValue
            elif prop.Name == "Elements":
                element_count = prop.StringValue
    except Exception as e:
        print "Comptage noeuds/elements impossible pour {} : {}".format(body.Name, str(e))
    return node_count, element_count


def export_body_mesh_image(body, all_bodies, image_name):
    """
    Fait : isole une piece (autres corps masques) et exporte une image de son maillage.
    Depend de : show_only_body, Model.Mesh.AddFigure, export_current_view_image, show_all_bodies.
    Retourne : str, le chemin du PNG genere.
    """
    # ExtAPI.Graphics.ViewOptions.ShowMesh doit deja etre force a True par l'appelant (build_mesh_part_slides) : un seul forcage/reset pour tout un groupe de captures.
    show_only_body(body, all_bodies)
    mesh = ExtAPI.DataModel.Project.Model.Mesh
    figure = mesh.AddFigure()
    figure.Activate()
    # SetFit() est necessaire ici (contrairement au reste du projet, voir README) : show_only_body()
    # masque completement les autres corps (Visible=False, pas de contexte transparent comme pour la
    # geometrie), donc le seul contenu du viewport est la piece ciblee et SetFit() ne peut pas ecraser
    # une autre vue utile -- sans cet appel, la camera garde le cadrage de l'assemblage complet et une
    # petite piece isolee (ex: un boulon) apparait minuscule sur l'image exportee.
    ExtAPI.Graphics.Camera.SetFit()
    image_path = export_current_view_image(image_name)
    show_all_bodies(all_bodies)
    return image_path


def export_body_mesh_summary_csv(directory, body):
    """
    Fait : exporte un tableau minimal de statistiques de maillage (ElementSize, Nodes, Elements) pour UNE piece.
    Depend de : get_body_mesh_counts, Model.Mesh, _format_element_size (01_data_export.py) pour ElementSize, get_unique_file_path/to_csv_cell (00_constants.py), le module csv.
    Retourne : str, le chemin du CSV genere.
    """
    mesh = Model.Mesh
    node_count, element_count = get_body_mesh_counts(body)

    rows = [
        ["ElementSize", _format_element_size(mesh)],
        ["Nodes", node_count],
        ["Elements", element_count],
    ]

    filepath = get_unique_file_path(directory, "Mesh_" + safe_file_name(body.Name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Propriete", "Valeur"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "Export CSV termine : " + filepath
    return filepath


def add_mesh_multi_image_slide(report, image_paths, csv_paths):
    """
    Fait : ajoute une slide LAYOUT_MESH_MULTI (jusqu'a 4 images + 4 tableaux) pour un groupe de pieces.
    Depend de : LAYOUT_MESH_MULTI, MESH_MULTI_IMAGE_SHAPE_INDICES, MESH_MULTI_TABLE_SHAPE_INDICES, report.add_csv_table.
    Retourne : PPT.Slide, la slide creee.
    """
    layout = report.presentation.SlideMaster.CustomLayouts[LAYOUT_MESH_MULTI]
    slide = report.presentation.Slides.AddSlide(report.presentation.Slides.Count + 1, layout)

    slide.Shapes[1].TextFrame.TextRange.Text = "Mesh Details"

    # Tri par position reelle (haut->bas, gauche->droite), pas par indice de shape, meme precaution que add_multi_step_image_slide.
    image_placeholders = [slide.Shapes[i] for i in MESH_MULTI_IMAGE_SHAPE_INDICES]
    image_placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    table_placeholders = [slide.Shapes[i] for i in MESH_MULTI_TABLE_SHAPE_INDICES]
    table_placeholders.sort(key=lambda ph: (ph.Top, ph.Left))

    for i in range(min(len(image_paths), len(image_placeholders))):
        ph = image_placeholders[i]
        if image_paths[i]:
            slide.Shapes.AddPicture(image_paths[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                     ph.Left, ph.Top, ph.Width, ph.Height)

    for i in range(min(len(csv_paths), len(table_placeholders))):
        ph = table_placeholders[i]
        if csv_paths[i]:
            try:
                report.add_csv_table(slide, csv_paths[i], ph.Left, ph.Top, ph.Width)
            except Exception as e:
                print "Insertion de table impossible ({}) : {}".format(csv_paths[i], str(e))

    return slide


def build_mesh_part_slides(report, row_configs, all_bodies, views):
    """
    Fait : ajoute une ou plusieurs slides "mesh par piece isolee", en regroupant les pieces par MAX_MESH_MULTI_BODIES.
    Depend de : apply_view_if_exists, export_body_mesh_image, export_body_mesh_summary_csv, add_mesh_multi_image_slide.
    Retourne : rien (effet de bord : ajoute une ou plusieurs slides au rapport).
    """
    for start in range(0, len(row_configs), MAX_MESH_MULTI_BODIES):
        chunk = row_configs[start:start + MAX_MESH_MULTI_BODIES]

        ExtAPI.Graphics.ViewOptions.ShowMesh = True
        image_paths = []
        try:
            for cfg in chunk:
                body = cfg.obj
                try:
                    apply_view_if_exists(cfg.view_name, views)
                    image_paths.append(export_body_mesh_image(body, all_bodies, "Mesh_" + safe_file_name(body.Name)))
                except Exception as e:
                    print "Export image impossible pour {} : {}".format(body.Name, str(e))
                    image_paths.append(None)
        finally:
            ExtAPI.Graphics.ViewOptions.ShowMesh = False

        csv_paths = []
        for cfg in chunk:
            body = cfg.obj
            try:
                csv_paths.append(export_body_mesh_summary_csv(CSV_EXPORT_FOLDER, body))
            except Exception as e:
                print "Export CSV impossible pour {} : {}".format(body.Name, str(e))
                csv_paths.append(None)

        add_mesh_multi_image_slide(report, image_paths, csv_paths)
        print "Slide Mesh multi-image ajoutee ({} piece(s)).".format(len(chunk))


class ContactRowConfig(object):
    """
    Ligne de selection pour la slide Contact summary : juste le contact, rien a configurer (toutes les lignes cochees partagent UNE seule slide).
    """

    def __init__(self, contact_region):
        """
        Fait : initialise la configuration d'une ligne Contact summary (rien a configurer).
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.obj = contact_region
        self.configured = True  # pas de bouton "..." pour cette categorie : toujours "pret"


def build_contact_row_display_name(row_config):
    """
    Fait : construit le texte affiche dans la liste pour une Contact Region.
    Depend de : rien (lit row_config.obj.Name).
    Retourne : str, le nom du contact.
    """
    return row_config.obj.Name


def collect_contact_regions():
    """
    Fait : liste toutes les Contact Region du modele.
    Depend de : ExtAPI.DataModel.GetObjectsByType(ContactRegion).
    Retourne : list, les objets Contact Region (dossier Connections).
    """
    return list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion))


def build_contact_summary_slide(report, row_configs):
    """
    Fait : ajoute LA slide de resume des contacts, limitee aux contacts selectionnes.
    Depend de : export_contacts_summary_csv (01_data_export.py), report.add_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport).
    """
    contact_list = [cfg.obj for cfg in row_configs]
    csv_path = export_contacts_summary_csv(CSV_EXPORT_FOLDER, contact_list)
    report.add_table_slide("Contacts summary", "-- Contact --", csv_path)


CURVE_COLOR_OPTIONS = [
    ("Automatique", None),
    ("Rouge", Color.IndianRed),
    ("Bleu", Color.SteelBlue),
    ("Vert", Color.SeaGreen),
    ("Orange", Color.DarkOrange),
    ("Violet", Color.MediumPurple),
    ("Noir", Color.Black),
    ("Gris", Color.Gray),
]


def curve_color_label(color):
    """
    Fait : trouve le libelle affiche pour une couleur de courbe.
    Depend de : CURVE_COLOR_OPTIONS.
    Retourne : str, le libelle correspondant, ou "Automatique" si color est None ou inconnue.
    """
    if color is not None:
        for label, option_color in CURVE_COLOR_OPTIONS:
            if option_color is not None and option_color == color:
                return label
    return CURVE_COLOR_OPTIONS[0][0]


def curve_color_from_label(label):
    """
    Fait : trouve la couleur associee a un libelle de CURVE_COLOR_OPTIONS.
    Depend de : CURVE_COLOR_OPTIONS.
    Retourne : Color ou None, la couleur correspondante (None pour "Automatique" ou libelle inconnu).
    """
    for option_label, option_color in CURVE_COLOR_OPTIONS:
        if option_label == label:
            return option_color
    return None


class SolutionInfoRowConfig(object):
    """
    Configuration d'affichage pour un tracker de Solution Information : l'objet et les parametres du graphique (titre, axes, couleur), None = deduit du CSV.
    """

    def __init__(self, tracker, analysis=None):
        """
        Fait : initialise la configuration d'un tracker de Solution Information avec ses valeurs par defaut.
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.obj = tracker
        self.analysis = analysis  # voir SlideRowConfig.analysis / analysis_suffix
        self.chart_title = None
        self.x_axis_label = None
        self.y_axis_label = None
        self.curve_color = None
        self.configured = False


def build_solution_info_row_display_name(row_config):
    """
    Fait : construit le texte affiche dans la liste pour un tracker de Solution Information.
    Depend de : analysis_suffix, row_config (chart_title, x_axis_label, y_axis_label, curve_color).
    Retourne : str, le nom du tracker suivi des reglages de graphique choisis, separes par " | ".
    """
    parts = [row_config.obj.Name + analysis_suffix(row_config)]
    if row_config.chart_title:
        parts.append("titre=" + row_config.chart_title)
    if row_config.x_axis_label:
        parts.append("x=" + row_config.x_axis_label)
    if row_config.y_axis_label:
        parts.append("y=" + row_config.y_axis_label)
    if row_config.curve_color is not None:
        parts.append("couleur=" + curve_color_label(row_config.curve_color))
    return " | ".join(parts)


def collect_solution_information_trackers(analysis):
    """
    Fait : liste les trackers (enfants) de Solution Information pour une analyse.
    Depend de : analysis.Solution.Children[0] (1er enfant de la branche Solution).
    Retourne : list, les objets tracker, vide en cas d'erreur.
    """
    try:
        solution_information = analysis.Solution.Children[0]
        children = solution_information.Children
        return list(children) if children else []
    except Exception as e:
        print "Solution Information indisponible : " + str(e)
        return []


def collect_solution_information_trackers_multi(analyses):
    """
    Fait : liste les trackers de Solution Information de toutes les analyses fournies.
    Depend de : collect_solution_information_trackers.
    Retourne : list de tuples (obj, analysis).
    """
    pairs = []
    for analysis in analyses:
        for tracker in collect_solution_information_trackers(analysis):
            pairs.append((tracker, analysis))
    return pairs


def build_single_solution_info_slide(report, cfg):
    """
    Fait : construit la slide d'UN tracker de Solution Information, avec ses parametres de graphique eventuels.
    Depend de : export_result_tabular_data, export_chart_image_from_csv (02_image_export.py), get_scoped_contact_region_name (04_slides.py), report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport, ou rien si aucune donnee exportable).
    """
    tracker = cfg.obj
    csv_path = None
    try:
        csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, tracker)
    except Exception as e:
        print "Export CSV impossible pour {} : {}".format(tracker.Name, str(e))

    img_path = None
    if csv_path:
        try:
            img_path = export_chart_image_from_csv(
                csv_path, tracker.Name, chart_title=cfg.chart_title, x_axis_label=cfg.x_axis_label,
                y_axis_label=cfg.y_axis_label, curve_color=cfg.curve_color
            )
        except Exception as e:
            print "Construction du graphique impossible pour {} : {}".format(tracker.Name, str(e))

    if img_path or csv_path:
        contact_region_name = get_scoped_contact_region_name(tracker)
        title = tracker.Name + analysis_suffix(cfg)
        if contact_region_name:
            title = "{} - {}".format(title, contact_region_name)
        report.add_image_table_slide(title, "-- Solution Information --", img_path=img_path, csv_path=csv_path)
    else:
        print "Aucune donnee exportable pour " + tracker.Name + " : slide ignoree."


def build_solution_info_slides(report, row_configs):
    """
    Fait : ajoute une slide pour chaque tracker de Solution Information selectionne.
    Depend de : build_single_solution_info_slide.
    Retourne : rien (effet de bord : ajoute une slide par tracker au rapport).
    """
    for cfg in row_configs:
        build_single_solution_info_slide(report, cfg)


def export_mesh_summary_csv(directory):
    """
    Fait : exporte un tableau minimal des statistiques de maillage (ElementSize, Nodes, Elements).
    Depend de : Model.Mesh, le module csv, _format_element_size (01_data_export.py) pour ElementSize.
    Retourne : str, le chemin du CSV genere.
    """
    mesh = Model.Mesh
    rows = [
        ["ElementSize", _format_element_size(mesh)],
        ["Nodes", mesh.Nodes],
        ["Elements", mesh.Elements],
    ]

    filepath = os.path.join(directory, "mesh_summary.csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Propriete", "Valeur"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "Export CSV termine : " + filepath
    return filepath


def build_mesh_slide(report, use_full_table):
    """
    Fait : ajoute la slide maillage (vue + tableau complet ou resume selon use_full_table).
    Depend de : export_mesh_image, export_mesh_report_csv (01_data_export.py), export_mesh_summary_csv, report.add_image_table_slide.
    Retourne : rien (effet de bord : ajoute une slide au rapport).
    """
    img_path = export_mesh_image()
    if use_full_table:
        csv_path = export_mesh_report_csv(CSV_EXPORT_FOLDER)
    else:
        csv_path = export_mesh_summary_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Mesh and mesh details", "-- Mesh --", img_path=img_path, csv_path=csv_path)


class AnalysisContextRowConfig(object):
    """
    Ligne de selection pour une slide de contexte (Analysis Parameters) : l'analyse elle-meme et une vue (View Manager) optionnelle.
    """

    def __init__(self, analysis):
        """
        Fait : initialise la configuration d'une ligne Contexte d'analyse (vue optionnelle).
        Depend de : rien (affectations simples).
        Retourne : rien (constructeur).
        """
        self.obj = analysis
        self.view_name = None
        self.configured = False  # passe a True des que le bouton "..." a ete valide (OK)


def build_analysis_context_row_display_name(row_config):
    """
    Fait : construit le texte affiche dans la liste pour une analyse.
    Depend de : rien (lit row_config.obj.Name/view_name).
    Retourne : str, le nom de l'analyse, suivi de la vue choisie si definie.
    """
    parts = [row_config.obj.Name]
    if row_config.view_name:
        parts.append("vue=" + row_config.view_name)
    return " | ".join(parts)


def build_analysis_context_slides(report, row_configs, views):
    """
    Fait : ajoute une slide de contexte (Analysis Parameters) pour chaque analyse selectionnee, avec sa vue configuree.
    Depend de : apply_view_if_exists, create_analysis_parameters_slide (04_slides.py).
    Retourne : rien (effet de bord : ajoute une slide par analyse au rapport).
    """
    for cfg in row_configs:
        apply_view_if_exists(cfg.view_name, views)
        create_analysis_parameters_slide(report, cfg.obj)
