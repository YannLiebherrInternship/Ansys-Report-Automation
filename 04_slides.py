# 04_slides.py : construction des slides - chaque create_..._slide(report) extrait les donnees dont elle a besoin (CSV + image) puis ajoute une slide a l'instance PPTReportBuilder fournie. Depend de 00_constants.py, 01_data_export.py, 02_image_export.py et 03_ppt_utils.py (doivent etre executes avant ce fichier).


def create_geometry_slide(report):
    """
    Fait : ajoute la slide de contexte geometrie + materiaux.
    Depend de : export_geometry_image, export_materials_csv, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report).
    """
    img_path = export_geometry_image()
    csv_path = export_materials_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Geometry and materials details", "-- Geometry and materials --",
                                  img_path=img_path, csv_path=csv_path)


def create_mesh_slide(report):
    """
    Fait : ajoute la slide de contexte maillage.
    Depend de : export_mesh_image, export_mesh_report_csv, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report).
    """
    img_path = export_mesh_image()
    csv_path = export_mesh_report_csv(CSV_EXPORT_FOLDER)
    report.add_image_table_slide("Mesh and mesh details", "-- Mesh --",
                                  img_path=img_path, csv_path=csv_path)


def create_analysis_parameters_slide(report, analysis=None):
    """
    Fait : ajoute la slide de contexte Analysis Parameters (vue d'ensemble + tableau des steps + tableau des infos de resolution) pour une analyse donnee.
    Depend de : export_analysis_overview_image, export_analysis_settings_csv, export_solution_info_csv, report.add_analysis_context_slide.
    Retourne : rien (effet de bord sur report).
    """
    # analysis=None (comportement d'origine) : utilise Analyses[0], conserve pour les appelants sans argument (ex : code obsolete AnsysReportGenerator_GUI.py).
    analysis = analysis or ExtAPI.DataModel.Project.Model.Analyses[0]
    img_path = export_analysis_overview_image(analysis)
    settings_csv_path = export_analysis_settings_csv(CSV_EXPORT_FOLDER, analysis)
    solution_csv_path = export_solution_info_csv(CSV_EXPORT_FOLDER, analysis.Solution, analysis.Name)
    report.add_analysis_context_slide(analysis.Name, "Analysis Parameters",
                                       img_path, settings_csv_path, solution_csv_path)


def create_bc_slide(report):
    """
    Fait : ajoute une slide pour chaque Boundary Condition trouvee dans le modele.
    Depend de : ExtAPI.DataModel, export_object_image, export_bc_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report ; ne fait rien si aucune BC n'est trouvee).
    """
    bc_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.GenericBoundaryCondition)
    if not bc_list:
        print "No Boundary Condition: slide skipped."
        return
    for bc in bc_list:
        img_path = export_object_image(bc, bc.Name)
        csv_path = export_bc_tabular_data(CSV_EXPORT_FOLDER, bc)
        report.add_image_table_slide(bc.Name, "-- Boundary Conditions --",
                                      img_path=img_path, csv_path=csv_path)


def create_bp_slide(report):
    """
    Fait : ajoute une slide pour chaque Bolt Pretension trouvee dans le modele.
    Depend de : ExtAPI.DataModel, export_object_image, export_bp_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report ; ne fait rien si aucune Bolt Pretension n'est trouvee).
    """
    bp_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltPretension)
    if not bp_list:
        print "No Bolt Pretension: slide skipped."
        return
    for bp in bp_list:
        img_path = export_object_image(bp, bp.Name)
        csv_path = export_bp_tabular_data(CSV_EXPORT_FOLDER, bp)
        report.add_image_table_slide(bp.Name, "-- Bolt Pretension --",
                                      img_path=img_path, csv_path=csv_path)


def create_contact_summary_slide(report):
    """
    Fait : ajoute la slide de resume des contacts (table seule, sans image).
    Depend de : export_contacts_summary_csv, report.add_table_slide.
    Retourne : rien (effet de bord sur report).
    """
    csv_path = export_contacts_summary_csv(CSV_EXPORT_FOLDER)
    report.add_table_slide("Contacts summary", "-- Contact --", csv_path)


def create_tool_children_slides(report, category, subtitle, include_table=False):
    """
    Fait : ajoute une slide pour chaque ENFANT de chaque objet outil trouve pour la categorie donnee (les resultats a exporter sont les enfants de l'outil, pas l'outil lui-meme).
    Depend de : ExtAPI.DataModel, export_object_image, export_result_tabular_data (si include_table), report.add_image_table_slide.
    Retourne : rien (effet de bord sur report ; ignore les outils sans objet trouve ou sans enfant).
    """
    tools = ExtAPI.DataModel.GetObjectsByType(category)
    if not tools:
        print "No object found for: " + subtitle
        return

    for tool in tools:
        children = tool.Children
        if children is None or len(children) == 0:
            print "No child under " + tool.Name + ": skipped."
            continue
        for child in children:
            try:
                img_path = export_object_image(child, child.Name)
            except Exception as e:
                print "Unable to export image for {}: {}".format(child.Name, str(e))
                img_path = None

            csv_path = None
            if include_table:
                try:
                    csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, child)
                except Exception as e:
                    print "Unable to export CSV for {}: {}".format(child.Name, str(e))

            if img_path or csv_path:
                report.add_image_table_slide(child.Name, subtitle, img_path=img_path, csv_path=csv_path)
            else:
                print "No exportable data for " + child.Name + ": slide skipped."


def create_contact_tool_slide(report):
    """
    Fait : ajoute une slide (avec donnees tabulaires) pour chaque resultat enfant de chaque Contact Tool trouve.
    Depend de : create_tool_children_slides.
    Retourne : rien (effet de bord sur report).
    """
    create_tool_children_slides(report, DataModelObjectCategory.ContactTool, "-- Contact Tool --", include_table=True)


def create_bolt_tool_slide(report):
    """
    Fait : ajoute une slide (avec donnees tabulaires) pour chaque resultat enfant de chaque Bolt Tool trouve.
    Depend de : create_tool_children_slides.
    Retourne : rien (effet de bord sur report).
    """
    create_tool_children_slides(report, DataModelObjectCategory.BoltTool, "-- Bolt Tool --", include_table=True)


def get_scoped_contact_region_name(obj):
    """
    Fait : recupere le nom de la Contact Region associee a un objet (champ Scope > Contact Region), quand il en a une.
    Depend de : obj.ContactRegion (API Ansys).
    Retourne : str, le nom de la Contact Region, ou None si non disponible.
    """
    try:
        contact_region = obj.ContactRegion
        if contact_region is None:
            return None
        return contact_region.Name
    except Exception:
        return None


def create_solution_information_slide(report):
    """
    Fait : ajoute une slide pour chaque tracker enfant de "Solution Information" (ex : Pressure, Max Normal Stiffness, Elastic Slip).
    Depend de : ExtAPI.DataModel, export_result_tabular_data, export_chart_image_from_csv, get_scoped_contact_region_name, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report ; ne fait rien si aucun tracker n'est trouve).
    """
    # Ces objets n'affichent qu'un graphique 2D, pas de vue 3D : on exporte les donnees tabulaires puis on reconstruit le graphique en image a partir du CSV plutot que de capturer le viewport.
    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    solution_information = analysis.Solution.Children[0]
    children = solution_information.Children

    if children is None or len(children) == 0:
        print "No tracker under Solution Information: slide skipped."
        return

    for child in children:
        csv_path = None
        try:
            csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, child)
        except Exception as e:
            print "Unable to export CSV for {}: {}".format(child.Name, str(e))

        img_path = None
        if csv_path:
            try:
                img_path = export_chart_image_from_csv(csv_path, child.Name)
            except Exception as e:
                print "Unable to build chart for {}: {}".format(child.Name, str(e))

        if img_path or csv_path:
            contact_region_name = get_scoped_contact_region_name(child)
            title = child.Name
            if contact_region_name:
                # Titre suffixe par la Contact Region scopee (Details > Scope), pour savoir a quel contact ces valeurs se rapportent.
                title = "{} - {}".format(child.Name, contact_region_name)

            report.add_image_table_slide(title, "-- Solution Information --",
                                          img_path=img_path, csv_path=csv_path)
        else:
            print "No exportable data for " + child.Name + ": slide skipped."


def get_all_simple_results():
    """
    Fait : renvoie tous les resultats simples de la branche Solution (tous les enfants sauf Solution Information et les dossiers d'outils Contact/Bolt Tool, deja traites ailleurs).
    Depend de : ExtAPI.DataModel.
    Retourne : list, les objets resultat restants.
    """
    excluded_categories = [DataModelObjectCategory.ContactTool, DataModelObjectCategory.BoltTool]

    # Exclusion par identite en plus de la categorie : garantit que tout objet deja traite par create_contact_tool_slide / create_bolt_tool_slide est ignore ici, meme si sa categorie ne matche pas exactement.
    already_handled = list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactTool))
    already_handled += list(ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.BoltTool))

    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    children = analysis.Solution.Children

    results = []
    for i in range(1, len(children)):  # index 0 = Solution Information, exclu de la boucle
        child = children[i]
        if child.DataModelObjectCategory in excluded_categories:
            continue
        if child in already_handled:
            continue
        results.append(child)
    return results


def create_result_slide(report):
    """
    Fait : ajoute une slide pour chaque resultat simple trouve sous Solution.
    Depend de : get_all_simple_results, export_solution_image, export_result_tabular_data, report.add_image_table_slide.
    Retourne : rien (effet de bord sur report ; ne fait rien si aucun resultat simple n'est trouve).
    """
    analysis = ExtAPI.DataModel.Project.Model.Analyses[0]
    results = get_all_simple_results()
    if not results:
        print "No simple result: slide skipped."
        return
    for result in results:
        img_path = export_solution_image(result)
        csv_path = export_result_tabular_data(CSV_EXPORT_FOLDER, result)
        report.add_image_table_slide(result.Name, analysis.Name, img_path=img_path, csv_path=csv_path)
