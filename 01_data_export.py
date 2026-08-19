# 01_data_export.py : extraction de donnees - tout ce qui lit le modele / le Tabular Data pane et ecrit des fichiers CSV. Depend de 00_constants.py (doit etre execute avant ce fichier).

import csv
import os
import materials


def export_active_tabular_data(directory, filename, start_col=3):
    """
    Fait : exporte le contenu du Tabular Data pane actuellement affiche vers un fichier CSV.
    Depend de : ExtAPI.UserInterface (pane Mechanical actif) ; l'appelant doit avoir fait Activate() sur l'objet avant l'appel.
    Retourne : str, le chemin complet du CSV ecrit.
    """
    # start_col saute les premieres colonnes du pane (numero de ligne / step), non pertinentes pour l'export.
    pane = ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    control = pane.ControlUnknown
    num_columns = control.ColumnsCount + 1
    num_rows = control.RowsCount + 1

    rows = []
    for row in range(1, num_rows):
        line = [clean_cell_text(control.cell(row, col).Text) for col in range(start_col, num_columns)]
        if any(cell != "" for cell in line):
            rows.append(line)

    filepath = os.path.join(directory, filename)
    with open(filepath, 'wb') as f:  # 'wb' pour eviter les doubles retours a la ligne sous Windows
        writer = csv.writer(f, delimiter=';')
        for line in rows:
            writer.writerow([to_csv_cell(cell) for cell in line])

    print "CSV exporte : " + filepath
    return filepath


def export_bc_tabular_data(directory, bc):
    """
    Fait : exporte les donnees tabulaires d'une Boundary Condition donnee.
    Depend de : export_active_tabular_data (apres Activate() de bc).
    Retourne : str, le chemin du CSV.
    """
    bc.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(bc.Name), start_col=3)


def export_bp_tabular_data(directory, bp):
    """
    Fait : exporte les donnees tabulaires d'une Bolt Pretension donnee.
    Depend de : export_active_tabular_data (apres Activate() de bp).
    Retourne : str, le chemin du CSV.
    """
    bp.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(bp.Name), start_col=3)


def export_result_tabular_data(directory, obj):
    """
    Fait : exporte les donnees tabulaires d'un objet resultat quelconque (resultat de solution, enfant de Contact/Bolt Tool, tracker de Solution Information...).
    Depend de : export_active_tabular_data ; obj doit supporter .Activate() et .Name.
    Retourne : str, le chemin du CSV.
    """
    # Tous ces types d'objets partagent le meme layout de Tabular Data pane (colonne 1 ignoree, donnees a partir de la colonne 2).
    obj.Activate()
    return export_active_tabular_data(directory, "{}.csv".format(obj.Name), start_col=2)


def export_contacts_summary_csv(directory, contact_list=None):
    """
    Fait : exporte un tableau resume (type, frottement, raideur, tolerances, traitement d'interface) pour chaque Contact Region du modele.
    Depend de : ExtAPI.DataModel (si contact_list est None) et _get_prop pour lire les proprietes de facon securisee.
    Retourne : str, le chemin du CSV.
    """
    filepath = os.path.join(directory, "contact_info_export.csv")
    if contact_list is None:
        # None = comportement historique : toutes les Contact Region du modele (voir create_contact_summary_slide de 04_slides.py).
        contact_list = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion)

    with open(filepath, 'wb') as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "Name", "Contact Type", "Friction Coefficient", "Normal Stiffness Factor",
            "Penetration Tolerance", "Penetration Tolerance Value",
            "Elastic Slip Tolerance", "Elastic Slip Tolerance Value",
            "Interface Treatment", "Offset",
        ])

        for contact in contact_list:
            try:
                friction = ""
                if contact.ContactType == ContactType.Frictional:
                    friction = contact.FrictionCoefficient

                penetration_tolerance = _get_prop(contact, "PenetrationTolerance")
                penetration_tolerance_value = ""
                if penetration_tolerance is not None and "Value" in str(penetration_tolerance):
                    penetration_tolerance_value = _get_prop(contact, "PenetrationToleranceValue")

                elastic_slip_tolerance = _get_prop(contact, "ElasticSlipTolerance")
                elastic_slip_tolerance_value = ""
                if elastic_slip_tolerance is not None and "Value" in str(elastic_slip_tolerance):
                    elastic_slip_tolerance_value = _get_prop(contact, "ElasticSlipToleranceValue")

                interface_treatment = _get_prop(contact, "InterfaceTreatment")
                offset_value = ""
                if interface_treatment is not None and "Offset" in str(interface_treatment):
                    offset_value = _get_prop(contact, "Offset")

                writer.writerow([
                    to_csv_cell(contact.Name), to_csv_cell(contact.ContactType),
                    to_csv_cell(friction), to_csv_cell(contact.NormalStiffnessFactor),
                    to_csv_cell(penetration_tolerance), to_csv_cell(penetration_tolerance_value),
                    to_csv_cell(elastic_slip_tolerance), to_csv_cell(elastic_slip_tolerance_value),
                    to_csv_cell(interface_treatment), to_csv_cell(offset_value),
                ])
            except Exception as e:
                print "Erreur sur contact {} : {}".format(contact.Name, str(e))

    print "Export CSV termine : " + filepath
    return filepath


def export_mesh_report_csv(directory):
    """
    Fait : exporte un rapport complet des parametres et statistiques du maillage (defauts, sizing, qualite, inflation, avance, statistiques).
    Depend de : Model.Mesh et _get_prop pour lire les proprietes de facon securisee, _format_element_size pour ElementSize.
    Retourne : str, le chemin du CSV.
    """
    mesh = Model.Mesh
    rows = []

    rows.append(["Defaults", "PhysicsPreference", mesh.PhysicsPreference])
    rows.append(["Defaults", "ElementOrder", mesh.ElementOrder])
    rows.append(["Defaults", "ElementSize", _format_element_size(mesh)])

    for p in ["UseAdaptiveSizing", "Resolution", "MeshDefeaturing", "DefeatureSize",
              "TransitionOption", "SpanAngleCenter", "CurvatureNormalAngle",
              "MinSize", "MaxSize", "GrowthRate"]:
        rows.append(["Sizing", p, _get_prop(mesh, p)])

    for p in ["MeshMetric", "ErrorLimits", "TargetQuality", "Smoothing"]:
        rows.append(["Quality", p, _get_prop(mesh, p)])

    for p in ["UseAutomaticInflation", "InflationOption", "NumberOfLayers", "InflationGrowthRate"]:
        rows.append(["Inflation", p, _get_prop(mesh, p)])

    for p in ["NumberOfCPUsForParallelPartMeshing", "StraightSidedElements", "RigidBodyBehavior"]:
        rows.append(["Advanced", p, _get_prop(mesh, p)])

    rows.append(["Statistics", "Nodes", mesh.Nodes])
    rows.append(["Statistics", "Elements", mesh.Elements])

    try:
        metric_data = mesh.MeshMetricValues
        if metric_data and len(metric_data) > 0:
            rows.append(["Statistics", "MeshMetricValueMin", min(metric_data)])
            rows.append(["Statistics", "MeshMetricValueMax", max(metric_data)])
            rows.append(["Statistics", "MeshMetricValueAvg", sum(metric_data) / len(metric_data)])
    except Exception:
        rows.append(["Statistics", "MeshMetricValues", "Non disponible"])

    filepath = os.path.join(directory, "mesh_report.csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Section", "Propriete", "Valeur"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "Export CSV termine : " + filepath
    return filepath


def _get_prop(obj, prop_name):
    """
    Fait : recupere une propriete d'un objet Mechanical de maniere securisee.
    Depend de : getattr natif Python.
    Retourne : la valeur de la propriete, ou None si absente / si l'acces leve une exception.
    """
    try:
        return getattr(obj, prop_name)
    except Exception:
        return None


def _format_element_size(mesh):
    """
    Fait : lit ElementSize et distingue le sizing "Default" (aucune valeur saisie par l'utilisateur,
    Quantity nulle) d'une valeur reellement definie.
    Depend de : _get_prop, mesh.ElementSize (Quantity .NET, expose .Value) - quand Element Size est
    laisse sur "Default" dans Mechanical, la taille reelle est calculee dynamiquement au maillage et
    n'est jamais ecrite dans cette propriete, qui reste alors a 0.
    Retourne : str "Default" si ElementSize vaut 0, sinon la Quantity brute (traitee normalement par to_csv_cell).
    """
    element_size = _get_prop(mesh, "ElementSize")
    try:
        if element_size is not None and float(element_size.Value) == 0.0:
            return "Default"
    except Exception:
        pass
    return element_size


def export_materials_csv(directory):
    """
    Fait : exporte une ligne par materiau distinct utilise par les bodies du modele (module, densite, coefficient de Poisson, dilatation, conductivite, chaleur specifique).
    Depend de : ExtAPI.DataModel, module materials (API Ansys), _material_property_values / _material_property_units.
    Retourne : str, le chemin du CSV.
    """
    bodies = ExtAPI.DataModel.Project.Model.GetChildren(DataModelObjectCategory.Body, True)

    seen_materials = []
    for body in bodies:
        if body.Material not in seen_materials:
            seen_materials.append(body.Material)

    filepath = os.path.join(directory, "materials_export.csv")

    with open(filepath, 'wb') as f:
        first_material = bodies[0].GetGeoBody().Material
        units = _material_property_units(first_material)

        header = (
            "Material;Young's Modulus [{}];Density [{}];Poisson's Ratio [-];"
            "Thermal Expansion [{}];Thermal Conductivity [{}];Specific Heat [{}]\n"
        ).format(units["Young's Modulus"], units["Density"], units["Coefficient of Thermal Expansion"],
                 units["Thermal Conductivity"], units["Specific Heat"])
        f.write(to_csv_cell(header))

        for mat in seen_materials:
            body_for_mat = None
            for body in bodies:
                if body.Material == mat:
                    body_for_mat = body
                    break

            if body_for_mat is None:
                continue

            material = body_for_mat.GetGeoBody().Material
            values = _material_property_values(material)

            line = "{};{};{};{};{};{};{}\n".format(
                mat,
                values["Young's Modulus"],
                values["Density"],
                values["Poisson's Ratio"],
                values["Coefficient of Thermal Expansion"],
                values["Thermal Conductivity"],
                values["Specific Heat"],
            )
            f.write(to_csv_cell(line))

    print "Export CSV termine : " + filepath
    return filepath


def _convert_young_modulus_to_gpa(prop_name, unit, value):
    """
    Fait : convertit le module de Young de Pa en GPa (plus lisible dans le tableau materiaux de la slide geometrie).
    Depend de : rien (conversion numerique simple) ; ne s'applique que si prop_name == "Young's Modulus" et
    l'unite source vaut "Pa", pour ne jamais reconvertir une valeur deja recuperee dans une autre unite
    (ex : projet configure en MPa/psi).
    Retourne : tuple (unit, value) convertis si applicable, sinon (unit, value) inchanges.
    """
    if prop_name == "Young's Modulus" and unit == "Pa":
        converted_value = (value / 1.0e9) if value is not None else None
        return "GPa", converted_value
    return unit, value


def _material_property_values(material):
    """
    Fait : aplatit les 5 groupes de proprietes materiau utilises dans le rapport en un seul dict de valeurs.
    Depend de : materials.GetMaterialPropertyByName (API Ansys), _convert_young_modulus_to_gpa.
    Retourne : dict, nom de propriete -> valeur (premiere valeur seulement si la propriete depend de la temperature ; Young's Modulus converti en GPa si recupere en Pa).
    """
    values = {}
    for group in ["Elasticity", "Density", "Coefficient of Thermal Expansion",
                  "Thermal Conductivity", "Specific Heat"]:
        for prop_name, prop_data in materials.GetMaterialPropertyByName(material, group).items():
            # prop_data = (unit, value) pour une propriete constante, ou (unit, value_T1, value_T2, ...) si dependante de la temperature : seule la 1ere valeur est gardee, pour rester sur une ligne par materiau.
            unit = prop_data[0]
            value = prop_data[1] if len(prop_data) > 1 else prop_data[0]
            _, value = _convert_young_modulus_to_gpa(prop_name, unit, value)
            values[prop_name] = value
    return values


def _material_property_units(material):
    """
    Fait : aplatit les 5 groupes de proprietes materiau utilises dans le rapport en un seul dict d'unites.
    Depend de : materials.GetMaterialPropertyByName (API Ansys), _convert_young_modulus_to_gpa.
    Retourne : dict, nom de propriete -> unite ("GPa" pour Young's Modulus si recupere en Pa).
    """
    units = {}
    for group in ["Elasticity", "Density", "Coefficient of Thermal Expansion",
                  "Thermal Conductivity", "Specific Heat"]:
        for prop_name, prop_data in materials.GetMaterialPropertyByName(material, group).items():
            unit, _ = _convert_young_modulus_to_gpa(prop_name, prop_data[0], None)
            units[prop_name] = unit
    return units


def export_analysis_settings_csv(directory, analysis):
    """
    Fait : exporte un tableau des parametres de steps de l'analyse (Analysis Settings), transpose (un Loadcase par colonne, une propriete par ligne) pour eviter la repetition d'un tableau vertical classique.
    Depend de : analysis.AnalysisSettings (API Ansys), get_unique_file_path/safe_file_name/to_csv_cell (00_constants.py), le module csv.
    Retourne : str, le chemin du CSV genere.
    """
    settings = analysis.AnalysisSettings

    try:
        num_steps = settings.NumberOfSteps
    except Exception:
        num_steps = 0

    end_times = []
    define_bys = []
    auto_steppings = []
    substep_counts = []

    for step in range(1, num_steps + 1):
        try:
            end_times.append(settings.GetStepEndTime(step))
        except Exception:
            end_times.append(None)
        try:
            define_bys.append(settings.GetDefineBy(step))
        except Exception:
            define_bys.append(None)
        try:
            auto_steppings.append(settings.GetAutomaticTimeStepping(step))
        except Exception:
            auto_steppings.append(None)
        try:
            substep_counts.append(settings.GetNumberOfSubSteps(step))
        except Exception:
            substep_counts.append(None)

    header = ["Propriete"] + ["Loadcase {}".format(step) for step in range(1, num_steps + 1)]
    rows = [
        ["End time"] + end_times,
        ["Define by"] + define_bys,
        ["Auto time stepping"] + auto_steppings,
        ["Substeps"] + substep_counts,
    ]

    filepath = get_unique_file_path(directory, "AnalysisSettings_" + safe_file_name(analysis.Name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([to_csv_cell(v) for v in header])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "Export CSV termine : " + filepath
    return filepath


SOLUTION_INFO_PROPERTIES = [  # nom de propriete Solution -> libelle affiche ; liste (et non dict) pour garder un ordre d'affichage stable
    ("ElapsedRunTime", "Elapsed run time"),
    ("MemoryUsed", "Memory used"),
    ("ResultFileSize", "Result file size"),
]


def export_solution_info_csv(directory, solution, analysis_name):
    """
    Fait : exporte un tableau des informations de resolution (temps de calcul, memoire utilisee, taille du fichier resultat).
    Depend de : solution.PropertyByName (API Ansys), constante SOLUTION_INFO_PROPERTIES, get_unique_file_path/safe_file_name/to_csv_cell (00_constants.py), le module csv.
    Retourne : str, le chemin du CSV genere.
    """
    rows = []
    for prop_name, label in SOLUTION_INFO_PROPERTIES:
        try:
            rows.append([label, solution.PropertyByName(prop_name).StringValue])
        except Exception:
            pass

    # analysis_name (pas solution.Name, generique type "Solution" sur toutes les analyses) pour un
    # nom de fichier distinct par analyse en projet multi-analyses.
    filepath = get_unique_file_path(directory, "SolutionInfo_" + safe_file_name(analysis_name), ".csv")
    with open(filepath, "wb") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Propriete", "Valeur"])
        for row in rows:
            writer.writerow([to_csv_cell(v) for v in row])

    print "Export CSV termine : " + filepath
    return filepath
