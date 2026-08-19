# 02_image_export.py : export d'images - capture de la vue graphique Mechanical pour geometrie/maillage/BC/resultats, et reconstruction de graphique a partir d'un CSV pour les trackers de Solution Information. Depend de 00_constants.py (doit etre execute avant ce fichier).

import csv

import clr
clr.AddReference("System.Drawing")

from System.Drawing import (Bitmap, Color, Font, FontFamily, FontStyle, Graphics, Pen, PointF,
                             RectangleF, SolidBrush, StringAlignment, StringFormat)
from System.Drawing.Drawing2D import SmoothingMode
from System.Drawing.Imaging import ImageFormat

CHART_COLORS = [Color.IndianRed, Color.SteelBlue, Color.SeaGreen, Color.DarkOrange, Color.MediumPurple]


def _parse_float(text):
    """
    Fait : convertit un texte de cellule Mechanical en float, virgule ou point decimal.
    Depend de : rien (traitement de chaine pur).
    Retourne : float, ou None si non convertible.
    """
    # Mechanical tourne en locale francaise et renvoie des valeurs type "1,234E-05" dans le Tabular Data pane.
    if text is None:
        return None
    try:
        return float(text.strip().replace(",", "."))
    except ValueError:
        return None


def export_current_view_image(image_name):
    """
    Fait : exporte la vue graphique actuelle en PNG (logo Ansys masque) avec les parametres par defaut du rapport.
    Depend de : ExtAPI.Graphics.ExportImage/ViewOptions (API Ansys), get_unique_file_path (00_constants.py).
    Retourne : str, le chemin du fichier PNG ecrit.
    """
    # ShowLogo=False force ici (point de passage commun a tous les exports d'image) plutot que
    # dans chaque fonction d'export haut niveau : garantit qu'aucune image du rapport n'affiche
    # le logo Ansys, sans devoir y penser a chaque nouvel appelant.
    ExtAPI.Graphics.ViewOptions.ShowLogo = False

    settings = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
    settings.CurrentGraphicsDisplay = False
    settings.Background = GraphicsBackgroundType.White
    settings.Width = DEFAULT_IMAGE_WIDTH
    settings.Height = DEFAULT_IMAGE_HEIGHT

    image_path = get_unique_file_path(IMAGE_EXPORT_FOLDER, image_name, ".png")
    ExtAPI.Graphics.ExportImage(image_path, GraphicsImageExportFormat.PNG, settings)
    return image_path


def set_material_display():
    """
    Fait : bascule la vue graphique en coloration par materiau, maillage masque.
    Depend de : ExtAPI.Graphics.ViewOptions (API Ansys).
    Retourne : rien (effet de bord sur l'etat d'affichage).
    """
    ExtAPI.Graphics.ViewOptions.ModelColoring = ModelColoring.ByMaterial
    ExtAPI.Graphics.ViewOptions.ShowMesh = False


def export_geometry_image():
    """
    Fait : exporte la vue de la geometrie, coloree par materiau.
    Depend de : set_material_display, ExtAPI.DataModel.Project.Model.Geometry, export_current_view_image.
    Retourne : str, le chemin du PNG.
    """
    set_material_display()
    ExtAPI.DataModel.Project.Model.Geometry.Activate()
    return export_current_view_image("geometry")


def export_mesh_image():
    """
    Fait : exporte la vue du maillage.
    Depend de : set_material_display, ExtAPI.DataModel.Project.Model.Mesh, export_current_view_image.
    Retourne : str, le chemin du PNG.
    """
    set_material_display()
    ExtAPI.DataModel.Project.Model.Mesh.Activate()
    return export_current_view_image("mesh")


def export_analysis_overview_image(analysis=None):
    """
    Fait : exporte la vue d'ensemble annotee (BC A, B, C...) affichee quand on selectionne la racine de l'analyse dans l'arbre.
    Depend de : export_current_view_image ; analysis ou, si None, ExtAPI.DataModel.Project.Model.Analyses[0].
    Retourne : str, le chemin du PNG.
    """
    # Reste generique quel que soit le type d'analyse en activant l'analyse fournie plutot qu'un nom code en dur.
    analysis = analysis or ExtAPI.DataModel.Project.Model.Analyses[0]
    analysis.Activate()
    return export_current_view_image(analysis.Name)


def export_object_image(obj, image_name):
    """
    Fait : active un objet et exporte son image via un "Figure" snapshot (plus fiable qu'un export direct en usage repete).
    Depend de : obj.AddFigure() si disponible, sinon repli sur Activate() ; export_current_view_image.
    Retourne : str, le chemin du PNG.
    """
    # Pas de SetFit() ici : le cadrage de la camera (vue choisie via apply_view_if_exists, ou
    # position manuelle courante) est laisse tel quel, a la responsabilite de l'utilisateur - un
    # SetFit() ecraserait silencieusement toute vue personnalisee juste avant la capture.
    add_figure = getattr(obj, "AddFigure", None)
    if add_figure is not None:
        try:
            figure = add_figure()
            figure.Activate()
            image_path = export_current_view_image(image_name)
            obj.Activate()  # restaure l'etat de l'arbre pour les appels suivants (ex: extraction du tabular data)
            return image_path
        except Exception as e:
            print "AddFigure() a echoue pour {} ({}) : export direct utilise.".format(obj.Name, str(e))

    obj.Activate()
    return export_current_view_image(image_name)


def export_solution_image(result):
    """
    Fait : exporte la vue d'un resultat de solution donne.
    Depend de : export_object_image.
    Retourne : str, le chemin du PNG.
    """
    return export_object_image(result, result.Name)


def _read_chart_data(csv_path):
    """
    Fait : lit un CSV (delimiteur ';') et separe l'eventuelle ligne d'en-tete des lignes de donnees numeriques.
    Depend de : le module csv, _parse_float.
    Retourne : tuple (headers, rows) - headers est une liste de noms de colonnes (None si absente), rows une liste de lignes de float.
    """
    # La 1ere ligne est consideree comme un en-tete si elle n'est pas entierement convertible en nombres.
    raw_rows = []
    with open(csv_path, "rb") as f:
        reader = csv.reader(f, delimiter=';')
        for row in reader:
            if row:
                raw_rows.append(row)

    if not raw_rows:
        return None, []

    headers = None
    data_rows = raw_rows
    first_row_values = [_parse_float(cell) for cell in raw_rows[0]]
    if not all(v is not None for v in first_row_values):
        headers = raw_rows[0]
        data_rows = raw_rows[1:]

    rows = []
    for row in data_rows:
        values = [_parse_float(cell) for cell in row]
        if values and all(v is not None for v in values):
            rows.append(values)

    return headers, rows


def export_chart_image_from_csv(csv_path, image_name, chart_title=None, x_axis_label=None,
                                 y_axis_label=None, curve_color=None):
    """
    Fait : construit une image de graphique (titre, axes gradues, grille, courbes + points) a partir d'un CSV exporte par export_active_tabular_data.
    Depend de : _read_chart_data, System.Drawing (Bitmap/Graphics/Pen...), CHART_COLORS.
    Retourne : str, le chemin du PNG genere, ou None si le CSV n'a pas assez de donnees numeriques exploitables.
    """
    # Utilise pour les objets qui n'affichent qu'un graphique 2D dans Mechanical (trackers de Solution Information) : pas de vue 3D a capturer, le graphique est donc redessine depuis les donnees exportees.
    # 1ere colonne = axe X (Time/Step) et chaque colonne suivante = une courbe si le CSV a 2+ colonnes ; sinon le numero de ligne sert d'axe X.
    headers, rows = _read_chart_data(csv_path)
    if len(rows) < 2:
        print "Pas assez de donnees numeriques pour tracer un graphique : " + csv_path
        return None

    num_columns = min(len(r) for r in rows)
    rows = [r[:num_columns] for r in rows]

    if num_columns >= 2:
        x_values = [r[0] for r in rows]
        series = [[r[c] for r in rows] for c in range(1, num_columns)]
    else:
        x_values = [float(i + 1) for i in range(len(rows))]
        series = [[r[0] for r in rows]]

    if headers and len(headers) >= num_columns:
        x_label = headers[0] if num_columns >= 2 else "N"
        series_labels = [headers[c] for c in range(1, num_columns)] if num_columns >= 2 else [headers[0]]
    else:
        x_label = "X"
        series_labels = ["Serie {}".format(i + 1) for i in range(len(series))]

    if x_axis_label:
        x_label = x_axis_label
    if y_axis_label and len(series) == 1:
        series_labels[0] = y_axis_label

    width, height = 900, 550
    plot_left, plot_top = 90, 50
    plot_right, plot_bottom = width - 30, height - 80

    x_min, x_max = min(x_values), max(x_values)
    y_all = [v for serie in series for v in serie]
    y_min, y_max = min(y_all), max(y_all)
    if x_max == x_min:
        x_max += 1.0
    if y_max == y_min:
        y_max += 1.0

    def to_pixel(x, y):
        # Fait : convertit un point (x, y) en donnees vers un pixel de la zone de tracage. Depend de : x_min/x_max/y_min/y_max, plot_left/top/right/bottom (portee englobante). Retourne : PointF.
        px = plot_left + (x - x_min) / (x_max - x_min) * (plot_right - plot_left)
        py = plot_bottom - (y - y_min) / (y_max - y_min) * (plot_bottom - plot_top)
        return PointF(px, py)

    bitmap = Bitmap(width, height)
    g = Graphics.FromImage(bitmap)
    try:
        g.Clear(Color.White)
        g.SmoothingMode = SmoothingMode.AntiAlias

        title_font = Font(FontFamily.GenericSansSerif, 14, FontStyle.Bold)
        axis_title_font = Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold)
        label_font = Font(FontFamily.GenericSansSerif, 9)
        text_brush = SolidBrush(Color.Black)

        title_format = StringFormat()
        title_format.Alignment = StringAlignment.Center
        g.DrawString(chart_title if chart_title else image_name, title_font, text_brush,
                      RectangleF(0, 10, width, 30), title_format)

        # --- Grille fine (mineure, non graduee) ---
        minor_grid_pen = Pen(Color.FromArgb(240, 240, 240), 1)
        minor_divisions = 20
        for i in range(1, minor_divisions):
            t = float(i) / minor_divisions
            py = plot_bottom - t * (plot_bottom - plot_top)
            g.DrawLine(minor_grid_pen, plot_left, py, plot_right, py)
            px = plot_left + t * (plot_right - plot_left)
            g.DrawLine(minor_grid_pen, px, plot_top, px, plot_bottom)

        # --- Grille majeure + graduations (4 divisions sur chaque axe) ---
        major_grid_pen = Pen(Color.Gainsboro, 1)
        divisions = 4
        for i in range(divisions + 1):
            t = float(i) / divisions

            y_val = y_min + t * (y_max - y_min)
            py = plot_bottom - t * (plot_bottom - plot_top)
            g.DrawLine(major_grid_pen, plot_left, py, plot_right, py)
            g.DrawString("{:.3g}".format(y_val), label_font, text_brush, plot_left - 75, py - 7)

            x_val = x_min + t * (x_max - x_min)
            px = plot_left + t * (plot_right - plot_left)
            g.DrawLine(major_grid_pen, px, plot_top, px, plot_bottom)
            g.DrawString("{:.3g}".format(x_val), label_font, text_brush, px - 20, plot_bottom + 8)

        # --- Axes ---
        axis_pen = Pen(Color.Black, 2)
        g.DrawLine(axis_pen, plot_left, plot_bottom, plot_right, plot_bottom)
        g.DrawLine(axis_pen, plot_left, plot_top, plot_left, plot_bottom)

        # --- Titres des axes (noms de colonnes du CSV, ou surcharges x_axis_label/y_axis_label) ---
        x_title_format = StringFormat()
        x_title_format.Alignment = StringAlignment.Center
        g.DrawString(x_label, axis_title_font, text_brush,
                      RectangleF(plot_left, plot_bottom + 28, plot_right - plot_left, 20), x_title_format)

        if len(series) == 1:
            state = g.Save()
            g.TranslateTransform(22, (plot_top + plot_bottom) / 2.0)
            g.RotateTransform(-90)
            y_title_format = StringFormat()
            y_title_format.Alignment = StringAlignment.Center
            g.DrawString(series_labels[0], axis_title_font, text_brush, RectangleF(-100, -15, 200, 20), y_title_format)
            g.Restore(state)

        # --- Courbes (+ points marques, + legende si plusieurs series) ---
        legend_y = plot_top
        for series_index in range(len(series)):
            serie = series[series_index]
            color = curve_color if curve_color is not None else CHART_COLORS[series_index % len(CHART_COLORS)]
            curve_pen = Pen(color, 3)
            marker_brush = SolidBrush(color)

            pixel_points = [to_pixel(x_values[i], serie[i]) for i in range(len(serie))]
            for i in range(len(pixel_points) - 1):
                g.DrawLine(curve_pen, pixel_points[i], pixel_points[i + 1])
            for p in pixel_points:
                g.FillEllipse(marker_brush, p.X - 4, p.Y - 4, 8, 8)

            if len(series) > 1:
                g.FillRectangle(marker_brush, plot_right - 140, legend_y, 12, 12)
                g.DrawString(series_labels[series_index], label_font, text_brush,
                              plot_right - 122, legend_y - 2)
                legend_y += 18

        image_path = get_unique_file_path(IMAGE_EXPORT_FOLDER, image_name, ".png")
        bitmap.Save(image_path, ImageFormat.Png)
    finally:
        g.Dispose()
        bitmap.Dispose()

    return image_path
