# 03_ppt_utils.py : constructeur de rapport PowerPoint - possede la session COM Interop et expose des methodes "add slide" basees sur les layouts du template corporate. Depend de 00_constants.py (doit etre execute avant ce fichier).

import clr
import csv
import datetime
import shutil

clr.AddReference("Microsoft.Office.Interop.PowerPoint")
clr.AddReference("Office")

import Microsoft.Office.Interop.PowerPoint as PPT
import Microsoft.Office.Core as Office


def _build_working_copy_base_name():
    """
    Fait : construit un nom de base (sans extension), unique par jour.
    Depend de : datetime.date.today().
    Retourne : str, ex. "automatic_report_generation_17072025".
    """
    today = datetime.date.today()
    return "automatic_report_generation_{:02d}{:02d}{:04d}".format(today.day, today.month, today.year)


def rename_time_header_to_loadcase(data):
    """
    Fait : remplace, dans la ligne d'en-tete d'un tableau CSV, toute cellule valant exactement "Time [s]" par "Loadcase".
    Depend de : rien (modifie data en place).
    Retourne : rien (effet de bord sur data).
    """
    # Plus parlant pour un ingenieur que le nom brut de colonne renvoye par le Tabular Data pane de Mechanical.
    if not data:
        return
    header = data[0]
    for i in range(len(header)):
        if header[i] == u"Time [s]":
            header[i] = u"Loadcase"


class PPTReportBuilder(object):
    """
    Encapsule une unique session PowerPoint Interop ouverte sur une copie de travail du template corporate. Toutes les methodes create_..._slide ajoutent des slides a LA MEME presentation au lieu de rouvrir PowerPoint a chaque slide.
    """

    def __init__(self, template_path):
        """
        Fait : cree une copie de travail du template et ouvre une session PowerPoint COM dessus.
        Depend de : shutil.copyfile, get_unique_file_path (00_constants.py), Microsoft.Office.Interop.PowerPoint.
        Retourne : rien (initialise self.app / self.presentation / self.working_copy_path).
        """
        # Le template original n'est JAMAIS ouvert directement : si l'utilisateur fait Ctrl+S dans PowerPoint pendant la generation, c'est cette copie qui est ecrasee, jamais l'original.
        self.working_copy_path = get_unique_file_path(
            REPORT_OUTPUT_FOLDER, _build_working_copy_base_name(), ".pptx")
        shutil.copyfile(template_path, self.working_copy_path)
        print "Template working copy opened: " + self.working_copy_path

        self.app = PPT.ApplicationClass()
        # self.app.Visible = True est necessaire : une session laissee invisible s'est reveleee instable sur un rapport avec beaucoup de slides (COMException "Presentation.SlideMaster : Object does not exist" en cours de route, plus aucune slide ne peut alors etre ajoutee/sauvegardee). La fenetre se ferme normalement en fin de generation (voir close()).
        self.app.Visible = True
        self.presentation = self.app.Presentations.Open(self.working_copy_path, WithWindow=True)

    def _add_slide(self, layout_index):
        """
        Fait : ajoute une slide vierge a la fin de la presentation, sur le layout personnalise donne.
        Depend de : self.presentation.SlideMaster.CustomLayouts.
        Retourne : PPT.Slide, la slide creee.
        """
        layout = self.presentation.SlideMaster.CustomLayouts[layout_index]
        return self.presentation.Slides.AddSlide(self.presentation.Slides.Count + 1, layout)

    def add_image_table_slide(self, title, subtitle, img_path=None, csv_path=None, comment=" "):
        """
        Fait : ajoute une slide "image + table" (titre, sous-titre, une image, une table, un commentaire).
        Depend de : self._add_slide (LAYOUT_IMAGE_TABLE), self.add_csv_table.
        Retourne : PPT.Slide, la slide creee.
        """
        slide = self._add_slide(LAYOUT_IMAGE_TABLE)

        # Affectation du texte APRES la creation de la slide : le faire sur le layout directement modifierait tout le master template.
        slide.Shapes[8].TextFrame.TextRange.Text = comment
        slide.Shapes[2].TextFrame.TextRange.Text = title
        slide.Shapes[4].TextFrame.TextRange.Text = subtitle

        if img_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[3]
            try:
                slide.Shapes.AddPicture(img_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                         coord.Left, coord.Top, coord.Width, coord.Height)
            except Exception as e:
                print "Unable to insert image ({}): {}".format(img_path, str(e))

        if csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[1]
            try:
                self.add_csv_table(slide, csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(csv_path, str(e))

        return slide

    def add_table_slide(self, title, subtitle, csv_path):
        """
        Fait : ajoute une slide "table seule" (titre, sous-titre, table).
        Depend de : self._add_slide (LAYOUT_TABLE_ONLY), self.add_csv_table.
        Retourne : PPT.Slide, la slide creee.
        """
        slide = self._add_slide(LAYOUT_TABLE_ONLY)
        slide.Shapes[1].TextFrame.TextRange.Text = title
        slide.Shapes[3].TextFrame.TextRange.Text = subtitle

        coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_TABLE_ONLY].Shapes[2]
        try:
            self.add_csv_table(slide, csv_path, coord.Left, coord.Top, coord.Width)
        except Exception as e:
            print "Unable to insert table ({}): {}".format(csv_path, str(e))
        return slide

    def add_analysis_context_slide(self, title, subtitle, img_path, settings_csv_path, solution_csv_path):
        """
        Fait : ajoute la slide de contexte "Analysis Parameters" (image de vue d'ensemble + 2 tableaux : parametres de steps, infos de resolution).
        Depend de : self._add_slide (LAYOUT_IMAGE_TABLE), self.add_csv_table.
        Retourne : PPT.Slide, la slide creee.
        """
        # Reutilise le layout LAYOUT_IMAGE_TABLE : son emplacement "table" (shape 1) recoit le tableau
        # des steps (meme position que sur les autres slides image+table), et son emplacement
        # "commentaire" (shape 8, plus reduit) recoit le second tableau (infos de resolution, 3 lignes max).
        slide = self._add_slide(LAYOUT_IMAGE_TABLE)

        slide.Shapes[2].TextFrame.TextRange.Text = title
        slide.Shapes[4].TextFrame.TextRange.Text = subtitle

        if img_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[3]
            try:
                slide.Shapes.AddPicture(img_path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                         coord.Left, coord.Top, coord.Width, coord.Height)
            except Exception as e:
                print "Unable to insert image ({}): {}".format(img_path, str(e))

        if settings_csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[1]
            try:
                self.add_csv_table(slide, settings_csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(settings_csv_path, str(e))

        if solution_csv_path:
            coord = self.presentation.SlideMaster.CustomLayouts[LAYOUT_IMAGE_TABLE].Shapes[8]
            try:
                self.add_csv_table(slide, solution_csv_path, coord.Left, coord.Top, coord.Width)
            except Exception as e:
                print "Unable to insert table ({}): {}".format(solution_csv_path, str(e))

        return slide

    def add_csv_table(self, slide, csv_path, left, top, width):
        """
        Fait : lit un CSV (delimiteur ';') et l'insere en table formatee sur la slide (en-tete gras/grise, bordures fines, texte centre).
        Depend de : le module csv, rename_time_header_to_loadcase, MAX_TABLE_ROWS/MAX_TABLE_COLUMNS (00_constants.py).
        Retourne : rien (modifie slide ; ne fait rien si le CSV est vide ou depasse les limites de taille).
        """
        data = []
        # Lecture binaire + decodage UTF-8 explicite : les CSV sont ecrits en UTF-8 (unites avec caracteres speciaux type degre/micro/exposants) ; une lecture texte sans encodage explicite provoque un DecoderFallbackException.
        with open(csv_path, "rb") as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                data.append([cell.decode("utf-8") for cell in row])

        if not data:
            print "Empty CSV, no table inserted: " + csv_path
            return

        rename_time_header_to_loadcase(data)

        rows = len(data)
        cols = len(data[0])

        if rows > MAX_TABLE_ROWS or cols > MAX_TABLE_COLUMNS:
            print ("The table exceeds 50 rows / 50 columns ({} rows, {} columns): it will "
                   "therefore not be shown in PowerPoint but is available in csv format at "
                   "this location: {}").format(rows, cols, csv_path)
            return

        table = slide.Shapes.AddTable(rows, cols, left, top, width).Table

        # Bordures posees UNE FOIS PAR LIGNE (Rows(r).Cells.Borders accepte une plage de cellules), pas par cellule x par cote : chaque aller-retour COM est couteux, et c'etait la partie la plus lente du formatage (jusqu'a 45s pour 8 lignes avant cette optimisation). Font/TextRange, eux, restent necessairement par cellule (pas d'equivalent en plage).
        for r in range(1, rows + 1):
            row_cells = table.Rows(r).Cells
            for border_index in range(1, 5):
                row_cells.Borders(border_index).ForeColor.RGB = 0x000000
                row_cells.Borders(border_index).Weight = 1

        for r in range(rows):
            for c in range(cols):
                cell = table.Cell(r + 1, c + 1)
                shape = cell.Shape
                text_frame = shape.TextFrame
                text_range = text_frame.TextRange
                text_range.Text = data[r][c]
                text_range.Font.Size = 7
                text_range.ParagraphFormat.Alignment = 2  # 2 = centre
                text_frame.VerticalAnchor = 3  # 3 = milieu
                # Marge haut/bas ramenee a 3pt (marges gauche/droite inchangees) : c'est cette marge interne, pas la taille de police, qui empeche PowerPoint de reduire la hauteur de ligne en dessous d'un certain seuil.
                text_frame.MarginTop = 3
                text_frame.MarginBottom = 3

                if r == 0:
                    text_range.Font.Bold = True
                    shape.Fill.ForeColor.RGB = 0x545454
                    shape.Fill.Solid()

        # Hauteur forcee a une valeur volontairement trop petite : PowerPoint la ramene automatiquement au minimum reel necessaire pour loger le texte - seul moyen de resserrer un tableau deja cree (AddTable alloue par defaut une hauteur bien superieure au necessaire pour du texte en taille 7, ce qui peut faire deborder la slide).
        for r in range(1, rows + 1):
            table.Rows(r).Height = 1

    def save_as(self, output_path):
        """
        Fait : enregistre la presentation dans un nouveau fichier (le template original n'est jamais ecrase).
        Depend de : self.presentation.SaveAs.
        Retourne : rien (effet de bord : cree/ecrase output_path).
        """
        self.presentation.SaveAs(output_path)
        print "Report saved: " + output_path

    def close(self):
        """
        Fait : enregistre la copie de travail, ferme la presentation et quitte l'application PowerPoint.
        Depend de : self.presentation.Save/Close, self.app.Quit.
        Retourne : rien (effet de bord ; appele en cas d'echec de la generation, pour ne jamais laisser une session PowerPoint invisible en memoire).
        """
        self.presentation.Save()  # Save() simple, pas SaveAs : le fichier a deja son nom/chemin final (working_copy_path)
        self.presentation.Close()
        self.app.Quit()

    def keep_open(self):
        """
        Fait : ne fait rien sur la presentation (pas de Save, pas de Close, pas de Quit) - la laisse telle quelle, ouverte et affichee.
        Depend de : rien.
        Retourne : rien (effet de bord : aucun ; existe pour rendre explicite, a la fin d'une generation reussie, le choix de laisser le rapport ouvert sans l'enregistrer).
        """
        pass
