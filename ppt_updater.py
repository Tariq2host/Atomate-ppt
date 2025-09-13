import win32com.client as win32
import time

class PPTUpdater:
    def __init__(self, excel_path, ppt_path, sheet_name=None, visible_excel=False, visible_ppt=True):
        """
        excel_path : chemin vers le fichier Excel (.xlsm)
        ppt_path   : chemin vers le fichier PowerPoint
        sheet_name : nom de la feuille Excel contenant les données (optionnel)
        visible_excel : bool, affiche Excel pendant l'exécution
        visible_ppt   : bool, affiche PowerPoint pendant l'exécution
        """
        self.excel_path = excel_path
        self.ppt_path = ppt_path
        self.sheet_name = sheet_name
        self.visible_excel = visible_excel
        self.visible_ppt = visible_ppt
        self.excel = None
        self.wb = None
        self.sheet = None
        self.ppt = None
        self.presentation = None

    def __enter__(self):
        # Lancer Excel
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = self.visible_excel
        self.wb = self.excel.Workbooks.Open(self.excel_path)

        # Déterminer la feuille à utiliser
        if self.sheet_name:
            self.sheet = self.wb.Sheets(self.sheet_name)
        else:
            self.sheet = self.wb.Sheets(1)  # première feuille par défaut

        # Lancer PowerPoint
        self.ppt = win32.Dispatch("PowerPoint.Application")
        self.ppt.Visible = self.visible_ppt
        self.presentation = self.ppt.Presentations.Open(self.ppt_path)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.presentation:
            self.presentation.Save()
            self.presentation.Close()
        if self.wb:
            self.wb.Close(SaveChanges=False)
        if self.excel:
            self.excel.Quit()

    def _refresh_slide(self, slide_index):
        """
        Force PowerPoint à afficher correctement les mises à jour sur la slide.
        """
        try:
            self.presentation.Windows(1).View.GotoSlide(slide_index)
            self.presentation.Windows(1).Activate()
            time.sleep(0.1)  # petite pause pour laisser PowerPoint traiter le rendu
        except Exception:
            pass

    def update_text_shapes(self, slide_index, mapping):
        """
        mapping: dict
            clé = nom de la zone de texte dans PPT
            valeur = {"cell": "A1"}
        """
        slide = self.presentation.Slides(slide_index)
        for ppt_name, info in mapping.items():
            cell_address = info["cell"]

            try:
                # Exécuter la macro Excel pour copier la cellule avec mise en forme
                self.excel.Application.Run("CopyFormattedCell", self.sheet.Name, cell_address)
            except Exception as e:
                print(f"❌ Erreur macro Excel sur {cell_address}: {e}")
                continue

            shape = next((s for s in slide.Shapes if s.Name == ppt_name), None)
            if shape:
                try:
                    shape.TextFrame.TextRange.Text = ""  # Vider texte existant
                    shape.TextFrame.TextRange.Paste()    # Coller avec format source
                except Exception as e:
                    print(f"⚠ Impossible de coller dans '{ppt_name}': {e}")
            else:
                print(f"⚠ Forme '{ppt_name}' introuvable sur slide {slide_index}")

        self._refresh_slide(slide_index)

    def update_graphs_tables(self, slide_index, mapping):
        """
        mapping: dict
            clé = nom de l'objet dans PPT
            valeur = nom du graphique ou de la plage Excel (ChartObject ou Range)
        """
        slide = self.presentation.Slides(slide_index)
        for ppt_name, excel_name in mapping.items():
            old_shape = next((s for s in slide.Shapes if s.Name == ppt_name), None)

            if "Graph" in excel_name:
                chart = self.sheet.ChartObjects(excel_name)
                chart.Copy()
                new_shape = slide.Shapes.Paste()[0]
            else:
                # Exécuter la macro Excel pour copier la plage avec format
                self.excel.Application.Run("CopyFormattedRange", self.sheet.Name, excel_name)
                shape_range = slide.Shapes.PasteSpecial(DataType=8)  # OLE Object
                new_shape = shape_range.Item(1)

            # Conserver position et taille
            if old_shape:
                new_shape.Left = old_shape.Left
                new_shape.Top = old_shape.Top
                new_shape.Width = old_shape.Width
                new_shape.Height = old_shape.Height
                old_shape.Delete()

            new_shape.Name = ppt_name

        self._refresh_slide(slide_index)
