from ppt_updater import PPTUpdater

excel_file = r"C:\Users\tariq\Desktop\ppt-automate\tests\sheet_test\test_sheet.xlsm"
ppt_file   = r"C:\Users\tariq\Desktop\ppt-automate\tests\ppt_test\test.pptx"

text_mapping = {
    "ZoneTexte1": {"cell": "C13"},
    "ZoneTexte2": {"cell": "B13"}
}

graph_mapping = {
    "Graphique1": "Graphique1",
    "Tableau1": "A1:D11"
}

with PPTUpdater(excel_file, ppt_file, sheet_name="Feuil1") as updater:
    updater.update_text_shapes(slide_index=1, mapping=text_mapping)
    updater.update_graphs_tables(slide_index=1, mapping=graph_mapping)
