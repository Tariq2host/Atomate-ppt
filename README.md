
# Projet PPT Automate

Automatisation de la mise à jour de fichiers PowerPoint à partir de données Excel.  
Ce projet permet de :  

- Mettre à jour des **zones de texte** dans PowerPoint avec les valeurs et formats Excel.  
- Mettre à jour des **graphiques et tableaux** en conservant le format d’Excel via des macros.  
- Gérer plusieurs slides et feuilles Excel de manière flexible et réutilisable.

---

## 📁 Structure du projet

```

draft/
tests/
├── ppt\_test/
│   └── test.pptx
├── sheet\_test/
│   ├── test\_sheet.xlsm   # Excel macro activé
│   └── test\_sheet.xlsx
└── test\_ppt.ipynb
.gitignore
LICENCE
macro.bash               # Script pour installer ou exécuter les macros si nécessaire
main.py                  # Exemple d'exécution du module
ppt\_updater.py           # Module principal pour mettre à jour PPT depuis Excel
README.md
requirements.txt

````

---

## ⚙️ Installation

1. Cloner le projet :

```bash
git clone <votre-repo-url>
cd <nom-du-projet>
````

2. Installer les dépendances :

```bash
pip install -r requirements.txt
```

> Le projet utilise principalement `pywin32` pour automatiser Excel et PowerPoint.

3. Assurez-vous que les fichiers Excel sont **.xlsm** pour activer les macros et que les macros suivantes sont présentes :

```vb
Sub CopyFormattedRange(sheetName As String, rngName As String)
    Dim rng As Range
    On Error GoTo ErrHandler
    
    ' Référence la feuille passée en paramètre
    Set rng = ThisWorkbook.Sheets(sheetName).Range(rngName)
    
    ' Copier valeurs + formats
    rng.Copy
    Exit Sub

ErrHandler:
    MsgBox "Erreur: impossible de copier la plage " & rngName & " sur la feuille " & sheetName, vbCritical
End Sub


Sub CopyFormattedCell(sheetName As String, cellAddress As String)
    Dim rng As Range
    On Error GoTo ErrHandler
    
    ' Référence la feuille passée en paramètre
    Set rng = ThisWorkbook.Sheets(sheetName).Range(cellAddress)
    
    ' Copier valeur + format
    rng.Copy
    Exit Sub

ErrHandler:
    MsgBox "Erreur: impossible de copier la cellule " & cellAddress & " sur la feuille " & sheetName, vbCritical
End Sub

```



* `CopyFormattedCell(sheetName, cellAddress)` : copie une cellule avec son format

* `CopyFormattedRange(sheetName, rngName)` : copie une plage avec son format

---

## 📝 Utilisation

### Exemple d'utilisation avec le module `ppt_updater.py` :

```python
from ppt_updater import PPTUpdater

excel_file = r"tests/sheet_test/test_sheet.xlsm"
ppt_file   = r"tests/ppt_test/test.pptx"

# Mapping des zones de texte
text_mapping = {
    "ZoneTexte1": {"cell": "C13"},
    "ZoneTexte2": {"cell": "B13"}
}

# Mapping des graphiques et tableaux
graph_mapping = {
    "Graphique1": "Graphique1",
    "Tableau1": "A1:D11"
}

# Mettre à jour la slide 1
with PPTUpdater(excel_file, ppt_file, sheet_name="Données") as updater:
    updater.update_text_shapes(slide_index=1, mapping=text_mapping)
    updater.update_graphs_tables(slide_index=1, mapping=graph_mapping)
```

> `slide_index` commence à **1** (PowerPoint est 1-indexé).
> `sheet_name` est le nom de la feuille Excel à utiliser. Si non précisé, la première feuille est utilisée.

---

## 🛠 Fonctionnalités principales

* Mise à jour **texte** et **format** depuis Excel.
* Mise à jour **graphique** et **tableau** en conservant la taille et position d’origine dans PPT.
* Compatible avec plusieurs slides et feuilles Excel.
* Gestion automatique de l’ouverture et fermeture d’Excel/PowerPoint.

---

## 📦 Dépendances

* Python >= 3.10
* [pywin32](https://pypi.org/project/pywin32/) >= 305

```
pip install pywin32>=305
```

---

## 💡 Remarques

* Le fichier Excel doit être **macro-enabled** (.xlsm) pour que les macros fonctionnent.
* Les macros doivent être présentes dans le fichier Excel ou dans un module séparé pour copier les cellules/plages avec leur format.
* Le projet fonctionne uniquement sous **Windows** avec Microsoft Office installé.

---

## 🔖 Licence

Ce projet est sous licence **MIT**. Voir le fichier [LICENCE](./LICENCE).

