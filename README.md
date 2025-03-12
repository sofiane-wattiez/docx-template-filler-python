# docx-template-filler-python

## Description

### English
This Python script generates `.docx` documents from templates (`template_remise.docx` and `template_restitution.docx`). The user enters information via a graphical user interface (Tkinter), and the script replaces the placeholders in the templates with the provided values. The generated documents are saved into specific folders.

The script creates a main folder based on the user's first and last name, then adds subfolders for each generated document (e.g., `Document_Remise` and `Document_Restitution`).

### Français
Ce script Python génère des documents `.docx` à partir de modèles (`template_remise.docx` et `template_restitution.docx`).
L'utilisateur entre des informations via une interface graphique (Tkinter), et le script remplace les champs de texte dans les modèles avec les valeurs saisies. Les documents générés sont ensuite enregistrés dans des dossiers spécifiques.

Le script crée un dossier principal basé sur le nom et le prénom de l'utilisateur, puis y ajoute des sous-dossiers pour chaque document généré (par exemple, `Document_Remise` et `Document_Restitution`).

---

## Dependencies / Prérequis

- Python 3.12.8 "Development Version"
- The `python-docx` library to handle `.docx` files.  

```bash
pip install python-docx
```

- The `tkinter` library for GUI.
```bash
pip install tkinter
```


### ENGLISH 

**This Python script generates `.docx` documents from templates (`template_remise.docx` and `template_restitution.docx`).**

The user enters information through a graphical user interface (`Tkinter`), and the script replaces the placeholders in the templates with the provided values.

The generated documents are then saved into specific folders.

The script creates a main folder based on the user's first and last name, and then adds subfolders for each generated document (e.g., `Document_Remise` and `Document_Restitution`).

### Français

**Ce script Python génère des documents `.docx` à partir de modèles (`template_remise.docx` et `template_restitution.docx`).**

L'utilisateur entre des informations via une interface graphique (`Tkinter`), et le script remplace les champs de texte dans les modèles avec les valeurs saisies.

Les documents générés sont ensuite enregistrés dans des dossiers spécifiques.

Le script crée un dossier principal basé sur le nom et le prénom de l'utilisateur, puis y ajoute des sous-dossiers pour chaque document généré (par exemple, `Document_Remise` et `Document_Restitution`).

---

## Installation

1. **Clonez le repository :**

```bash
git clone https://github.com/sofiane-wattiez/docx-template-filler-python.git
```

## Usage / Utilisation

2. **Script Execution / Exécutez le script :**

**In the directory of the project execute script like this:**

**Dans le répertoire du projet, exécutez le script Python :**

```bash
python script.py
```

**3. Insert Data / Saisie des informations :**

**Lorsque vous lancez le script, une interface graphique (GUI) s'ouvrira. Vous devrez remplir les champs suivants:**

```bash
Name
Surname
Date (date of day)
Brand
Modele
New ID
New S/N
Old ID (facultatif) # If Value is 0 or null the 2nd docx don't be created
Old S/N (facultatif) # If Value is 0 or null the 2nd docx don't be created
```

## Result 
```bash
Document_MI_John_Doe/
├── Document_Remise/
│   └── Document_Remise_MI_John_Doe_SL_PT_12345.docx
└── Document_Restitution/
    └── Document_Restitution_MI_John_Doe_SL_PT_67890.docx
```

## Author : Sofiane Wattiez
