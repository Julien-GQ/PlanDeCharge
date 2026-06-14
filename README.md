# Mise en forme charge

Projet Python de base pour manipuler et formater des valeurs de charge.

## Objectif

Ce depot fournit:
- une structure de projet moderne basee sur `src/`
- un package Python `mise_en_forme_charge`
- une CLI minimale pour executer un exemple
- des tests unitaires avec `pytest`
- un socle qualite avec `ruff` et `black`

## Prerequis

- Python 3.10 ou plus
- PowerShell (Windows)

## Installation

1. Creer l'environnement virtuel:

```powershell
python -m venv .venv
```

2. Activer l'environnement:

```powershell
.\.venv\Scripts\Activate.ps1
```

3. Installer les dependances de developpement:

```powershell
pip install -e .[dev]
```

## Utilisation

Lancer la CLI du projet:

```powershell
mise-en-forme-charge
```

Lancer sans installation de script (via module):

```powershell
python -m mise_en_forme_charge.cli
```

## Developpement

Executer les tests:

```powershell
pytest -q
```

Verifier le style et la qualite:

```powershell
ruff check .
black --check .
```

## Structure du projet

```text
.
|-- pyproject.toml
|-- README.md
|-- src/
|   `-- mise_en_forme_charge/
|       |-- __init__.py
|       |-- cli.py
|       `-- formatter.py
`-- tests/
    `-- test_formatter.py
```

## Prochaine etape

Ajouter les premieres regles metier de mise en forme dans `src/mise_en_forme_charge/formatter.py`, puis completer les tests dans `tests/test_formatter.py`.


#### note Julien

LECTURE / Le projet :

Le projet a pour objectif de recuperer un fichier soit .xlsm ou .txt.
Ce fichier aura toujours la même structure car il résulte d'une requete effectuer sur la base de données entreprise. (A terme il sera eventuellement une tache planifié.)
Le fichier contient des ordre de fabrication en cours avec different site de production, avec le numéro d'OF, le nom du client, la date de livraison, le code article et la quantité a livrer.... l'objectif est de faire un tableau des commande en cours a la semaine connaitre les retard et connaitre les quantité à livrer a la semaine sur les differentes famille d'article.

LECTURE / programmation :

Je voudrais externaliser le maximum de fonction pour epurer les pages .py 

LECTURE / Les evolutions :

La 1ere étape serait de creer un programme python qui sache lire les fichiers et mettre en page comme souhaité le fichier sous format excel.

La 2eme etape serait de faire une programme qui mette en page directement dans le programme python via une interface en tkinter par exemple.

Lé 3eme etape sera de creer une appli en flutter et dark pour synchroniser les donnees et les posseder en direct dans mon telephone.

ACTION / lire les fichiers :

le fichier excel ou txt est disponible dans le dossier "base". Peux tu les analyser et ecrire ci dessous ce que reussit a ressortir .

RESULTAT / analyse des fichiers du dossier base

Les 2 fichiers sont bien exploitables comme source de donnees tabulaires.

1. Fichier texte analyse

- Fichier lu : DELAI_FAB_MJULIEN1.txt
- Format : texte tabule
- Nombre de lignes de donnees : 1259
- Nombre de colonnes : 17

Colonnes detectees :

- DATECDE
- DT_LIV_CONFIRMEE
- COMMANDE
- NOMREDCLI_CDE
- CODE2
- DESIGNATION
- RESTE_A_LIV_UV
- MOR_ATELIER
- CAPREV_LCDE
- PXVENTE
- FAM_VENTE
- MOR_OF
- OF_LANCE
- MOR_60EMB
- OF_SOLDFAB
- CPAYS_CDE
- ARTICLE

Ce que le programme pourra deja ressortir depuis ce fichier :

- les commandes en cours
- le numero de commande / OF
- le client
- la date de commande
- la date de livraison confirmee
- le code article et la designation
- la quantite restante a livrer
- le site de production
- la famille de vente
- le pays client
- les montants / valorisations presents dans CAPREV_LCDE et PXVENTE

Premiers constats sur les donnees texte :

- 549 commandes uniques detectees
- 228 lignes n'ont pas de date de livraison confirmee
- plage de dates de livraison confirmee detectee : du 2025-03-20 au 2028-12-31
- principaux sites de production : ANNOEULL (450), PROVIN (426), MWF (245), WILLEFER (83)
- principales familles de vente : VTEGAZ (534), VTELIQ (338), VTENEGG (170), VTEAIR (99), VTEMSNOR (89)
- principaux pays clients : FRA (1130), BEL (24), MAR (14), ESP (12), DZA (12), TUR (12)

Points d'attention detectes sur le fichier texte :

- certaines valeurs ont des caracteres mal encodes, par exemple sur les accents ou le symbole degre
- certaines colonnes peuvent etre vides, notamment DT_LIV_CONFIRMEE et parfois MOR_ATELIER

2. Fichier Excel analyse

- Fichier lu : DELAI_FAB_MJULIEN.xlsm
- Feuille detectee : DELAI_FAB_MJULIEN
- Table detectee : A1:V1436
- Nombre de lignes de donnees : 1435
- Nombre de colonnes : 22

Colonnes detectees dans le fichier Excel :

- DATECDE
- DT_LIV_CONFIRMEE
- COMMANDE
- NOMREDCLI_CDE
- CODE2
- DESIGNATION
- RESTE_A_LIV_UV
- MOR_ATELIER
- CAPREV_LCDE
- PXVENTE
- FAM_VENTE
- MOR_OF
- OF_LANCE
- MOR_10PREC
- MOR_20COUPE
- MOR_30RENF
- MOR_40CONF
- MOR_50RETCOU
- MOR_60EMB
- OF_SOLDFAB
- CPAYS_CDE
- ARTICLE

Conclusion sur la source Excel :

- le fichier Excel contient les colonnes du fichier texte
- il contient aussi des colonnes supplementaires de suivi de fabrication
- il semble donc plus riche pour suivre l'avancement atelier et construire un tableau de suivi hebdomadaire

3. Ce qu'il sera possible de produire dans la premiere version du programme

- importer un fichier .txt ou .xlsm depuis le dossier base
- normaliser les colonnes utiles
- filtrer les commandes en cours
- regrouper les quantites a livrer par semaine
- regrouper les donnees par site de production
- regrouper les donnees par famille d'article / famille de vente
- identifier les lignes sans date de livraison
- preparer un export Excel mis en forme pour le suivi hebdomadaire

4. Recommandation pour la suite

Pour la V1, la meilleure base semble etre le fichier Excel .xlsm car il contient plus d'informations de suivi de production. Le fichier .txt reste une bonne source de secours ou un format d'import simplifie.
