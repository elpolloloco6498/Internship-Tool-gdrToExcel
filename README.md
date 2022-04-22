# GDR to Excel

## A propos
Cet outil permet de récupérer les données de l'API GDR sous la forme d'un tableau excel automatiquement.

## Installation de l'outil
Avant d'utiliser cet outil il est nécessaire d'installer Python 3.9.\
De plus il faut installer quelques modules Python. Voici le code permettant d'installer ces modules :
```
pip install openpyxl
```

## Configuration
Il faut maintenant configurer l'outil.
Dans le fichier config.json vous trouverez toutes informations qu'il faut renseigner.
config.json\
```
{
  "api-key": "clé api",
  "gdr": [
    {
      "date": "20210615",
      "startDate": "20220101",
      "endDate": "20221030",
      "siren": "802633693"
    },
    {
      "date": "20210615",
      "startDate": "20220101",
      "endDate": "20221030",
      "cib": "99988"
    },
    {
      "date": "20210615",
      "startDate": "20220101",
      "endDate": "20221030",
      "cib": "34000"
    }
  ]
}
```
Le fichier de configuration ui.json permet de configurer les requêtes vers GDR et la clé de l'API.
ui.json\

## Lancement
Dans un terminal dans le repertoire de l'outil, tapez:\
`python3 main.py`
