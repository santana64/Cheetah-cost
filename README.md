# Cheetah Cost - Version 0.1.1

**Cheetah Cost** est une application de gestion des coûts de projet, conçue pour permettre une gestion et une visualisation simplifiées des coûts, des dépenses réelles, et de la valeur acquise en utilisant la méthode des **courbes en S**. Elle propose également des outils pour générer des rapports et des graphiques basés sur les données de coûts.

## Fonctionnalités

- Gestion des coûts de projet : coûts estimés, coûts réels, reste à faire.
- Support de la méthode des courbes en S pour visualiser les coûts et les écarts.
- Calcul automatique des indicateurs clés : CPI, SPI.
- Exportation des données vers Excel et Word.
- Interface utilisateur moderne et intuitive basée sur **PyQt5**.
- Prise en charge de l’ajout de périodes supplémentaires.
- Génération de graphiques en temps réel.

## Installation

### Pré-requis

1. **Python 3.8+** installé sur votre machine.
2. Les bibliothèques suivantes doivent être installées :
   - `PyQt5`
   - `matplotlib`
   - `pandas`
   - `python-docx`
   - `qt_material`

Vous pouvez installer ces bibliothèques à l'aide de `pip`.

### Étapes d'installation

1. Clonez le dépôt Git sur votre machine locale :

   
   git clone https://github.com/santana64/cheetah-cost.git
   cd cheetah-cost



Installez les dépendances requises via pip :

pip install PyQt5 matplotlib pandas python-docx qt_material
Utilisation

Lancer l'application

Après avoir lancé l'application, vous serez accueilli par l'écran d'accueil Cheetah Cost. Cliquez sur le bouton Commencer pour créer un nouveau projet.
Création d'un projet

Remplissez les informations du projet, telles que le nom, la description, les dates de début et de fin, ainsi que les coûts estimés. Vous pouvez également choisir d'ajouter des périodes supplémentaires si le projet s'étend au-delà de la période prévue.
Suivi et Analyse des Coûts

Une fois le projet créé, vous accédez au tableau de suivi des coûts. Entrez les coûts réels et utilisez les boutons pour calculer les indicateurs de performance (CPI, SPI) ou générer un graphique des courbes en S.
Exportation des Données

Utilisez les options d'exportation pour enregistrer vos données dans un fichier Excel ou Word. Cela vous permet de partager ou de sauvegarder vos analyses.
Contribution

Les contributions sont les bienvenues ! Si vous souhaitez contribuer à Cheetah Cost, veuillez :

Forker le projet.
Créer une nouvelle branche pour vos modifications.
Soumettre une pull request.
Licence
Cheetah Cost est distribué sous la licence MIT. Veuillez consulter le fichier LICENSE pour plus d’informations.

Problèmes connus
L’implémentation de l'annulation/refaire des actions n’est pas encore disponible.
Des ajustements peuvent être nécessaires pour l’utilisation sur différentes résolutions d’écran.
