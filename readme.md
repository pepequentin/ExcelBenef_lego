# Bot Idéalo - Calculateur de Bénéfices LEGO

## Description

Le **Bot Idéalo** est un programme conçu pour automatiser le suivi des prix des ensembles LEGO sur le site Idéalo. Le programme permet de calculer le bénéfice potentiel pour chaque ensemble LEGO, en comparant le prix d'achat initial avec le prix actuel idéalo trouvé sur le site Idéalo. De plus, le programme calcule également le coût total des achats, le montant total des ventes réalisées et le potentiel bénéficiaire global.

## Fonctionnalités

- **Suivi Automatisé des Prix**: Le bot parcourt les liens vers les ensembles LEGO spécifiés dans un fichier Excel, récupère les prix actuels idéalo à partir du site Idéalo et les compare aux prix d'achat initiaux pour calculer les bénéfices potentiels.

- **Formatage et Couleurs**: Le bot applique un formatage des couleurs dans le fichier Excel résultant en fonction des valeurs de bénéfice potentiel. Les lignes avec des bénéfices positifs sont mises en <span style="color:green">vert</span>, celles avec des bénéfices négatifs en <span style="color:red">rouge</span>, et celles sans bénéfice en <span style="color:grey">gris</span>.

- **Calculs de Total**: Le programme calcule automatiquement le coût total des achats, le montant total des ventes réalisées et le potentiel bénéficiaire global à partir des données fournies.

## Utilisation

1. Assurez-vous d'avoir toutes les dépendances nécessaires installées. Vous pouvez installer les dépendances en exécutant `pip install -r module_to_install.md`.

2. Placez les liens vers les ensembles LEGO, les prix d'achat initiaux et les quantités dans un fichier Excel nommé `Achat_lego.xlsx`.

3. Exécutez le script `get_idealo_price.py` pour lancer le bot. Celui-ci effectuera automatiquement le suivi des prix, calculera les bénéfices potentiels et mettra à jour le fichier Excel avec les informations nécessaires.

4. Ouvrez le fichier Excel `Achat_lego_updated.xlsx` pour consulter les résultats. Les lignes seront colorées en fonction des bénéfices potentiels et les totaux des coûts, des ventes et des bénéfices seront calculés.

## Notes

- Assurez-vous d'avoir une connexion Internet active pour que le bot puisse accéder au site Idéalo et récupérer les prix actuels idéalo.

- Veuillez vérifier que les prix et les quantités dans le fichier Excel sont correctement formatés pour éviter les erreurs de calcul.

- Chose à coder : 
    1.  Prendre le prix de vente `si et seulement si la case du excel est dif de NAN` au lieu du prix idéalo
    2.  Il est possible que le lego soit assez rare pour qu'il n'apparaisse pas dans idéalo, donc on doit prendre le prix ailleur `à def`.
    3.  Il est possible que le logi soit assez rare pour qu'il n'apparaisse quand état d'occasion, donc on doit prendre le prix ailleur `à def`.
    4.  Garder les couleurs de l'input file sur la colonne `Stoké`
    5.  Bot discord pour les notif de nouvelle annonce vinted