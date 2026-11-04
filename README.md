# OMAS Internat : choix du menu

Formulaire en ligne pour recueillir les choix de menu des participants à un événement OMAS Internat.
La page est hébergée sur GitHub Pages, et les réponses arrivent dans un Google Sheet.

## Mise en route (une seule fois, 10 minutes)

1. Créer un Google Sheet vide, puis ouvrir **Extensions > Apps Script**.
2. Remplacer le contenu de `Code.gs` par celui de [`apps-script/Code.gs`](apps-script/Code.gs), puis enregistrer.
3. Choisir la fonction `installer` et cliquer sur **Exécuter**. Accepter les autorisations demandées par Google.
   Les onglets « Réglages », « Menu », « Réponses » et « Récap » apparaissent dans le tableau.
4. **Déployer > Nouveau déploiement > Application Web**, avec « Exécuter en tant que : Moi » et « Qui a accès : Tout le monde ».
5. Copier l'adresse obtenue (elle se termine par `/exec`) et la coller dans `index.html`, à la ligne `const API = "";`.

Tant que cette ligne est vide, la page fonctionne en aperçu avec un menu d'exemple et n'enregistre rien.

## Au quotidien

Tout se règle dans le Google Sheet, sans toucher au code :

- **Réglages** : nom de l'événement, date limite, `Choix ouverts` (mettre `NON` pour clore), `Demander les allergies`.
- **Menu** : une ligne par plat, avec l'étape (Entrée, Plat, Dessert, Boisson…), le nom du plat et une description facultative.
- **Réponses** : une ligne par personne. Si quelqu'un répond deux fois avec le même nom, sa ligne est remplacée.
- **Récap** : le nombre de chaque plat et la liste des allergies, à transmettre au restaurant.

Après une modification du script (pas du tableau), il faut publier une nouvelle version : **Déployer > Gérer les déploiements > Modifier > Nouvelle version**. L'adresse ne change pas.

## Données personnelles

Les allergies sont des données de santé. Seuls les organisateurs ont accès au tableau. Supprimer les réponses après l'événement.
