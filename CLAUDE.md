# Korecto : consignes de travail

## Rythme de publication
- Une mise à jour par semaine : les corrections s'accumulent sur la branche de travail, la release est décidée par Maxime après usage.
- Ne jamais publier (étiquette `vX.Y.Z`, mise sur `main` pour une release) sans demande explicite. Avant la release : numéro de version dans `src/AppSettings.cs`, ligne en tête de `src/NOUVEAUTES.txt`, nombre de contrôles dans `README.md` et `.github/workflows/tests.yml`.
- Les tests Windows se lancent par GitHub Actions (`tests.yml`, workflow_dispatch) : on ne peut pas compiler sur Linux.

## Liste des corrections en attente
Ne pas la répéter à chaque message. La donner seulement quand Maxime la demande, ou quand une correction est ajoutée ou que la release se prépare (une ligne par correction : description, commit, résultat des tests Windows).

## Corrections en attente de release (prochaine : 1.0.5)
Copie indépendante préparée par Codex le 29 octobre : dictionnaire métier sans collecte de noms de clients, proposition d'objet à copier, correction limitée à la sélection Outlook, consigne Plus court renforcée. Relu et corrigé par Claude (tiret dans la validation de l’objet, message de sélection refusée, écriture atomique du dictionnaire, ligne keptKey retirée) ; compilation locale Windows et 320 contrôles locaux ; lecture/application COM dans Outlook réel à valider. Pas de publication. Voir POUR_INTEGRATION.md dans la copie de travail.
