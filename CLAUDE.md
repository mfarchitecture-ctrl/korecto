# Korecto : consignes de travail

## Rythme de publication
- Une mise à jour par semaine : les corrections s'accumulent sur la branche de travail, la release est décidée par Maxime après usage.
- Ne jamais publier (étiquette `vX.Y.Z`, mise sur `main` pour une release) sans demande explicite. Avant la release : numéro de version dans `src/AppSettings.cs`, ligne en tête de `src/NOUVEAUTES.txt`, nombre de contrôles dans `README.md` et `.github/workflows/tests.yml`.
- Les tests Windows se lancent par GitHub Actions (`tests.yml`, workflow_dispatch) : on ne peut pas compiler sur Linux.

## À chaque réponse
Terminer par la liste des corrections en attente de release (une ligne par correction : description, commit, résultat des tests Windows), pour que Maxime sache où il en est.

## Corrections en attente de release (prochaine : 1.0.3)
1. Korecto reste affiché quand un menu ou une boîte de dialogue d'Outlook s'ouvre — `382fbd6`
2. Appliquer garde les liens et mentions @ tels quels et corrige le reste mot par mot — `81b5ccd`
3. Reformuler disponible dès qu'il y a du texte, sans corriger d'abord — `1ac21e1`
4. Icône : K recentré visuellement (décalé de 3,5/128 vers la droite) — `e6437e1`

Aucune n'a été essayée dans un vrai Outlook. Vider cette liste à chaque release.
