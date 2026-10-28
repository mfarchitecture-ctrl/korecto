# Essayer Korecto sans publier ni mettre à jour la version installée

## Voir la version interne

Dans le dossier `dist`, double-cliquer sur `VOIR_KORECTO.cmd`.

La fenêtre porte la mention « TEST INTERNE ». Elle affiche un texte fictif et le bouton « Mise à jour disponible » simulé. Ce bouton ne télécharge rien dans ce mode. Le programme installé, ses réglages, son démarrage Windows et la Release GitHub ne sont pas remplacés. Les réglages de test sont conservés dans `dist/DONNEES_TEST`. Le choix IA existant est lu au premier lancement ; changer ce choix dans la fenêtre de test ne modifie pas les réglages installés.

Les corrections IA ne partent que sur clic et utilisent le compte du fournisseur choisi. Une correction réelle consomme les limites de ce compte, même en mode de test.

## Alt + K

Garder cette version interne ouverte. Dans une autre application, sélectionner du texte, puis appuyer sur Alt + K et relâcher les touches. La sélection arrive dans l’onglet « Texte d’origine », en mode Manuel, sans appel IA. Choisir ensuite Corriger ou Reformuler. Copier le résultat pour le recoller soi-même.

Le raccourci simule une copie de la sélection puis tente de rétablir le contenu précédent du presse-papiers. Il n’utilise pas Ctrl + A et n’effectue pas de collage automatique. L’application source doit accepter Ctrl + C. Certains champs protégés, applications exécutées comme administrateur ou raccourcis déjà enregistrés peuvent empêcher cette capture.

En cas de conflit, fermer l’autre version de Korecto ou l’application qui réserve Alt + K. Fermer la fenêtre avec la croix masque Korecto ; utiliser son icône près de l’horloge → Quitter pour libérer le raccourci.

## Vérifications réalisées

Contrôles C# locaux, dont import sans appel IA, rejet des réponses périmées, reformulation sur clic et limites de taille.

La capture physique Alt + K dans les applications réelles reste à valider sur le poste. Les tests automatisés n’actionnent pas le vrai clavier et ne lisent pas de mails personnels.

