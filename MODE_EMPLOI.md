# Recherche emails (Outlook / .eml) — Mode d'emploi

Petite application maison pour rechercher vos emails selon **plusieurs
critères en même temps**, depuis Outlook **ou** un dossier de fichiers `.eml`.

## Deux sources au choix (en haut de la fenêtre)

**Outlook** : se connecte à Outlook bureau.

- Laissez le champ « Boîte (email) » vide pour utiliser votre compte par défaut.
- Ou tapez une adresse précise (la vôtre ou une boîte partagée à laquelle vous
  avez accès) pour explorer ses dossiers.
- Cliquez sur **Connecter / Reconnecter**, puis choisissez le dossier dans la liste.

**Fichiers .eml (dossier local)** : recherche dans un dossier contenant des
fichiers `.eml`. **Ne nécessite pas Outlook.**

- Cliquez sur **Parcourir...** pour choisir le dossier.
- Cochez « Inclure les sous-dossiers » si besoin.

## Critères de recherche

- **Mots-clés** avec **ET** (tous les mots) ou **OU** (au moins un), dans
  l'objet, le corps, ou les deux.
- **Expéditeur**, **Objet contient**, **Corps contient**.
- **Plage de dates** (du… au…), au format AAAA-MM-JJ.
- **Uniquement les mails avec pièce jointe**.
- **Max résultats** pour limiter l'attente (500 par défaut).

## Résultats

- Tableau **triable** : cliquez sur un en-tête de colonne.
- **Double-clic** sur une ligne : ouvre le mail (dans Outlook, ou le fichier
  `.eml` avec votre logiciel par défaut).
- **Exporter (CSV)** : enregistre la liste, ouvrable dans Excel.

## Enregistrer / recharger votre configuration

- **Enregistrer config** : sauvegarde tous vos critères et la source dans un
  fichier `config_recherche_outlook.json` (à côté du programme).
- La configuration est **rechargée automatiquement** au prochain démarrage.
- **Charger config** : la recharge manuellement.

## Prérequis

- **Mode Outlook** : Windows + **Outlook bureau** (Office / Microsoft 365,
  *pas* le « nouveau Outlook ») + le module `pywin32` :
  ```
  pip install pywin32
  ```
- **Mode .eml** : Python seul, aucune installation supplémentaire.

## Lancement

```
python recherche_outlook.py
```

## Si la connexion Outlook reste bloquée

Au bout de ~25 secondes, un message d'aide s'affiche. Causes fréquentes :

1. Une fenêtre Outlook cachée (choix de profil ou message de sécurité) attend
   une validation : cherchez-la dans la barre des tâches.
2. Vous utilisez le « nouveau Outlook » de Windows, qui ne permet pas ce type
   d'automatisation. Utilisez le Outlook classique, ou passez en mode `.eml`.
