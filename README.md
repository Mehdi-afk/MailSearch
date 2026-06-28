# MailSearch — Recherche avancée d'emails

Application Windows avec interface graphique pour rechercher des emails selon plusieurs critères simultanément, depuis **Outlook bureau** ou un dossier de **fichiers .eml**.

---

## Fonctionnalités

- **Deux sources** : Outlook (via COM) ou fichiers `.eml` locaux
- **Critères combinables** : mots-clés (ET / OU), expéditeur, objet, corps, plage de dates, pièce jointe
- **Résultats triables** par colonne (date, expéditeur, objet)
- **Double-clic** sur un résultat pour ouvrir le mail directement
- **Export CSV** compatible Excel
- **Sauvegarde de la configuration** (rechargée automatiquement au démarrage)

---

## Prérequis

| Mode | Prérequis |
|------|-----------|
| Outlook | Windows + Outlook bureau (Office / Microsoft 365) + `pywin32` |
| Fichiers .eml | Python seul, aucune dépendance supplémentaire |

> Le « nouveau Outlook » de Windows n'est pas compatible — utiliser le Outlook classique ou le mode `.eml`.

---

## Installation

```bash
pip install pywin32   # uniquement pour le mode Outlook
```

---

## Lancement

### Via Python

```bash
python recherche_outlook.py
```

### Via le raccourci fourni

Double-cliquer sur **`Lancer_Recherche.bat`**

---

## Utilisation

### 1. Choisir la source

**Outlook**
- Laissez le champ « Boîte » vide pour votre compte par défaut, ou saisissez une adresse pour une boîte partagée
- Cliquez sur **Connecter / Reconnecter**, puis sélectionnez le dossier

**Fichiers .eml**
- Cliquez sur **Parcourir...** et choisissez le dossier contenant les `.eml`
- Cochez « Inclure les sous-dossiers » si besoin

### 2. Définir les critères

- **Mots-clés** : séparés par des espaces, combinés en mode ET (tous présents) ou OU (au moins un)
- **Portée** : objet, corps, ou les deux
- **Expéditeur** / **Objet contient** / **Corps contient**
- **Plage de dates** : format `AAAA-MM-JJ`
- **Pièce jointe uniquement** : filtre les mails sans attachement
- **Max résultats** : limite le nombre de résultats retournés (défaut : 500)

### 3. Lancer la recherche

Cliquez sur **Rechercher**. La barre de statut affiche la progression en temps réel.

### 4. Exploiter les résultats

- Cliquez sur un en-tête de colonne pour trier
- Double-cliquez sur une ligne pour ouvrir le mail
- Cliquez sur **Exporter (CSV)** pour sauvegarder les résultats

---

## Construire l'exécutable (.exe)

Un fichier `.exe` autonome (sans Python requis) peut être généré avec PyInstaller :

```bash
pip install pyinstaller
```

Double-cliquer sur **`Construire_EXE.bat`** — le `.exe` sera généré dans le dossier `dist/`.

---

## Dépannage — Connexion Outlook bloquée

Si la connexion reste bloquée plus de ~25 secondes :

1. Une fenêtre Outlook est peut-être cachée dans la barre des tâches (choix de profil ou alerte de sécurité) — la chercher et la valider
2. Le « nouveau Outlook » de Windows ne supporte pas ce type d'automatisation — utiliser le Outlook classique ou passer en mode `.eml`
