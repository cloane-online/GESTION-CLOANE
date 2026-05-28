# CLOANE — Commandes Clients B2B

Application web autonome pour gérer les commandes clients, avec **enregistrement en ligne
sur Supabase** (utilisable depuis plusieurs appareils en temps quasi réel) et **hébergement
sur Netlify**.

---

## 1. Créer la base Supabase

1. Aller sur [supabase.com](https://supabase.com) → **New project** (notez le mot de passe de la base).
2. Dans le menu **SQL Editor → New query**, collez le contenu de **`supabase/schema.sql`** puis **Run**.
3. Toujours dans SQL Editor, collez le contenu de **`supabase/seed.sql`** puis **Run**
   (cela insère les commandes du 03/01/2026 au 22/05/2026).
4. Dans **Project Settings → API**, récupérez :
   - **Project URL** → `VITE_SUPABASE_URL`
   - **anon public** key → `VITE_SUPABASE_ANON_KEY`

---

## 2. Lancer en local (sur Mac)

```bash
# Node.js requis (sinon : brew install node)
npm install

# Créer le fichier de configuration
cp .env.example .env
# puis éditez .env avec votre URL et votre clé anon

npm run dev
```

Ouvrez l'adresse affichée (ex. `http://localhost:5173`).

---

## 3. Déployer sur Netlify

### Option A — Connecté à GitHub (recommandé)
1. Poussez ce dossier sur un dépôt GitHub.
2. Sur [app.netlify.com](https://app.netlify.com) → **Add new site → Import from GitHub**.
3. Réglages (déjà dans `netlify.toml`) : build `npm run build`, publish `dist`.
4. **Site settings → Environment variables**, ajoutez :
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
5. **Deploy**. À chaque modification poussée sur GitHub, le site se reconstruit tout seul.

### Option B — Glisser-déposer
```bash
npm install
npm run build      # crée le dossier dist/
```
Glissez le dossier **`dist`** sur [app.netlify.com/drop](https://app.netlify.com/drop).
⚠️ Avec cette méthode, les variables Supabase doivent être présentes **au moment du build**
(dans votre `.env` local).

---

## 4. Fonctionnement

- **Données partagées** : toutes les commandes sont stockées dans Supabase. Chaque appareil
  qui ouvre le site voit les mêmes données (rechargez la page pour récupérer les ajouts récents).
- **Création / modification / suppression / changement de statut groupé** sont enregistrés
  automatiquement en ligne.
- **Purge automatique** au chargement : les commandes de plus de 60 jours (hors « Commande
  validée / réceptionnée ») et les « Produit encaissé » de plus de 15 jours sont retirées —
  de l'écran **et** de la base.
- **Sauvegarde Excel quotidienne** : bouton « 💾 Sauvegarde du jour » (rotation lundi→samedi,
  fichiers `CLOANE_Commandes.lundi.xlsx`, etc.). C'est un export local en complément de Supabase.
- **Ticket caisse**, **contact client WhatsApp/SMS/e-mail**, **bouton Commander (URL B2B)**,
  **filtres**, **couleurs + Autre**, **jour de récupération** : voir l'interface.

---

## 5. Sécurité

Le schéma active RLS avec une policy ouverte à la clé **anon** (lecture + écriture) : pratique
pour un outil interne, mais **toute personne ayant l'URL du site peut modifier les données**.
Pour restreindre l'accès, activez l'**authentification Supabase** et remplacez la policy
`acces_complet_anon` dans `schema.sql` par une règle basée sur `auth.uid()`.

---

## Structure

```
cloane-commandes/
├── index.html
├── package.json
├── vite.config.js
├── netlify.toml
├── .env.example
├── src/
│   ├── main.jsx
│   ├── App.jsx
│   ├── CommandesB2B.jsx     ← l'application
│   └── supabaseClient.js    ← connexion + lecture/écriture Supabase
└── supabase/
    ├── schema.sql           ← à exécuter en 1er
    └── seed.sql             ← données initiales (84 commandes)
```
