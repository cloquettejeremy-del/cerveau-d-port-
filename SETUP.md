# Décharge mentale — guide de mise en route

Cette application est un "PWA" : un site web que tu installes sur ton téléphone comme une vraie appli. Elle n'a pas de serveur — tout tourne dans ton téléphone, et elle parle directement à ton compte Microsoft (Outlook) pour créer les rendez-vous et les tâches.

Il y a deux choses à faire une seule fois avant de pouvoir l'utiliser :

1. Héberger les fichiers quelque part (une adresse web, gratuite)
2. Créer un "Client ID" Azure pour autoriser l'appli à parler à ton Outlook

Compte 10-15 minutes la première fois. Ensuite, plus rien à faire.

## Étape 1 — Héberger l'application

Il te faut une adresse en `https://` pour que Microsoft accepte de s'y connecter (ça ne marche pas en local).

### Option simple : GitHub Pages (gratuit, recommandé si tu as déjà un compte GitHub)

1. Crée un nouveau dépôt sur [github.com](https://github.com) (ex: `decharge-mentale`), public.
2. Mets-y tous les fichiers de ce dossier (`index.html`, `style.css`, `app.js`, `sw.js`, `manifest.json`, le dossier `icons/`).
3. Dans le dépôt : **Settings → Pages → Source : "Deploy from a branch"**, branche `main`, dossier `/root` → Save.
4. Après une minute, ton app est disponible à une adresse du type :
   `https://TON-PSEUDO.github.io/decharge-mentale/`
5. Note cette adresse, tu en auras besoin à l'étape 2.

### Option zéro-compte : Netlify Drop

1. Va sur [app.netlify.com/drop](https://app.netlify.com/drop).
2. Glisse-dépose le dossier entier de l'application dans la page.
3. Netlify te donne une adresse `https://un-nom-au-hasard.netlify.app`.
4. Tu peux créer un compte gratuit ensuite pour rendre le site permanent (sinon il peut expirer après un moment d'inactivité).

## Étape 2 — Créer le Client ID Azure (accès à Outlook)

C'est gratuit et ne nécessite pas d'abonnement payant, que ton compte Microsoft soit personnel ou professionnel.

1. Va sur [portal.azure.com](https://portal.azure.com) et connecte-toi avec ton compte Microsoft (celui qui a ton Outlook).
2. Cherche **"Inscriptions d'applications"** (App registrations) dans la barre de recherche en haut.
3. Clique **Nouvelle inscription**.
4. Nom : `Décharge mentale` (ou ce que tu veux).
5. **Types de comptes pris en charge** : choisis *"Comptes dans n'importe quel annuaire organisationnel et comptes Microsoft personnels"*.
6. **URI de redirection** : type = **"Application monopage (SPA)"**, valeur = l'adresse notée à l'étape 1 (ex: `https://ton-pseudo.github.io/decharge-mentale/index.html`).
7. Clique **Inscrire**.
8. Sur la page qui s'ouvre, copie la valeur **"ID d'application (client)"** — c'est une suite de lettres/chiffres du type `a1b2c3d4-....`. C'est ton **Client ID**.
9. Dans le menu de gauche, va dans **Authentification** → vérifie que sous "Application monopage" ton URI de redirection est bien listée, et coche **"Jetons d'accès"** et **"Jetons ID"** si ce n'est pas déjà fait → Enregistrer.
10. Dans **Autorisations API**, ajoute (si elles n'y sont pas déjà) : `Calendars.ReadWrite`, `Tasks.ReadWrite`, `User.Read` (type "Microsoft Graph" → "Autorisations déléguées"). Un compte personnel n'a pas besoin d'un admin pour valider ces autorisations ; un compte professionnel peut te demander l'accord de ton administrateur informatique la première fois.

Tu n'as rien d'autre à configurer (pas de "secret" à créer — l'appli n'en a pas besoin, elle est 100% côté téléphone).

## Étape 3 — Premier lancement

1. Ouvre l'adresse de ton app (étape 1) dans le navigateur de ton téléphone.
2. Va dans **⚙️ Réglages**.
3. Colle ton **Client ID** (étape 2) dans le champ prévu, puis **Enregistrer les réglages**.
4. Clique **Se connecter** et connecte-toi avec ton compte Microsoft. Accepte les autorisations demandées.
5. Toujours dans Réglages, personnalise tes catégories : renomme "Entreprise 1/2/3" avec les vrais noms de tes sociétés, et ajoute quelques mots-clés pour chacune (noms de clients, de projets...) — ça aide l'appli à deviner automatiquement la bonne catégorie.
6. Enregistre.

## Étape 4 — Installer sur l'écran d'accueil

- **iPhone (Safari)** : ouvre l'adresse → bouton Partager (carré avec flèche) → **"Sur l'écran d'accueil"**.
- **Android (Chrome)** : ouvre l'adresse → menu ⋮ → **"Ajouter à l'écran d'accueil"** (ou une bannière d'installation apparaît automatiquement).

L'icône apparaît comme une vraie application. C'est celle-ci que tu utilises au quotidien.

## Comment ça marche au quotidien

1. Tu ouvres l'app, tu touches le micro et tu parles (ou tu écris).
2. Tu touches "Continuer" : l'appli propose automatiquement une catégorie, un type (Tâche / Événement / Courses) et une date si elle en a détecté une.
3. Tu corriges si besoin en un tap, puis "Envoyer vers Outlook".
4. C'est dans ton calendrier ou ta liste de tâches Outlook (liste **"🧠 Décharge mentale"**), avec une couleur/catégorie Outlook selon l'entreprise concernée.

Si tu es hors-ligne (pas de réseau), l'élément reste dans le **Journal** (icône 📋) avec un badge, et repart automatiquement dès que le réseau revient. Tu peux aussi forcer un nouvel essai depuis le Journal.

## Limites connues (v1)

- La reconnaissance vocale fonctionne dans Chrome/Edge (Android et desktop) et Safari iOS récent. Si ton navigateur ne la supporte pas, le champ de texte reste utilisable au clavier.
- La détection de date/heure comprend les formulaires courants en français ("demain à 14h", "lundi", "dans 2 heures", "12/09", "12 septembre") mais n'est pas parfaite — tu peux toujours corriger la date à la main sur l'écran de confirmation.
- Un compte professionnel (Microsoft 365 entreprise) peut nécessiter que ton administrateur informatique valide les autorisations la première fois — si la connexion échoue avec un message d'"consentement administrateur requis", il faudra lui demander de valider l'inscription d'application une fois.

## Et après ?

Cette v1 est volontairement simple pour être fiable. Des pistes d'amélioration possibles pour plus tard : ajouter une intelligence plus fine (relié à Claude) pour reformuler et catégoriser les phrases les plus ambiguës, prise en charge de rappels multiples dans une seule phrase, synchronisation avec d'autres outils que Outlook. Dis-le quand tu veux qu'on y retourne.
