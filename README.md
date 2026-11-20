# RDV consultants

Prendre un RDV client dans l'agenda Outlook d'un consultant, sans passer par Outlook.

1. Sur la fiche Salesforce du client, un clic sur le favori **RDV** : la page s'ouvre avec Madame / Monsieur et le nom du client (on peut aussi les taper).
2. La page montre les créneaux libres des consultants côte à côte, lus dans leurs agendas : du lundi au vendredi, 9h–18h (heure de Paris), départ à l'heure pile ou à la demie, du lendemain jusqu'à dans un mois.
3. Un clic sur « Libre » puis « Créer le RDV » : le RDV d'1 h est créé dans l'agenda du consultant avec le titre `<début réglé> - Madame Jeanne DUPONT`, la réunion Teams et l'admin commerce invitée. Outlook ajoute lui-même le bloc « Réunion Microsoft Teams » dans le texte.

## Comment ça marche

Tout se passe dans le navigateur, sans serveur. La connexion se fait avec le compte Microsoft 365 de la personne (MSAL, sans secret) et l'outil n'appelle que Microsoft Graph, sur les agendas auxquels ce compte a déjà accès :

- `User.Read` : nom de la personne connectée ;
- `Calendars.ReadWrite.Shared` : lire les créneaux et créer le RDV dans les agendas partagés.

Les agendas des consultants, l'e-mail de l'admin commerce et le début du titre se règlent une fois dans « Réglages » et restent dans le navigateur. Le favori s'installe depuis « Réglages » ; il lit seulement la civilité et le nom sur la fiche.

## Mise en route

1. L'admin Microsoft 365 inscrit une application dans Entra ID : type « Application monopage (SPA) », adresse de retour `https://69heavxn.github.io/rdv-consultants/`, les deux autorisations déléguées ci-dessus.
2. Mettre son `clientId` et le `tenantId` dans `src/config.js`, puis `npm run build`.
3. Tant que `clientId` est vide, la page propose la démo (`?demo`, faux agendas).

Durée, horaires et période proposée : `src/config.js`.

## Développement

```
npm install
npm test
npm run build   # -> docs/ (GitHub Pages)
```

`test/fiche-test.html` est une fausse fiche Salesforce pour essayer le favori.
