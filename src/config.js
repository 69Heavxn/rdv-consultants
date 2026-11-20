// Réglages de l'outil de prise de RDV.

// Identifiants de l'application déclarée par l'admin Microsoft 365 (Entra ID).
// Ce ne sont pas des secrets. Tant qu'ils sont vides, la page propose la démo.
export const MICROSOFT = {
  clientId: '',
  tenantId: '',
};
export const SCOPES = ['User.Read', 'Calendars.ReadWrite.Shared'];
// Heure de Paris au format attendu par Outlook / Microsoft Graph
export const GRAPH_TZ = 'Romance Standard Time';

// Le RDV (le début du titre se règle dans « Réglages »)
export const DURATION = 60; // minutes
export const HOURS = { start: '09:00', end: '18:00' }; // heure de Paris, du lundi au vendredi
export const STEP = 30; // un créneau toutes les 30 min
export const MONTHS_AHEAD = 1; // créneaux proposés du lendemain jusqu'à dans 1 mois (pas de RDV le jour même)
