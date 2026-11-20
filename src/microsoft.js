// Connexion Microsoft 365 (MSAL) et appels Microsoft Graph.
// Tout se passe dans le navigateur : le jeton reste dans ce navigateur et
// n'est envoyé qu'à graph.microsoft.com.
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';
import { MICROSOFT, SCOPES, GRAPH_TZ } from './config.js';
import { wallToDate } from './slots.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';

// La page sert aussi de page de retour après la connexion (fenêtre Microsoft).
export function isAuthResponse() {
  const s = location.hash + '&' + location.search;
  return /[#?&]state=/.test(s) && /[#?&](code|error)=/.test(s);
}
export const finishAuthPopup = () => broadcastResponseToMainFrame();

export async function microsoftBackend() {
  const pca = new PublicClientApplication({
    auth: {
      clientId: MICROSOFT.clientId,
      authority: `https://login.microsoftonline.com/${MICROSOFT.tenantId || 'organizations'}`,
      redirectUri: new URL('./', location.href).href,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  const pickAccount = () => pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  if (pickAccount()) pca.setActiveAccount(pickAccount());

  async function token() {
    const account = pickAccount();
    if (!account) throw friendly({ code: 'not_signed_in' });
    try {
      return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        try { return (await pca.acquireTokenPopup({ scopes: SCOPES, account })).accessToken; } catch (e2) { throw friendly(e2); }
      }
      throw friendly(e);
    }
  }

  async function graph(path, { method = 'GET', body, headers = {} } = {}) {
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
      method,
      headers: { Authorization: `Bearer ${await token()}`, 'Content-Type': 'application/json', ...headers },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw friendly({ status: res.status, code: j.error && j.error.code, message: j.error && j.error.message });
    }
    return res.status === 204 ? null : res.json();
  }

  return {
    demo: false,
    account() {
      const a = pickAccount();
      return a ? { name: a.name || a.username, email: (a.username || '').toLowerCase() } : null;
    },
    async signIn() {
      try {
        const r = await pca.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
        pca.setActiveAccount(r.account);
      } catch (e) { throw friendly(e); }
      return this.account();
    },
    async signOut() {
      const account = pickAccount();
      // on oublie le compte dans ce navigateur, sans déconnecter Outlook
      if (account) await pca.clearCache({ account });
    },
    // agendas partagés avec moi (ceux des consultants), avec droit de modification
    async calendars() {
      const me = this.account();
      const r = await graph('/me/calendars?$select=id,name,owner,canEdit&$top=200');
      const seen = new Map();
      for (const c of r.value || []) {
        const email = (c.owner && c.owner.address || '').toLowerCase();
        if (!email || (me && email === me.email) || seen.has(email)) continue;
        seen.set(email, { email, name: c.owner.name || email, canEdit: !!c.canEdit });
      }
      return [...seen.values()];
    },
    // événements entre deux heures murales de Paris, renvoyés en heure de Paris
    async events(email, fromWall, toWall) {
      const q = new URLSearchParams({
        startDateTime: wallToDate(fromWall).toISOString(),
        endDateTime: wallToDate(toWall).toISOString(),
        $select: 'start,end,showAs,isCancelled,isAllDay',
        $top: '500',
      });
      let url = `/users/${encodeURIComponent(email)}/calendar/calendarView?${q}`;
      const out = [];
      while (url) {
        const r = await graph(url, { headers: { Prefer: `outlook.timezone="${GRAPH_TZ}"` } });
        for (const e of r.value || []) {
          out.push({ start: e.start.dateTime, end: e.end.dateTime, showAs: e.showAs, isCancelled: e.isCancelled });
        }
        url = r['@odata.nextLink'] || null;
      }
      return out;
    },
    async createEvent(email, ev) {
      const r = await graph(`/users/${encodeURIComponent(email)}/calendar/events`, {
        method: 'POST',
        headers: { Prefer: `outlook.timezone="${GRAPH_TZ}"` },
        body: {
          subject: ev.subject,
          start: { dateTime: ev.start + ':00', timeZone: GRAPH_TZ },
          end: { dateTime: ev.end + ':00', timeZone: GRAPH_TZ },
          attendees: ev.attendees.map((address) => ({ emailAddress: { address }, type: 'required' })),
          isOnlineMeeting: true,
          onlineMeetingProvider: 'teamsForBusiness',
          showAs: 'busy',
        },
      });
      return {
        webLink: r.webLink || null,
        teams: !!(r.isOnlineMeeting && r.onlineMeeting && r.onlineMeeting.joinUrl),
      };
    },
  };
}

// Messages compréhensibles pour les erreurs les plus probables
function friendly(e) {
  const raw = `${e && (e.errorCode || e.code) || ''} ${e && (e.errorMessage || e.message) || ''}`;
  let msg = '';
  if (/user_cancelled|popup_window_error|empty_window_error/i.test(raw)) msg = 'Connexion annulée ou fenêtre bloquée : autorise les fenêtres pop-up pour cette page, puis réessaie.';
  else if (/not_signed_in/.test(raw)) msg = 'Connecte-toi avec ton compte Microsoft.';
  else if (/AADSTS65001|consent/i.test(raw)) msg = "L'admin Microsoft 365 doit encore « accorder le consentement » pour l'outil.";
  else if (/AADSTS50105|AADSTS50177|not assigned/i.test(raw)) msg = "Ton compte n'est pas autorisé à utiliser l'outil : demande à l'admin de t'ajouter dans « Utilisateurs et groupes ».";
  else if (/AADSTS700016|AADSTS90002/i.test(raw)) msg = "Identifiant de l'application ou de l'annuaire incorrect dans les réglages de l'outil.";
  else if (/ErrorAccessDenied|AccessDenied|403/.test(raw) || e.status === 403) msg = "Accès refusé à cet agenda : il faut avoir le droit de modifier l'agenda du consultant dans Outlook.";
  else if (/ErrorItemNotFound|ErrorInvalidUser|MailboxNotEnabled/.test(raw) || e.status === 404) msg = 'Agenda introuvable : vérifie l’adresse e-mail du consultant dans les réglages.';
  else if (e.status === 429 || e.status >= 500) msg = 'Microsoft ne répond pas pour le moment, réessaie dans un instant.';
  const err = new Error(msg || `Erreur Microsoft : ${raw.trim() || 'inconnue'}`);
  err.detail = raw.trim();
  return err;
}
