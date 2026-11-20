// Page "RDV consultants" : créneaux libres des consultants et création du RDV
// dans leur agenda Outlook (titre, heure, réunion Teams, admin commerce invitée).
import { h, copyText } from './dom.js';
import { MICROSOFT, DURATION, HOURS, STEP, MONTHS_AHEAD } from './config.js';
import { daySlots, daysAhead, parisNow, dayLabel, timeOf, wallToMin, minToWall } from './slots.js';
import { makeTitle } from './text.js';
import { demoBackend } from './demo.js';
import { isAuthResponse, finishAuthPopup, microsoftBackend } from './microsoft.js';

const VERSION = __VERSION__;
const DEMO = new URLSearchParams(location.search).has('demo');
const CFG_KEY = DEMO ? 'rdv.cfg.demo' : 'rdv.cfg';
const app = document.getElementById('app');
const EMAIL = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

// Réglages gardés dans ce navigateur (les adresses ne sont pas dans le code)
function loadCfg() {
  const empty = { consultants: [], admin: '', titre: '' };
  try { return { ...empty, ...JSON.parse(localStorage.getItem(CFG_KEY) || '{}') }; } catch { return empty; }
}
function saveCfg() {
  try { localStorage.setItem(CFG_KEY, JSON.stringify(S.cfg)); } catch { /* stockage bloqué */ }
}

const S = {
  backend: null,
  cfg: loadCfg(),
  client: { civilite: '', nom: '' },
  days: [],
  day: null,
  events: {}, // email -> événements
  errors: {}, // email -> message
  loadedAt: 0,
  loading: false,
  pick: null, // { slot, consultant }
  done: null,
  view: '',
};
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

// ------------------------------------------------------------- en-tête
function header() {
  const acc = S.backend && S.backend.account && S.backend.account();
  const right = [];
  if (acc) {
    right.push(h('span', { class: 'muted' }, acc.name));
    right.push(h('button', { class: 'link', type: 'button', onclick: () => showSettings() }, 'Réglages'));
    if (!S.backend.demo) {
      right.push(h('button', { class: 'link', type: 'button', onclick: async () => { await S.backend.signOut(); showSignIn(); } }, 'Déconnexion'));
    }
  }
  return h('header', null,
    h('span', { class: 'logo' }),
    h('h1', null, 'RDV consultants'),
    S.backend && S.backend.demo && h('span', { class: 'badge' }, 'démo, faux agendas'),
    h('span', { class: 'sp' }),
    right);
}
function screen(...kids) {
  app.replaceChildren(header(), ...kids, h('footer', { class: 'muted' }, `RDV consultants ${VERSION}`));
}

// ------------------------------------------------------------- écrans simples
function showNotConfigured() {
  S.view = 'wait';
  screen(h('section', null,
    h('h2', null, "En attente de l'admin Microsoft 365"),
    h('p', null, "L'outil pourra lire et remplir les agendas des consultants dès que l'admin aura inscrit l'application dans Microsoft 365."),
    h('p', null, 'En attendant, tu peux l’essayer avec de faux agendas : rien n’est envoyé nulle part.'),
    h('a', { class: 'primary', href: '?demo' }, 'Essayer la démo')));
}

function showSignIn(err) {
  S.view = 'signin';
  const msg = h('div', { class: 'err' }, err || '');
  const btn = h('button', { class: 'primary', type: 'button' }, 'Se connecter avec Microsoft');
  btn.addEventListener('click', async () => {
    btn.disabled = true;
    msg.textContent = '';
    try { await S.backend.signIn(); start(); } catch (e) { msg.textContent = e.message; btn.disabled = false; }
  });
  screen(h('section', null,
    h('h2', null, 'Connexion'),
    h('p', null, 'Connecte-toi avec ton compte Microsoft pro. Ça ne se fait qu’une fois sur ce navigateur.'),
    h('p', { class: 'muted' }, 'L’outil voit seulement ce que ton compte voit déjà dans Outlook.'),
    btn, msg));
}

// ------------------------------------------------------------- réglages
function showSettings() {
  S.view = 'settings';
  const list = h('div', { class: 'cals' }, h('span', { class: 'muted' }, 'Recherche des agendas partagés avec toi…'));
  const chosen = new Map(S.cfg.consultants.map((c) => [c.email, c]));
  let known = [...S.cfg.consultants];

  const drawList = () => {
    if (!known.length) {
      list.replaceChildren(h('span', { class: 'muted' }, 'Aucun agenda partagé trouvé. Ajoute les consultants par leur e-mail ci-dessous.'));
      return;
    }
    list.replaceChildren(...known.map((c) => {
      const box = h('input', { type: 'checkbox' });
      box.checked = chosen.has(c.email);
      box.addEventListener('change', () => { if (box.checked) chosen.set(c.email, c); else chosen.delete(c.email); });
      return h('label', { class: 'cal' }, box, h('b', null, c.name), h('span', { class: 'muted' }, c.email),
        c.canEdit === false && h('span', { class: 'warn-txt' }, 'lecture seule'));
    }));
  };
  S.backend.calendars().then((cals) => {
    for (const c of cals) {
      if (!known.some((k) => k.email === c.email)) known.push(c);
      else known = known.map((k) => (k.email === c.email ? { ...k, canEdit: c.canEdit } : k));
    }
    drawList();
  }).catch((e) => { drawList(); list.prepend(h('div', { class: 'err' }, e.message)); });

  const addName = h('input', { type: 'text', placeholder: 'Nom du consultant' });
  const addMail = h('input', { type: 'email', placeholder: 'e-mail du consultant' });
  const addBtn = h('button', { class: 'secondary', type: 'button' }, 'Ajouter');
  addBtn.addEventListener('click', () => {
    const email = addMail.value.trim().toLowerCase();
    if (!EMAIL.test(email)) { addMail.focus(); return; }
    const c = { email, name: addName.value.trim() || email };
    if (!known.some((k) => k.email === email)) known.push(c);
    chosen.set(email, c);
    addName.value = '';
    addMail.value = '';
    drawList();
  });

  const admin = h('input', { type: 'email', value: S.cfg.admin, placeholder: 'admin.commerce@…' });
  const titre = h('input', { type: 'text', value: S.cfg.titre, placeholder: 'ex. Consultation + nom de la société' });
  const preview = h('div', { class: 'muted' });
  const drawPreview = () => { preview.textContent = `Le RDV s'appellera : « ${makeTitle(titre.value, { civilite: 'Madame', nom: 'Jeanne DUPONT' })} »`; };
  titre.addEventListener('input', drawPreview);
  drawPreview();
  const msg = h('div', { class: 'err' });
  const save = h('button', { class: 'primary', type: 'button' }, 'Enregistrer');
  save.addEventListener('click', () => {
    const consultants = known.filter((k) => chosen.has(k.email)).map(({ email, name }) => ({ email, name }));
    const adminMail = admin.value.trim();
    if (!consultants.length) { msg.textContent = 'Choisis au moins un agenda de consultant.'; return; }
    if (!EMAIL.test(adminMail)) { msg.textContent = "Indique l'e-mail de l'admin commerce."; admin.focus(); return; }
    if (!titre.value.trim()) { msg.textContent = 'Indique le début du titre du RDV.'; titre.focus(); return; }
    S.cfg = { consultants, admin: adminMail, titre: titre.value.trim() };
    saveCfg();
    S.loadedAt = 0;
    showMain();
  });
  const cancel = S.cfg.consultants.length > 0 && h('button', { class: 'link', type: 'button', onclick: () => showMain() }, 'Annuler');

  screen(h('section', { class: 'settings' },
    h('h2', null, 'Réglages'),
    h('p', { class: 'muted' }, 'À faire une seule fois sur ce navigateur.'),
    h('h3', null, 'Agendas des consultants'),
    list,
    h('div', { class: 'row' }, addName, addMail, addBtn),
    h('h3', null, 'Admin commerce (invitée à chaque RDV)'),
    admin,
    h('h3', null, 'Titre du RDV'),
    h('div', { class: 'muted' }, 'Ce qui vient avant Madame/Monsieur et le nom du client.'),
    titre,
    preview,
    favoriBlock(),
    msg,
    h('div', { class: 'row' }, save, cancel)));
  drawList();
}

// Favori à glisser dans la barre de favoris : sur la fiche Salesforce, il lit
// Madame/Monsieur et le nom du client et ouvre cette page avec.
function favoriBlock() {
  const page = location.origin + location.pathname + (DEMO ? '?demo' : '');
  // le code du favori est une adresse javascript: déjà encodée
  const code = __BOOKMARKLET__.split('__RDV_URL__').join(page.replace(/%/g, '%25').replace(/#/g, '%23'));
  const link = h('a', { class: 'bm', href: code, title: 'Glisse-moi dans ta barre de favoris' }, 'RDV');
  link.addEventListener('click', (e) => { e.preventDefault(); alert('Glisse ce bouton dans ta barre de favoris, puis clique dessus depuis la fiche Salesforce du client.'); });
  const copy = h('button', { class: 'secondary', type: 'button' }, 'Copier le code du favori');
  copy.addEventListener('click', async () => { copy.textContent = (await copyText(code)) ? 'Code copié ✓' : 'Échec de la copie'; });
  return h('div', null,
    h('h3', null, 'Favori Salesforce'),
    h('div', { class: 'row' }, link, h('span', { class: 'muted' }, '← glisse ce bouton dans ta barre de favoris (Ctrl+Maj+B pour l’afficher)')),
    h('div', { class: 'muted' }, 'Sur la fiche du client, clique sur RDV : cette page s’ouvre avec Madame/Monsieur et le nom déjà remplis.'),
    h('div', { class: 'muted' }, 'Si le glisser ne marche pas : ', copy, ', puis clic droit sur la barre de favoris → Ajouter une page…, nom « RDV », colle le code dans « URL ».'));
}

// Client envoyé par le favori dans l'adresse (#client=…), jamais envoyé au serveur
function readClientFromUrl() {
  const p = new URLSearchParams(location.hash.slice(1));
  if (!p.has('client')) return false;
  try {
    const d = JSON.parse(p.get('client'));
    S.client = {
      civilite: ['Madame', 'Monsieur'].includes(d.civilite) ? d.civilite : '',
      nom: String(d.nom || '').slice(0, 120),
    };
    S.pick = null;
    S.done = null;
  } catch { /* ignoré */ }
  history.replaceState(null, '', location.pathname + location.search);
  return true;
}

// ------------------------------------------------------------- écran principal
async function loadEvents(force = false) {
  if (S.loading || (!force && Date.now() - S.loadedAt < 60000)) return;
  S.loading = true;
  renderSlots();
  S.days = daysAhead(parisNow(), MONTHS_AHEAD);
  const from = S.days[0] + 'T00:00';
  const to = minToWall(wallToMin(S.days.at(-1) + 'T00:00') + 1440);
  S.errors = {};
  await Promise.all(S.cfg.consultants.map(async (c) => {
    try { S.events[c.email] = await S.backend.events(c.email, from, to); } catch (e) { S.errors[c.email] = e.message; S.events[c.email] = null; }
  }));
  S.loading = false;
  S.loadedAt = Date.now();
  renderSlots();
}

const rules = () => ({ start: HOURS.start, end: HOURS.end, step: STEP, duration: DURATION, minNotice: 0 });
function slotsFor(day) {
  const busy = {};
  for (const c of S.cfg.consultants) busy[c.email] = S.events[c.email] || [];
  const slots = daySlots(day, rules(), busy, parisNow());
  // agenda illisible : aucun créneau proposé pour ce consultant
  for (const s of slots) for (const c of S.cfg.consultants) if (!S.events[c.email]) s.free[c.email] = false;
  return slots;
}

let els = null;
function showMain() {
  S.view = 'main';
  const civ = h('select', { 'aria-label': 'Civilité' },
    h('option', { value: '' }, '—'), h('option', { value: 'Madame' }, 'Madame'), h('option', { value: 'Monsieur' }, 'Monsieur'));
  const nom = h('input', { type: 'text', placeholder: 'Nom du client', 'aria-label': 'Nom du client' });
  civ.value = S.client.civilite;
  nom.value = S.client.nom;
  const days = h('div', { class: 'days', role: 'tablist' });
  const grid = h('div', { class: 'gridwrap' });
  const sheet = h('div', { class: 'sheet', hidden: true });
  els = { civ, nom, days, grid, sheet };

  const sync = () => { S.client = { civilite: civ.value, nom: nom.value }; renderSheet(); };
  civ.addEventListener('change', sync);
  nom.addEventListener('input', sync);

  screen(
    h('section', { class: 'client' }, h('div', { class: 'who' }, civ, nom)),
    h('section', { class: 'slots' }, h('div', { class: 'bar' }, h('b', null, 'Créneaux libres')), days, grid),
    sheet);
  renderSlots();
  renderSheet();
  loadEvents(true);
}

function renderSlots() {
  if (!els) return;
  const { days, grid } = els;
  if (!S.days.length) S.days = daysAhead(parisNow(), MONTHS_AHEAD);
  const consultants = S.cfg.consultants;
  const counts = new Map(S.days.map((d) => [d, slotsFor(d).filter((s) => Object.values(s.free).some(Boolean)).length]));
  if (!S.day || !S.days.includes(S.day) || (!counts.get(S.day) && !S.loading)) {
    S.day = S.days.find((d) => counts.get(d)) || S.days[0];
  }
  days.replaceChildren(...S.days.map((d) => {
    const n = counts.get(d);
    const b = h('button', { type: 'button', class: `day${d === S.day ? ' on' : ''}`, role: 'tab', 'aria-selected': d === S.day ? 'true' : 'false' },
      h('b', null, cap(dayLabel(d, { weekday: 'short', day: 'numeric' }).replace('.', ''))),
      h('span', null, S.loading ? '…' : n ? `${n} créneau${n > 1 ? 'x' : ''}` : 'complet'));
    if (!n && !S.loading) b.disabled = true;
    b.addEventListener('click', () => { S.day = d; S.pick = null; renderSlots(); renderSheet(); });
    return b;
  }));

  if (S.loading && !S.loadedAt) { grid.replaceChildren(h('div', { class: 'muted pad' }, 'Lecture des agendas…')); return; }
  const slots = slotsFor(S.day);
  const errs = Object.entries(S.errors).map(([email, m]) => {
    const c = consultants.find((x) => x.email === email);
    return h('div', { class: 'err' }, `${c ? c.name : email} : ${m}`);
  });
  const thead = h('tr', null, h('th', null, 'Heure'), consultants.map((c) => h('th', null, c.name)));
  const rows = slots.map((s) => h('tr', { class: Object.values(s.free).some(Boolean) ? '' : 'full' },
    h('td', { class: 't' }, timeOf(s.start)),
    consultants.map((c) => {
      if (!s.free[c.email]) return h('td', { class: 'busy' }, '—');
      const on = S.pick && S.pick.slot.start === s.start && S.pick.consultant.email === c.email;
      const b = h('button', { type: 'button', class: `free${on ? ' on' : ''}`, 'aria-label': `${c.name}, ${timeOf(s.start)}` }, 'Libre');
      b.addEventListener('click', () => { S.pick = { slot: s, consultant: c }; S.done = null; renderSlots(); renderSheet(); });
      return h('td', null, b);
    })));
  grid.replaceChildren(...errs,
    slots.length
      ? h('table', { class: 'grid' }, h('thead', null, thead), h('tbody', null, rows))
      : h('div', { class: 'muted pad' }, 'Plus de créneau ce jour-là.'));
}

const titleNow = () => makeTitle(S.cfg.titre, S.client);

function renderSheet() {
  if (!els) return;
  const { sheet } = els;
  if (S.done) { sheet.hidden = false; sheet.replaceChildren(doneView()); return; }
  if (!S.pick) { sheet.hidden = true; sheet.replaceChildren(); return; }
  const { slot, consultant } = S.pick;
  const title = h('input', { type: 'text', value: titleNow(), 'aria-label': 'Titre du RDV' });
  const err = h('div', { class: 'err' });
  const go = h('button', { class: 'primary', type: 'button' }, 'Créer le RDV');
  const cancel = h('button', { class: 'link', type: 'button' }, 'Annuler');
  cancel.addEventListener('click', () => { S.pick = null; renderSlots(); renderSheet(); });
  go.addEventListener('click', async () => {
    if (!S.client.civilite || !S.client.nom.trim()) { err.textContent = 'Indique Madame ou Monsieur et le nom du client.'; els.nom.focus(); return; }
    go.disabled = true;
    go.textContent = 'Création…';
    err.textContent = '';
    const subject = title.value.trim() || titleNow();
    try {
      // le créneau est-il toujours libre ?
      const fresh = await S.backend.events(consultant.email, slot.start, slot.end);
      if (daySlots(slot.start.slice(0, 10), { ...rules(), start: slot.start.slice(11), end: slot.end.slice(11), minNotice: -1e9 },
        { x: fresh }, parisNow()).every((s) => !s.free.x)) {
        S.events[consultant.email] = (S.events[consultant.email] || []).concat(fresh);
        throw new Error('Ce créneau vient d’être pris. Choisis-en un autre.');
      }
      const r = await S.backend.createEvent(consultant.email, { subject, start: slot.start, end: slot.end, attendees: [S.cfg.admin] });
      (S.events[consultant.email] ||= []).push({ start: slot.start, end: slot.end, showAs: 'busy' });
      S.done = { ...S.pick, ...r, title: subject };
      S.pick = null;
      renderSlots();
      renderSheet();
    } catch (e) {
      err.textContent = e.message;
      go.disabled = false;
      go.textContent = 'Créer le RDV';
      renderSlots();
    }
  });
  sheet.hidden = false;
  sheet.replaceChildren(
    h('div', { class: 'when' }, h('b', null, `${cap(dayLabel(slot.start.slice(0, 10)))} · ${timeOf(slot.start)} – ${timeOf(slot.end)}`),
      ' avec ', h('b', null, consultant.name)),
    h('label', { class: 'field wide' }, h('span', null, 'Titre'), title),
    h('div', { class: 'muted' }, `Réunion Teams · Invitée : ${S.cfg.admin}`),
    err,
    h('div', { class: 'row' }, go, cancel));
}

function doneView() {
  const d = S.done;
  const close = h('button', { class: 'link', type: 'button' }, 'Fermer');
  close.addEventListener('click', () => { S.done = null; renderSheet(); });
  return h('div', { class: 'done' },
    h('div', { class: 'ok' }, `✓ RDV créé dans l’agenda de ${d.consultant.name}`),
    h('div', null, `${cap(dayLabel(d.slot.start.slice(0, 10)))} · ${timeOf(d.slot.start)} – ${timeOf(d.slot.end)} · ${d.title}`),
    !d.teams && h('div', { class: 'err' }, 'Le lien Teams n’a pas été ajouté : ouvre le RDV dans Outlook et active « Réunion Teams ».'),
    S.backend.demo && h('div', { class: 'muted' }, 'Démo : rien n’a été créé pour de vrai.'),
    h('div', { class: 'row' },
      d.webLink && h('a', { class: 'secondary', href: d.webLink, target: '_blank', rel: 'noopener' }, 'Ouvrir dans Outlook'),
      close));
}

// ------------------------------------------------------------- démarrage
function start() {
  if (!S.backend.account()) return showSignIn();
  if (!S.cfg.consultants.length || !S.cfg.admin || !S.cfg.titre) return showSettings();
  showMain();
}

(async function main() {
  if (isAuthResponse()) {
    // fenêtre de connexion Microsoft qui revient : on transmet et on se ferme
    try { await finishAuthPopup(); } catch (e) { app.textContent = `Connexion : ${e.message}`; }
    return;
  }
  readClientFromUrl();
  // favori recliqué alors que la page est déjà ouverte
  window.addEventListener('hashchange', () => {
    if (!readClientFromUrl() || S.view !== 'main') return;
    els.civ.value = S.client.civilite;
    els.nom.value = S.client.nom;
    renderSlots();
    renderSheet();
  });
  if (!DEMO && !MICROSOFT.clientId) return showNotConfigured();
  try {
    S.backend = DEMO ? demoBackend() : await microsoftBackend();
  } catch (e) {
    app.textContent = `Impossible de démarrer la connexion Microsoft : ${e.message}`;
    return;
  }
  if (DEMO && !S.cfg.consultants.length) {
    S.cfg.consultants = (await S.backend.calendars()).map(({ email, name }) => ({ email, name }));
    S.cfg.admin ||= 'admin.commerce@exemple.fr';
    S.cfg.titre ||= 'Consultation (démo)';
  }
  start();
  // agendas relus quand on revient sur l'onglet
  window.addEventListener('focus', () => { if (S.view === 'main') loadEvents(false); });
  setInterval(() => { if (S.view === 'main' && !document.hidden) loadEvents(false); }, 60000);
})();
