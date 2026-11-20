// Créneaux libres des consultants. Tout est en "heure murale" de Paris
// (2026-11-24T10:00) : c'est ce que renvoie Microsoft Graph quand on lui
// demande le fuseau de Paris, et c'est ce qu'on lui renvoie pour créer le RDV.

export const PARIS = 'Europe/Paris';
const pad = (n) => String(n).padStart(2, '0');

// "2026-11-24T10:00[:00.0000000]" -> minutes (horloge murale, comparables entre elles)
export function wallToMin(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (!m) return NaN;
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) / 60000;
}
export function minToWall(min) {
  const d = new Date(min * 60000);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
}
export const hhmmToMin = (s) => {
  const [h, m] = String(s).split(':').map(Number);
  return h * 60 + (m || 0);
};

// Heure murale de Paris "maintenant"
export function parisNow(now = new Date()) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: PARIS, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(now).map((x) => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}`;
}

// Heure murale d'un fuseau -> instant réel (Date)
export function wallToDate(wall, tz = PARIS) {
  const guess = wallToMin(wall) * 60000;
  // décalage du fuseau à cet instant (deux passes pour les changements d'heure)
  let t = guess;
  for (let i = 0; i < 2; i++) t = guess - offsetMin(tz, new Date(t)) * 60000;
  return new Date(t);
}
function offsetMin(tz, date) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(date).map((x) => [x.type, x.value]));
  const local = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute);
  return Math.round((local - Math.floor(date.getTime() / 60000) * 60000) / 60000);
}

// Jours ouvrés du lendemain jusqu'à la même date dans N mois (heure de Paris)
export function daysAhead(nowWall, months = 1, days = [1, 2, 3, 4, 5]) {
  const [y, m, d] = nowWall.slice(0, 10).split('-').map(Number);
  const lastOfTarget = new Date(Date.UTC(y, m - 1 + months + 1, 0)).getUTCDate();
  const end = Date.UTC(y, m - 1 + months, Math.min(d, lastOfTarget)) / 86400000;
  const out = [];
  for (let k = Date.UTC(y, m - 1, d) / 86400000 + 1; k <= end; k++) {
    if (days.includes(new Date(k * 86400000).getUTCDay())) out.push(minToWall(k * 1440).slice(0, 10));
  }
  return out;
}

// Un événement occupe le consultant sauf s'il est marqué "libre" ou "ailleurs"
export const isBusy = (ev) => !ev.isCancelled && !['free', 'workingElsewhere'].includes(ev.showAs);

/**
 * Créneaux d'une journée pour chaque consultant.
 * @param {string} day  "2026-11-24"
 * @param {{start:string,end:string,step:number,duration:number,minNotice:number}} rules
 * @param {Record<string, Array<{start:string,end:string,showAs?:string,isCancelled?:boolean}>>} busyBy
 * @param {string} nowWall heure murale de Paris maintenant
 * @returns {Array<{start:string,end:string,free:Record<string,boolean>}>}
 */
export function daySlots(day, rules, busyBy, nowWall) {
  const base = wallToMin(day + 'T00:00');
  const first = base + hhmmToMin(rules.start);
  const last = base + hhmmToMin(rules.end) - rules.duration;
  const earliest = wallToMin(nowWall) + (rules.minNotice ?? 30);
  const intervals = Object.fromEntries(Object.entries(busyBy).map(([k, evs]) =>
    [k, evs.filter(isBusy).map((e) => [wallToMin(e.start), wallToMin(e.end)])]));
  const out = [];
  for (let s = first; s <= last; s += rules.step) {
    if (s < earliest) continue;
    const e = s + rules.duration;
    const free = {};
    for (const [k, list] of Object.entries(intervals)) free[k] = !list.some(([bs, be]) => bs < e && be > s);
    out.push({ start: minToWall(s), end: minToWall(e), free });
  }
  return out;
}

// "mardi 24 novembre"
export function dayLabel(day, opts = { weekday: 'long', day: 'numeric', month: 'long' }) {
  return new Intl.DateTimeFormat('fr-FR', { ...opts, timeZone: 'UTC' }).format(new Date(wallToMin(day + 'T12:00') * 60000));
}
export const timeOf = (wall) => wall.slice(11, 16).replace(':', 'h');

