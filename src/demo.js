// Démo : faux agendas pour essayer l'outil sans compte Microsoft.
// Rien n'est envoyé nulle part, les RDV "créés" disparaissent au rechargement.
import { wallToMin, minToWall } from './slots.js';

const PEOPLE = [
  { email: 'consultant1@exemple.fr', name: 'Claire Martin (démo)' },
  { email: 'consultant2@exemple.fr', name: 'Paul Bernard (démo)' },
  { email: 'consultant3@exemple.fr', name: 'Sophie Petit (démo)' },
];

// pseudo-hasard stable : même jour = mêmes RDV
function rand(seed) {
  let x = seed % 2147483647 || 1;
  return () => ((x = (x * 48271) % 2147483647) / 2147483647);
}

export function demoBackend() {
  const created = [];
  return {
    demo: true,
    account: () => ({ name: 'Commercial (démo)', email: 'moi@exemple.fr' }),
    signIn() { return this.account(); },
    signOut() {},
    calendars: async () => PEOPLE.map((p) => ({ ...p, canEdit: true })),
    async events(email, fromWall, toWall) {
      const i = PEOPLE.findIndex((p) => p.email === email);
      const out = [];
      for (let d = wallToMin(fromWall.slice(0, 10) + 'T00:00'); d < wallToMin(toWall); d += 1440) {
        const r = rand(d / 1440 * 7 + i * 101);
        if (r() < 0.06) { // journée de congés
          out.push({ start: minToWall(d), end: minToWall(d + 1440), showAs: 'oof' });
          continue;
        }
        const n = 2 + Math.floor(r() * 4);
        for (let k = 0; k < n; k++) {
          const s = d + 9 * 60 + Math.floor(r() * 16) * 30;
          out.push({ start: minToWall(s), end: minToWall(s + (r() < 0.6 ? 60 : 90)), showAs: 'busy' });
        }
      }
      return out.concat(created.filter((c) => c.email === email));
    },
    async createEvent(email, ev) {
      await new Promise((r) => setTimeout(r, 600));
      created.push({ email, start: ev.start, end: ev.end, showAs: 'busy' });
      return { webLink: null, teams: true };
    },
  };
}
