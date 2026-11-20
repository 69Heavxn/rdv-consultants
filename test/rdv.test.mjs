import test from 'node:test';
import assert from 'node:assert/strict';
import { daySlots, daysAhead, wallToDate, parisNow, isBusy } from '../src/slots.js';
import { makeTitle, parseName, civiliteOf } from '../src/text.js';

const rules = { start: '09:00', end: '18:00', step: 30, duration: 60, minNotice: 30 };

test('créneaux libres : un RDV occupe le consultant, pas les autres', () => {
  const busy = {
    a: [{ start: '2026-11-24T10:00:00.0000000', end: '2026-11-24T11:00:00.0000000', showAs: 'busy' }],
    b: [],
  };
  const slots = daySlots('2026-11-24', rules, busy, '2026-11-20T12:00');
  assert.equal(slots[0].start, '2026-11-24T09:00');
  assert.equal(slots.at(-1).start, '2026-11-24T17:00'); // dernier RDV d'1 h finit à 18h
  const at = (t) => slots.find((s) => s.start === '2026-11-24T' + t).free;
  assert.deepEqual(at('09:00'), { a: true, b: true });
  assert.deepEqual(at('09:30'), { a: false, b: true }); // 9h30-10h30 chevauche 10h-11h
  assert.deepEqual(at('10:30'), { a: false, b: true });
  assert.deepEqual(at('11:00'), { a: true, b: true });
});

test('créneaux : événements libres, annulés et journée entière', () => {
  const busy = {
    a: [{ start: '2026-11-24T00:00:00', end: '2026-11-25T00:00:00', showAs: 'oof' }], // congés
    b: [{ start: '2026-11-24T09:00:00', end: '2026-11-24T18:00:00', showAs: 'free' },
      { start: '2026-11-24T14:00:00', end: '2026-11-24T15:00:00', showAs: 'busy', isCancelled: true }],
  };
  const slots = daySlots('2026-11-24', rules, busy, '2026-11-20T12:00');
  assert.ok(slots.every((s) => !s.free.a));
  assert.ok(slots.every((s) => s.free.b));
  assert.equal(isBusy({ showAs: 'tentative' }), true);
});

test("créneaux : pas de RDV dans le passé ni dans les 30 prochaines minutes", () => {
  const slots = daySlots('2026-11-24', rules, { a: [] }, '2026-11-24T10:10');
  assert.equal(slots[0].start, '2026-11-24T11:00');
});

test('heure de Paris', () => {
  assert.equal(wallToDate('2026-11-24T10:00').toISOString(), '2026-11-24T09:00:00.000Z'); // hiver UTC+1
  assert.equal(wallToDate('2026-07-14T10:00').toISOString(), '2026-07-14T08:00:00.000Z'); // été UTC+2
  assert.equal(parisNow(new Date('2026-11-24T09:00:00Z')), '2026-11-24T10:00');
});

test('titre du RDV et nom lu sur la fiche', () => {
  assert.equal(makeTitle('Consultation Société', { civilite: 'Madame', nom: 'Jeanne DUPONT' }), 'Consultation Société - Madame Jeanne DUPONT');
  assert.equal(makeTitle(' Consultation Société ', { civilite: '', nom: ' Jean  DUPONT ' }), 'Consultation Société - Jean DUPONT');
  assert.deepEqual(parseName('Mme Jeanne DUPONT'), { civilite: 'Madame', nom: 'Jeanne DUPONT' });
  assert.deepEqual(parseName('M. Jean DUPONT'), { civilite: 'Monsieur', nom: 'Jean DUPONT' });
  assert.deepEqual(parseName('Monsieur Jean DUPONT'), { civilite: 'Monsieur', nom: 'Jean DUPONT' });
  assert.deepEqual(parseName('Mlle Jeanne DUPONT'), { civilite: 'Madame', nom: 'Jeanne DUPONT' });
  assert.deepEqual(parseName('Jeanne DUPONT'), { civilite: '', nom: 'Jeanne DUPONT' });
  assert.deepEqual(parseName('Marc DUPONT'), { civilite: '', nom: 'Marc DUPONT' }); // "M" de Marc n'est pas "M."
  assert.equal(civiliteOf('Mme'), 'Madame');
  assert.equal(civiliteOf('M.'), 'Monsieur');
  assert.equal(civiliteOf('--Aucun--'), '');
});

test("créneaux proposés : du lendemain jusqu'à dans 1 mois, jours ouvrés", () => {
  const d = daysAhead('2026-11-20T15:00', 1); // vendredi 20 novembre
  assert.equal(d[0], '2026-11-23'); // pas le jour même, pas le week-end
  assert.equal(d.at(-1), '2026-12-18'); // le 20 décembre est un dimanche
  assert.equal(d.length, 20);
  assert.equal(daysAhead('2027-01-31T10:00', 1).at(-1), '2027-02-26'); // fin de mois : 28 février est un dimanche
});
