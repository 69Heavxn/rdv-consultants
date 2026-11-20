// Titre du RDV et nom du client.

// ("Consultation Société", { civilite: 'Madame', nom: 'Jeanne DUPONT' })
// -> "Consultation Société - Madame Jeanne DUPONT"
export function makeTitle(prefix, { civilite, nom }) {
  const who = [civilite, nom].map((s) => String(s || '').replace(/\s+/g, ' ').trim()).filter(Boolean).join(' ');
  return [String(prefix || '').trim(), who].filter(Boolean).join(' - ');
}

const CIV = [
  [/^(madame|mme\.?|mademoiselle|mlle\.?|mrs\.?|ms\.?)(\s+|$)/i, 'Madame'],
  [/^(monsieur|m\.|mr\.?)(\s+|$)/i, 'Monsieur'],
];
// "Mme Jeanne DUPONT" -> { civilite: 'Madame', nom: 'Jeanne DUPONT' }
export function parseName(raw) {
  let nom = String(raw || '').replace(/\s+/g, ' ').trim();
  for (const [re, civilite] of CIV) {
    if (re.test(nom)) return { civilite, nom: nom.replace(re, '').trim() };
  }
  return { civilite: '', nom };
}
// valeur d'un champ "Civilité" -> 'Madame' | 'Monsieur' | ''
export const civiliteOf = (raw) => parseName(String(raw || '') + ' x').civilite;
