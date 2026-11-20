// Favori « RDV » : sur la fiche Salesforce affichée, lit Madame/Monsieur et le
// nom du client, puis ouvre la page de prise de RDV avec. Rien d'autre.
import { parseName, civiliteOf } from './text.js';

const RDV_URL = '__RDV_URL__'; // mis par la page Réglages au moment d'installer le favori

// tous les éléments de la page, y compris dans les composants (shadow DOM)
function* elements(root) {
  const stack = [root];
  while (stack.length) {
    const walker = document.createTreeWalker(stack.pop(), NodeFilter.SHOW_ELEMENT);
    for (let el = walker.currentNode; el; el = walker.nextNode()) {
      if (el.nodeType !== 1) continue;
      yield el;
      if (el.shadowRoot) stack.push(el.shadowRoot);
    }
  }
}
// Salesforce garde les fiches des autres onglets dans la page : on ne lit que ce qui est visible
const visible = (el) => {
  try { return el.checkVisibility ? el.checkVisibility({ visibilityProperty: true }) : el.getClientRects().length > 0; } catch { return true; }
};
const text = (el) => (el.textContent || '').replace(/\s+/g, ' ').trim();
const parent = (el) => el.parentElement || (el.parentNode && el.parentNode.host) || null;

// nom en tête de fiche, ex. "Mme Jeanne DUPONT"
function readName() {
  const sels = ['records-highlights2 lightning-formatted-name', 'records-highlights2 h1', 'force-highlights2 h1',
    '.slds-page-header__title lightning-formatted-name', 'h1.slds-page-header__title', '.slds-page-header__title'];
  for (const sel of sels) {
    for (const el of elements(document)) {
      if (el.matches(sel) && visible(el) && text(el) && text(el).length < 80) return text(el);
    }
  }
  const t = document.title.split('|')[0].trim();
  return /salesforce/i.test(t) ? '' : t;
}

// champ "Civilité" de la fiche, si le nom en tête ne la donne pas
function readCiviliteField() {
  for (const el of elements(document)) {
    if (!el.matches('.slds-form-element__label, .test-id__field-label, label, dt') || !visible(el)) continue;
    const lbl = text(el);
    if (!/^(civilit[ée]|salutation)$/i.test(lbl)) continue;
    for (let box = parent(el); box; box = parent(box)) {
      if (box.matches && box.matches('.slds-form-element, records-record-layout-item, force-record-layout-item, .slds-form__item')) {
        const c = civiliteOf(text(box).replace(lbl, '').trim());
        if (c) return c;
        break;
      }
    }
  }
  return '';
}

(function main() {
  const { civilite, nom } = parseName(readName());
  const client = { civilite: civilite || readCiviliteField(), nom };
  const w = window.open(RDV_URL + '#client=' + encodeURIComponent(JSON.stringify(client)), 'rdv-consultants');
  if (!w) alert('RDV : la fenêtre a été bloquée. Autorise les pop-ups pour Salesforce puis reclique.');
})();
