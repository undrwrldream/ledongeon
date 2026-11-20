/**
 * Le Donjon — mémoire de la partie dans Google Sheets.
 *
 * La page du lanceur de dés lit et écrit l'état de la partie ici.
 * - Onglet « _donnees » : l'état complet (ne pas modifier à la main).
 * - Onglets « Partie », « Personnages », « Journal » : version lisible, mise à jour à chaque action
 *   (pratique pour l'IA maître du jeu qui a accès au Drive).
 *
 * Déploiement : Déployer → Nouveau déploiement → Application Web
 *   Exécuter en tant que : Moi · Qui a accès : Tout le monde
 */

const DONNEES = '_donnees';
const MORCEAU = 40000;          // une cellule accepte 50 000 caractères

function onglet_(nom) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nom) || ss.insertSheet(nom);
}
function version_() {
  return Number(PropertiesService.getScriptProperties().getProperty('v')) || 0;
}
function lire_() {
  const sh = onglet_(DONNEES);
  const n = sh.getLastRow();
  let json = '';
  if (n >= 2) json = sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).replace(/^~/, '')).join('');
  return json ? JSON.parse(json) : null;
}
function ecrire_(v, etat) {
  const sh = onglet_(DONNEES);
  const json = JSON.stringify(etat);
  const parts = [];
  for (let i = 0; i < json.length; i += MORCEAU) parts.push(['~' + json.slice(i, i + MORCEAU)]);   // « ~ » empêche Sheets d'y voir un nombre ou une formule
  sh.getRange('A:B').clearContent();
  sh.getRange('A1:B1').setValues([['État de la partie (ne pas modifier)', v]]);
  if (parts.length) {
    const r = sh.getRange(2, 1, parts.length, 1);
    r.setNumberFormat('@');
    r.setValues(parts);
  }
  PropertiesService.getScriptProperties().setProperty('v', String(v));
}
function sortie_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const v = version_();
  if (p.v !== undefined && Number(p.v) === v) return sortie_({ ok: true, v: v });   // rien de neuf
  return sortie_({ ok: true, v: v, state: v ? lire_() : null });
}

function doPost(e) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const v = version_();
    if (req.a !== 'save') return sortie_({ ok: false, error: 'action inconnue' });
    if (Number(req.v) !== v) return sortie_({ ok: false, conflict: true, v: v, state: v ? lire_() : null });
    ecrire_(v + 1, req.state);
    try { lisible_(req.state); } catch (err) { console.error(err); }
    return sortie_({ ok: true, v: v + 1 });
  } finally {
    verrou.releaseLock();
  }
}

/* ---------- Version lisible pour les humains et l'IA ---------- */
function texte_(html) { return String(html || '').replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"'); }
function remplir_(nom, lignes, largeur) {
  const sh = onglet_(nom);
  sh.clearContents();
  if (!lignes.length) return;
  const vals = lignes.map(l => { const r = l.slice(0, largeur); while (r.length < largeur) r.push(''); return r; });
  sh.getRange(1, 1, vals.length, largeur).setValues(vals);
  sh.getRange(1, 1, 1, largeur).setFontWeight('bold');
}
function lisible_(s) {
  const joueurs = s.players || [];
  const ordre = (s.order || []).map(i => joueurs[i]).filter(Boolean);
  const actuel = s.combat ? ordre[s.turn] : null;
  const ecran = s.screen || {};
  remplir_('Partie', [
    ['Résumé de la partie (mis à jour automatiquement à chaque action)', ''],
    ['Mis à jour', new Date()],
    ['Combat en cours', s.combat ? 'Oui' : 'Non'],
    ['Round', s.combat ? s.round : '—'],
    ['À qui le tour', actuel ? actuel.name : '—'],
    ['Dernier résultat', [texte_(ecran.label), texte_(ecran.total), texte_(ecran.detail)].filter(Boolean).join(' · ')]
  ], 2);
  const dmg = p => p.dmgN ? p.dmgN + 'd' + p.dmgF + (p.dmgM ? (p.dmgM > 0 ? '+' : '') + p.dmgM : '') : '—';
  remplir_('Personnages', [['Nom', 'Camp', 'PV', 'PV max', 'Attaque', 'Défense', 'Dégâts', 'Déplacement', 'Position (col, rang)', 'État']]
    .concat(joueurs.map(p => {
      const mob = p.side === 'mechant';
      return [p.name, mob ? 'Monstre' : 'Héros', p.hpMax ? p.hp : '—', p.hpMax || '—', p.mod, p.def == null ? '—' : p.def, dmg(p), p.spd || 3,
        p.x == null ? '—' : (p.x + 1) + ', ' + (p.y + 1), p.dead ? (mob ? 'Vaincu' : 'À terre') : (p.in ? 'Au combat' : 'Absent')];
    })), 10);
  const lignes = [['Round', 'Événement']];
  (s.log || []).forEach(e => {
    let t = '';
    if (e.t === 'atk') t = e.a + ' attaque ' + (e.tg || '(sans cible)') + ' : ' + e.total + (e.def != null ? ' contre Déf ' + e.def : '') + ' → ' + (e.hit === true ? 'TOUCHÉ' : e.hit === false ? 'RATÉ' : '—') + (e.nat === 20 ? ' (critique !)' : e.nat === 1 ? ' (échec critique)' : '');
    if (e.t === 'dmg') t = e.a + ' inflige ' + e.n + ' dégâts à ' + (e.tg || '?') + (e.hp != null ? ' (reste ' + e.hp + '/' + e.hpMax + ' PV)' : '');
    if (e.t === 'move') t = e.a + ' se déplace de ' + e.n + ' case(s) vers ' + e.tg + (e.contact ? ' et arrive au contact' : '');
    if (e.t === 'skip') t = e.a + ' passe son tour';
    if (e.t === 'kill') t = '💀 ' + e.tg + ' est vaincu par ' + e.a + (e.line ? ' — « ' + e.line + ' »' : '');
    if (e.t === 'down') t = '🩸 ' + e.tg + ' tombe au combat (frappé par ' + e.a + ')';
    if (e.t === 'rebirth') t = '🐉 ' + e.tg + ' se transforme en ' + e.into;
    if (e.t === 'free') t = '🎲 ' + e.a + ' : ' + e.lbl + ' = ' + e.total;
    if (t) lignes.push([e.r ? e.r : 'hors combat', t]);
  });
  remplir_('Journal', lignes, 2);
}
