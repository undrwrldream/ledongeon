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
 * Mise à jour : Déployer → Gérer les déploiements → ✎ → Version : Nouvelle version → Déployer (l'adresse /exec reste la même).
 *
 * IA (facultatif) : Paramètres du projet (⚙) → Propriétés du script → ajouter
 *   GEMINI_CLE = ta clé de Google AI Studio (commence par AIza…)
 * L'IA lit les documents « Consignes IA — Donjon Gaspésien » et « Univers — Donjon Gaspésien » (dossier le donjon).
 * Les images vont dans « le donjon / Images ».
 */

const DONNEES = '_donnees';
const MORCEAU = 40000;          // une cellule accepte 50 000 caractères

const FICHIER_PARTIE = '11hOjp_F1Zkh6hB7qq-4jvYLHbYvhemQtjJYTgR5LFY0';   // « Le Donjon — partie », si le script n'est rattaché à aucun fichier
function classeur_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('FICHIER_PARTIE') || FICHIER_PARTIE);
}
function onglet_(nom) {
  const ss = classeur_();
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
  if (p.a === 'images') return sortie_({ ok: true, images: images_() });
  const v = version_();
  const out = (p.v !== undefined && Number(p.v) === v) ? { ok: true, v: v } : { ok: true, v: v, state: v ? lire_() : null };
  if (p.c !== undefined) out.chat = discussion_(p.c);
  return sortie_(out);
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return sortie_({ ok: false, error: 'requête illisible' }); }
  try {
    if (req.a === 'save') return sortie_(sauver_(req));
    if (req.a === 'dire') return sortie_(dire_(req));
    if (req.a === 'ia') return sortie_(demanderIA_(req));
    if (req.a === 'publier') return sortie_(publier_(req));
    if (req.a === 'narrer') return sortie_(narrer_(req));
    return sortie_({ ok: false, error: 'action inconnue' });
  } catch (err) {
    return sortie_({ ok: false, error: String(err && err.message || err) });
  }
}

function sauver_(req) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    const v = version_();
    if (Number(req.v) !== v) return { ok: false, conflict: true, v: v, state: v ? lire_() : null };
    ecrire_(v + 1, req.state);
    try { lisible_(req.state); } catch (err) { console.error(err); }
    return { ok: true, v: v + 1 };
  } finally {
    verrou.releaseLock();
  }
}

/* ---------- Discussion (onglet « Discussion ») ---------- */
function discussion_(depuis) {
  const sh = onglet_('Discussion');
  const n = sh.getLastRow();
  if (n < 2) return [];
  let from = Number(depuis) < 0 ? Math.max(2, n - 59) : Number(depuis) + 2;   // première visite : les 60 derniers
  if (!(from >= 2)) from = 2;
  if (from > n) return [];
  return sh.getRange(from, 1, n - from + 1, 6).getValues().map(r => ({
    id: Number(r[0]), t: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    a: String(r[2]), k: String(r[3]), x: String(r[4]), img: String(r[5] || '')
  }));
}
function ajouter_(auteur, genre, texte, image) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    const sh = onglet_('Discussion');
    if (sh.getLastRow() === 0) { sh.appendRow(['#', 'Date', 'Auteur', 'Genre', 'Message', 'Image']); sh.setFrozenRows(1); }
    const id = sh.getLastRow();
    const sur = x => { x = String(x || ''); return /^[=+\-@]/.test(x) ? "'" + x : x; };   // jamais une formule
    sh.appendRow([id, new Date(), sur(auteur), genre, sur(texte), image || '']);
    return id;
  } finally {
    verrou.releaseLock();
  }
}
function dire_(req) {
  const texte = String(req.texte || '').trim().slice(0, 1500);
  if (!texte) return { ok: false, error: 'message vide' };
  const genre = req.genre === 'mj' ? 'mj' : 'joueur';
  return { ok: true, id: ajouter_(String(req.auteur || 'Joueur').slice(0, 40), genre, texte) };
}

/* ---------- IA maître du jeu (Gemini) ---------- */
const PROPS = PropertiesService.getScriptProperties();
const MODELES_TEXTE = ['gemini-3.6-flash', 'gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-2.5-flash'];
const MODELES_IMAGE = ['gemini-3.1-flash-lite-image', 'gemini-3.1-flash-image', 'gemini-2.5-flash-image'];
const CONSIGNES_DEFAUT = [
  "Tu es le maître du jeu d'une partie de Donjons et Dragons 5e entre amis québécois : « Le Donjon Gaspésien ».",
  "Écris en français québécois, avec l'humour et le ton de l'univers (téléréalité cosmique, Gaspésie, chantier, débrouillardise).",
  "Le lanceur de dés calcule les règles : dés, PV, initiative, positions et distances. Ces données sont EXACTES. Ne les change jamais, n'invente jamais un résultat de dé, un nombre de PV ou une distance.",
  "Ne décide jamais à la place des joueurs. Termine chaque intervention en leur laissant la main (ce qu'ils voient, ce qu'ils peuvent faire).",
  "Quand plusieurs monstres jouent d'affilée, tu reçois toutes leurs actions ensemble : raconte-les en UNE seule description, dans l'ordre où elles se sont passées.",
  "Garde tes interventions courtes : 3 à 8 phrases, sauf pour décrire une nouvelle scène."
].join('\n');

function cle_() {
  const k = PROPS.getProperty('GEMINI_CLE');
  if (!k) throw new Error("Clé Gemini absente : ajoute GEMINI_CLE dans Paramètres du projet → Propriétés du script.");
  return k.trim();
}
function hash_(s) {
  let h = 5381;
  for (const c of String(s)) h = ((h << 5) + h + c.codePointAt(0)) | 0;
  return 'h' + (h >>> 0).toString(36);
}
function verifierMJ_(code) { /* plus de code MJ : le bouton « Je suis le MJ » suffit */ }
function texteOnglet_(nom, max) {
  const sh = classeur_().getSheetByName(nom);
  if (!sh || sh.getLastRow() === 0) return '';
  return sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0])).join('\n').slice(0, max || 40000);
}
function gemini_(modeles, cleProp, corps) {
  const pref = PROPS.getProperty(cleProp);
  const liste = pref ? [pref].concat(modeles.filter(m => m !== pref)) : modeles;
  let derniere = '';
  for (const m of liste) {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': cle_() }, payload: JSON.stringify(corps)
    });
    const code = r.getResponseCode(), txt = r.getContentText();
    if (code === 200) { PROPS.setProperty(cleProp, m); return JSON.parse(txt); }
    derniere = code + ' ' + txt.slice(0, 300);
    if (code === 404 || (code === 400 && /not found|not supported|unknown name/i.test(txt))) continue;   // modèle retiré : on essaie le suivant
    if (code === 429) throw new Error(/limit: 0/.test(txt) ? 'Ce modèle demande la facturation Google (pas offert gratuitement).' : 'Limite de Gemini atteinte, réessaie dans une minute.');
    if (code === 400 && /API key/i.test(txt)) throw new Error('Clé Gemini refusée : vérifie GEMINI_CLE (elle commence par AIza).');
    throw new Error('Erreur Gemini ' + derniere);
  }
  throw new Error('Aucun modèle Gemini disponible : ' + derniere);
}
function texteDe_(rep) {
  const c = rep.candidates && rep.candidates[0];
  return c && c.content && c.content.parts ? c.content.parts.map(p => p.text || '').join('') : '';
}
function historique_() {
  return discussion_(-1).slice(-30).filter(m => m.k !== 'image').map(m => m.a + ' : ' + m.x).join('\n');
}
const DOC_CONSIGNES = '1EIGD0Fly7R8LCQxRaZBGs7-Rr1c-UOljB1taU9F0oyk';   // « Consignes IA — Donjon Gaspésien » (modifiable par le MJ)
const DOC_UNIVERS = '12SA8sL_NComI_Pi9pyjezG7RD9TswO2LBETfWTb7J0s';     // « Univers — Donjon Gaspésien »
const DOC_BESTIAIRE = '1vw9alx8x7fKupJpMx4m_WXu79paDl6YPc7hFJQXp1o0';   // « Bestiaire des 30 — pour l'IA »
function texteDoc_(id, max) {
  try { return DocumentApp.openById(id).getBody().getText().slice(0, max); } catch (err) { return ''; }
}
function systeme_() {
  const consignes = texteDoc_(PROPS.getProperty('DOC_CONSIGNES') || DOC_CONSIGNES, 8000) || texteOnglet_('Consignes IA', 8000) || CONSIGNES_DEFAUT;
  const univers = texteDoc_(PROPS.getProperty('DOC_UNIVERS') || DOC_UNIVERS, 60000) || texteOnglet_('Univers', 40000);
  const bestiaire = texteDoc_(PROPS.getProperty('DOC_BESTIAIRE') || DOC_BESTIAIRE, 40000);
  return consignes + (univers ? '\n\n=== UNIVERS, PNJ ET CRÉATURES ===\n' + univers : '') + (bestiaire ? '\n\n=== BESTIAIRE DES 30 ===\n' + bestiaire : '');
}

function demanderIA_(req) {
  verifierMJ_(req.mj);
  const tache = [
    "Rédige la prochaine intervention du maître du jeu.",
    req.consigne ? 'Indication du MJ humain : ' + String(req.consigne).slice(0, 800) : '',
    "Si des monstres apparaissent, choisis-les EN PRIORITÉ dans le BESTIAIRE DISPONIBLE (nom exact) et indique combien. N'invente un nouveau monstre que si aucun ne convient.",
    'Réponds en JSON : {"narration": texte, "scene": description visuelle courte de la scène si une nouvelle scène commence ou si le décor change, sinon null, ' +
    '"monstres": liste des monstres qui ARRIVENT maintenant dans ta narration (sinon []), chacun {"nom","nombre"} pour un monstre du bestiaire, ' +
    'ou {"nom","nombre","apparence","pv","defense","attaque","degats" (ex. "2d6+3"),"deplacement" (en cases de 1,5 m)} pour un nouveau monstre}.'
  ].filter(Boolean).join('\n');
  const rep = gemini_(MODELES_TEXTE, 'MODELE_TEXTE', {
    systemInstruction: { parts: [{ text: systeme_() }] },
    contents: [{ role: 'user', parts: [{ text: 'ÉTAT DU JEU (calculé par le lanceur, exact) :\n' + String(req.contexte || '').slice(0, 12000) +
      '\n\nDISCUSSION RÉCENTE :\n' + historique_() + '\n\n' + tache }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.9 }
  });
  let b;
  try { b = JSON.parse(texteDe_(rep)); } catch (err) { b = { narration: texteDe_(rep), scene: null, monstres: [] }; }
  return { ok: true, brouillon: { narration: String(b.narration || ''), scene: b.scene || null, monstres: Array.isArray(b.monstres) ? b.monstres.slice(0, 6) : [] } };
}

function narrer_(req) {
  const derniere = Number(PROPS.getProperty('derniereNarration')) || 0;
  if (Date.now() - derniere < 15000) return { ok: true, saute: true };     // pas deux narrations en même temps
  PROPS.setProperty('derniereNarration', String(Date.now()));
  const rep = gemini_(MODELES_TEXTE, 'MODELE_TEXTE', {
    systemInstruction: { parts: [{ text: systeme_() }] },
    contents: [{ role: 'user', parts: [{ text: 'ÉTAT DU JEU (calculé par le lanceur, exact) :\n' + String(req.contexte || '').slice(0, 12000) +
      '\n\nDISCUSSION RÉCENTE :\n' + historique_() +
      '\n\nRaconte en une seule description les actions des monstres ci-dessus (la séquence qui vient de se terminer), en respectant exactement les résultats. Termine en annonçant à qui c\'est le tour. Texte seulement, pas de JSON.' }] }],
    generationConfig: { temperature: 0.9 }
  });
  const texte = texteDe_(rep).trim();
  if (!texte) return { ok: false, error: 'réponse vide' };
  return { ok: true, id: ajouter_('IA du Donjon', 'ia', texte) };
}

/* ---------- Images (modèle d'images Gemini : payant) ---------- */
const STYLE = "Illustration de bande dessinée détaillée, encrage marqué, couleurs saturées, ambiance industrielle souterraine sombre et humoristique, univers « Donjon Gaspésien ».";
const DOSSIER_IMAGES = '1yNjt6TM1h40ic_f9gk2-Agyk3bJ9VPlv';             // « le donjon / Images »
function dossierImages_() {
  const id = PROPS.getProperty('DOSSIER_IMAGES') || DOSSIER_IMAGES;
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const d = DriveApp.createFolder('Le Donjon — images');
  PROPS.setProperty('DOSSIER_IMAGES', d.getId());
  return d;
}
function image_(prompt, nom) {
  const rep = gemini_(MODELES_IMAGE, 'MODELE_IMAGE', {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { responseModalities: ['IMAGE'] }
  });
  const parts = (rep.candidates && rep.candidates[0] && rep.candidates[0].content && rep.candidates[0].content.parts) || [];
  const d = parts.find(p => p.inlineData && p.inlineData.data);
  if (!d) throw new Error("Gemini n'a pas renvoyé d'image.");
  const blob = Utilities.newBlob(Utilities.base64Decode(d.inlineData.data), d.inlineData.mimeType || 'image/png', nom + '.png');
  const f = dossierImages_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1000';
}
/* Les images du dossier, par nom de fichier (ex. ratus_web.png), partagées « tous ceux qui ont le lien » pour que la page les affiche */
function images_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('images');
  if (c) return JSON.parse(c);
  const out = {};
  const it = dossierImages_().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!/^image\//.test(f.getMimeType())) continue;
    try { if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
    out[f.getName()] = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600';
  }
  cache.put('images', JSON.stringify(out), 600);
  return out;
}
function publier_(req) {
  verifierMJ_(req.mj);
  const texte = String(req.texte || '').trim();
  if (texte) ajouter_('IA du Donjon', 'ia', texte.slice(0, 6000));
  const out = { ok: true, imgs: {}, erreurs: [] };
  if (!req.images) return out;
  if (req.scene) {                                   // la scène d'ambiance d'abord, puis les monstres
    try {
      const url = image_(STYLE + ' Scène large, sans texte : ' + req.scene, 'scene-' + Date.now());
      ajouter_('IA du Donjon', 'image', String(req.scene).slice(0, 300), url);
      out.scene = url;
    } catch (err) { out.erreurs.push('Scène : ' + err.message); }
  }
  (req.monstres || []).slice(0, 6).forEach(m => {
    try {
      out.imgs[m.nom] = image_(STYLE + ' Portrait en pied d\'un seul personnage, centré, entier, sur fond uni sombre, sans texte, style pion de jeu : ' +
        m.nom + '. ' + (m.apparence || ''), 'monstre-' + m.nom);
    } catch (err) { out.erreurs.push(m.nom + ' : ' + err.message); }
  });
  return out;
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
