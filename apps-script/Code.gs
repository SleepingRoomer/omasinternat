/**
 * OMAS Internat : réception des choix de menu dans ce Google Sheet.
 * La page du formulaire est hébergée sur GitHub Pages ; ce script lui sert de « boîte aux lettres ».
 *
 * Tout se règle dans le tableau, pas ici :
 *   onglet « Réglages » : nom de l'événement, date limite, choix ouverts ou non
 *   onglet « Menu »     : une ligne par plat proposé
 *   onglet « Réponses » : rempli automatiquement
 *   onglet « Récap »    : le décompte à transmettre au restaurant, mis à jour à chaque réponse
 */

const ONGLET_REGLAGES = 'Réglages';
const ONGLET_MENU = 'Menu';
const ONGLET_REPONSES = 'Réponses';
const ONGLET_RECAP = 'Récap';
const COL_DATE = 'Dernière mise à jour';
const COL_NOM = 'Prénom et nom';
const COL_REMARQUES = 'Allergies et remarques';

/** À lancer une fois depuis l'éditeur : crée les onglets et demande les autorisations. */
function installer() {
  initialiser_();
  majRecap_();
}

/** La page demande le menu et les réglages. */
function doGet() {
  initialiser_();
  return json_(lireConfig_());
}

/** La page envoie un choix de menu validé. */
function doPost(e) {
  try {
    const donnees = JSON.parse(e && e.postData && e.postData.contents || '{}');
    const resultat = enregistrer(donnees);
    return json_({ ok: true, misAJour: resultat.misAJour });
  } catch (err) {
    return json_({ ok: false, erreur: err.message });
  }
}

function json_(objet) {
  return ContentService.createTextOutput(JSON.stringify(objet)).setMimeType(ContentService.MimeType.JSON);
}

function enregistrer(donnees) {
  const config = lireConfig_();
  if (!config.ouvert) throw new Error('les choix sont clos');

  const nom = String(donnees && donnees.nom || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  if (nom.length < 2) throw new Error('nom manquant');

  const choix = {};
  config.etapes.forEach(function (e) {
    const valeur = donnees.choix && donnees.choix[e.titre];
    const ok = e.options.some(function (o) { return o.titre === valeur; });
    if (!ok) throw new Error('choix manquant pour « ' + e.titre + ' »');
    choix[e.titre] = valeur;
  });
  const remarques = config.remarques ? String(donnees.remarques || '').trim().slice(0, 500) : '';

  const verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    const feuille = SpreadsheetApp.getActive().getSheetByName(ONGLET_REPONSES);
    const entetes = [COL_DATE, COL_NOM].concat(config.etapes.map(function (e) { return e.titre; }));
    if (config.remarques) entetes.push(COL_REMARQUES);
    const colonnes = preparerEntetes_(feuille, entetes);

    const ligneValeurs = new Array(colonnes.total).fill('');
    ligneValeurs[colonnes.index[COL_DATE]] = new Date();
    ligneValeurs[colonnes.index[COL_NOM]] = securiser_(nom);
    config.etapes.forEach(function (e) { ligneValeurs[colonnes.index[e.titre]] = securiser_(choix[e.titre]); });
    if (config.remarques) ligneValeurs[colonnes.index[COL_REMARQUES]] = securiser_(remarques);

    // Si la même personne répond deux fois, on remplace sa ligne
    const cle = simplifier_(nom);
    let ligne = -1;
    const derniere = feuille.getLastRow();
    if (derniere >= 2) {
      const noms = feuille.getRange(2, colonnes.index[COL_NOM] + 1, derniere - 1, 1).getValues();
      for (let i = 0; i < noms.length; i++) {
        if (simplifier_(noms[i][0]) === cle) { ligne = i + 2; break; }
      }
    }
    if (ligne > 0) {
      // On garde les colonnes d'anciennes étapes éventuellement supprimées du menu
      const existante = feuille.getRange(ligne, 1, 1, colonnes.total).getValues()[0];
      for (let j = 0; j < colonnes.total; j++) {
        if (ligneValeurs[j] === '' && entetes.indexOf(colonnes.noms[j]) < 0) ligneValeurs[j] = existante[j];
      }
      feuille.getRange(ligne, 1, 1, colonnes.total).setValues([ligneValeurs]);
    } else {
      feuille.appendRow(ligneValeurs);
    }
    feuille.getRange(2, colonnes.index[COL_DATE] + 1, Math.max(feuille.getLastRow() - 1, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    majRecap_();
    return { misAJour: ligne > 0 };
  } finally {
    verrou.releaseLock();
  }
}

// ---------------------------------------------------------------------------

function lireConfig_() {
  const classeur = SpreadsheetApp.getActive();
  const reglages = {};
  classeur.getSheetByName(ONGLET_REGLAGES).getDataRange().getDisplayValues().slice(1).forEach(function (l) {
    if (l[0]) reglages[simplifier_(l[0])] = String(l[1]).trim();
  });
  const oui = function (v) { return ['oui', 'o', 'yes', 'vrai', 'true', '1'].indexOf(simplifier_(v)) >= 0; };

  const etapes = [];
  const parTitre = {};
  classeur.getSheetByName(ONGLET_MENU).getDataRange().getDisplayValues().slice(1).forEach(function (l) {
    const etape = String(l[0]).trim();
    const titre = String(l[1]).trim();
    if (!etape || !titre) return;
    if (!parTitre[etape]) { parTitre[etape] = { titre: etape, options: [] }; etapes.push(parTitre[etape]); }
    const option = { titre: titre };
    if (String(l[2]).trim()) option.description = String(l[2]).trim();
    parTitre[etape].options.push(option);
  });

  return {
    evenement: reglages[simplifier_("Nom de l'événement")] || 'Événement OMAS',
    dateLimite: reglages[simplifier_('Date limite')] || '',
    ouvert: oui(reglages[simplifier_('Choix ouverts')] || 'OUI') && etapes.length > 0,
    remarques: oui(reglages[simplifier_('Demander les allergies')] || 'OUI'),
    etapes: etapes
  };
}

function initialiser_() {
  const classeur = SpreadsheetApp.getActive();
  if (!classeur.getSheetByName(ONGLET_REGLAGES)) {
    const f = classeur.insertSheet(ONGLET_REGLAGES, 0);
    f.getRange(1, 1, 5, 2).setValues([
      ['Paramètre', 'Valeur'],
      ["Nom de l'événement", "Dîner de rentrée des internes"],
      ['Date limite', '15 novembre'],
      ['Choix ouverts', 'OUI'],
      ['Demander les allergies', 'OUI']
    ]);
    f.getRange('B2:B5').setNumberFormat('@');
    styliserEntete_(f, 2);
    f.setColumnWidth(1, 220); f.setColumnWidth(2, 260);
  }
  if (!classeur.getSheetByName(ONGLET_MENU)) {
    const f = classeur.insertSheet(ONGLET_MENU, 1);
    f.getRange(1, 1, 11, 3).setValues([
      ['Étape', 'Choix', 'Description (facultatif)'],
      ['Entrée', 'Soupe de lentilles', ''],
      ['Entrée', 'Salade orientale', ''],
      ['Entrée', 'Briouates', ''],
      ['Plat', "Tajine d'agneau", 'Pruneaux et amandes'],
      ['Plat', 'Poulet rôti', 'Pommes de terre au four'],
      ['Plat', 'Pavé de saumon', 'Riz et légumes'],
      ['Plat', 'Couscous végétarien', ''],
      ['Dessert', 'Fondant au chocolat', ''],
      ['Dessert', 'Salade de fruits', ''],
      ['Dessert', 'Pâtisseries orientales', '']
    ]);
    styliserEntete_(f, 3);
    f.setColumnWidth(1, 120); f.setColumnWidth(2, 220); f.setColumnWidth(3, 260);
  }
  if (!classeur.getSheetByName(ONGLET_REPONSES)) {
    const f = classeur.insertSheet(ONGLET_REPONSES, 2);
    f.getRange(1, 1, 1, 2).setValues([[COL_DATE, COL_NOM]]);
    styliserEntete_(f, 2);
  }
  if (!classeur.getSheetByName(ONGLET_RECAP)) classeur.insertSheet(ONGLET_RECAP, 3);
  // Supprime la feuille vide créée par défaut avec le classeur
  const vide = classeur.getSheetByName('Feuille 1') || classeur.getSheetByName('Sheet1');
  if (vide && classeur.getSheets().length > 1 && vide.getLastRow() === 0) classeur.deleteSheet(vide);
}

function preparerEntetes_(feuille, entetes) {
  let actuels = feuille.getRange(1, 1, 1, Math.max(feuille.getLastColumn(), 1)).getValues()[0].map(String);
  if (actuels.length === 1 && actuels[0] === '') actuels = [];
  entetes.forEach(function (h) {
    if (actuels.indexOf(h) < 0) {
      actuels.push(h);
      feuille.getRange(1, actuels.length).setValue(h);
    }
  });
  styliserEntete_(feuille, actuels.length);
  const index = {};
  actuels.forEach(function (h, i) { index[h] = i; });
  return { index: index, total: actuels.length, noms: actuels };
}

function majRecap_() {
  const classeur = SpreadsheetApp.getActive();
  const config = lireConfig_();
  const rep = classeur.getSheetByName(ONGLET_REPONSES).getDataRange().getValues();
  const entetes = rep[0].map(String);
  const lignes = rep.slice(1).filter(function (l) { return String(l[entetes.indexOf(COL_NOM)]).trim(); });

  const sortie = [[config.evenement, ''], [lignes.length + ' réponse(s)', ''], ['', '']];
  const titres = [];
  config.etapes.forEach(function (e) {
    const col = entetes.indexOf(e.titre);
    titres.push(sortie.length);
    sortie.push([e.titre, 'Nombre']);
    e.options.forEach(function (o) {
      const n = col < 0 ? 0 : lignes.filter(function (l) { return l[col] === o.titre; }).length;
      sortie.push([o.titre, n]);
    });
    sortie.push(['', '']);
  });
  const colRem = entetes.indexOf(COL_REMARQUES);
  if (colRem >= 0) {
    titres.push(sortie.length);
    sortie.push(['Allergies et remarques', 'Personne']);
    const avec = lignes.filter(function (l) { return String(l[colRem]).trim(); });
    if (!avec.length) sortie.push(['Aucune', '']);
    avec.forEach(function (l) { sortie.push([String(l[colRem]), String(l[entetes.indexOf(COL_NOM)])]); });
  }

  const f = classeur.getSheetByName(ONGLET_RECAP);
  f.clear();
  f.getRange(1, 1, sortie.length, 2).setValues(sortie);
  f.getRange('A1').setFontSize(14).setFontWeight('bold');
  f.getRange('A2').setFontColor('#5f6f6a');
  titres.forEach(function (r) { f.getRange(r + 1, 1, 1, 2).setFontWeight('bold').setBackground('#e3f3ec'); });
  f.setColumnWidth(1, 280); f.setColumnWidth(2, 180);
}

function styliserEntete_(feuille, nbColonnes) {
  feuille.getRange(1, 1, 1, nbColonnes).setFontWeight('bold').setBackground('#0f5a4e').setFontColor('#ffffff');
  feuille.setFrozenRows(1);
}

/** Empêche qu'une réponse commençant par = soit interprétée comme une formule. */
function securiser_(v) {
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function simplifier_(t) {
  return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}
