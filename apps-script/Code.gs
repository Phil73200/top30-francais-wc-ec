/**
 * Top 30 Français WC/EC — mise à jour automatique depuis db.fisski.com
 *
 * Réglages (Paramètres du projet > Propriétés du script) :
 *   DB_HOST  ex. db.fisski.com
 *   DB_PORT  ex. 3306
 *   DB_NAME  FISmainDB (majuscules exactes)
 *   DB_USER  ex. ffsuser
 *   DB_PASS  (à saisir par Phil)
 *
 * Fonctions :
 *   majTop30()            lance la mise à jour (1 seule connexion à la base)
 *   installerDeclencheur() crée le déclencheur quotidien (6 h)
 */

const ONGLET_DONNEES = 'donnees';
const ONGLET_JOURNAL = 'journal';
const PREMIERE_SAISON = 2022;
const HEURE_MAJ = 6;

function saisonEnCours_() {
  const d = new Date();
  // juillet à décembre => saison de l'année suivante
  return d.getMonth() >= 6 ? d.getFullYear() + 1 : d.getFullYear();
}

function requete_(saisonFin) {
  return `
WITH departs AS (
    SELECT rac.Seasoncode, res.Competitorid, rac.Catcode, rac.Disciplinecode,
           CASE WHEN res.Status = 'QLF'
                 AND CAST(res.Position AS UNSIGNED) BETWEEN 1 AND 30
                THEN 1 ELSE 0 END AS top30
    FROM A_resultal res
    JOIN A_raceal   rac ON rac.Raceid = res.Raceid
    WHERE rac.Seasoncode BETWEEN ${PREMIERE_SAISON} AND ${saisonFin}
      AND rac.Catcode IN ('WC', 'EC')
      AND rac.Disciplinecode IN ('GS', 'SL')
      AND res.Status NOT LIKE 'DNS%'
)
SELECT d.Seasoncode AS Saison, c.Gender AS Sexe, c.Lastname AS Nom, c.Firstname AS Prenom,
       YEAR(c.Birthdate) AS Annee, d.Disciplinecode AS Discipline,
       SUM(d.Catcode = 'WC')                 AS Departs_WC,
       SUM(d.Catcode = 'WC' AND d.top30 = 1) AS Top30_WC,
       SUM(d.Catcode = 'EC')                 AS Departs_EC,
       SUM(d.Catcode = 'EC' AND d.top30 = 1) AS Top30_EC
FROM departs d
JOIN A_competitor c ON c.Competitorid = d.Competitorid
WHERE c.Nationcode = 'FRA'
GROUP BY d.Seasoncode, c.Competitorid, c.Gender, c.Lastname, c.Firstname, c.Birthdate, d.Disciplinecode
ORDER BY d.Seasoncode, c.Gender, c.Lastname, c.Firstname, d.Disciplinecode`;
}

function majTop30() {
  const ss = SpreadsheetApp.getActive();
  const debut = new Date();
  let conn;
  try {
    const p = PropertiesService.getScriptProperties().getProperties();
    ['DB_HOST', 'DB_NAME', 'DB_USER', 'DB_PASS'].forEach(k => {
      if (!p[k]) throw new Error('Propriété manquante : ' + k);
    });
    const url = `jdbc:mysql://${p.DB_HOST}:${p.DB_PORT || 3306}/${p.DB_NAME}`;
    conn = Jdbc.getConnection(url, p.DB_USER, p.DB_PASS);
    const stmt = conn.createStatement();
    const rs = stmt.executeQuery(requete_(saisonEnCours_()));
    const meta = rs.getMetaData();
    const n = meta.getColumnCount();
    const lignes = [[]];
    for (let i = 1; i <= n; i++) lignes[0].push(meta.getColumnLabel(i));
    while (rs.next()) {
      const l = [];
      for (let i = 1; i <= n; i++) {
        const v = rs.getString(i);
        l.push(/^\d+$/.test(v || '') ? Number(v) : (v || ''));
      }
      lignes.push(l);
    }
    rs.close(); stmt.close();

    if (lignes.length < 2) throw new Error('Requête vide : aucune ligne, données précédentes conservées');

    let sh = ss.getSheetByName(ONGLET_DONNEES) || ss.insertSheet(ONGLET_DONNEES);
    sh.clearContents();
    sh.getRange(1, 1, lignes.length, n).setValues(lignes);
    journal_(ss, debut, 'OK', (lignes.length - 1) + ' lignes');
  } catch (e) {
    journal_(ss, debut, 'ERREUR', String(e && e.message || e));
    throw e;
  } finally {
    if (conn) try { conn.close(); } catch (e) {}
  }
}

function journal_(ss, date, statut, detail) {
  let sh = ss.getSheetByName(ONGLET_JOURNAL);
  if (!sh) {
    sh = ss.insertSheet(ONGLET_JOURNAL);
    sh.appendRow(['Date', 'Statut', 'Détail']);
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 3).setValues([[date, statut, detail]]);
  if (sh.getLastRow() > 200) sh.deleteRows(201, sh.getLastRow() - 200);
}

function installerDeclencheur() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'majTop30')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('majTop30').timeBased().everyDays(1).atHour(HEURE_MAJ).create();
}

/** Fournit l'onglet "donnees" en CSV au graphique (adresse web /exec) */
function doGet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ONGLET_DONNEES);
  const v = sh ? sh.getDataRange().getDisplayValues() : [];
  const csv = v.map(r => r.map(x => /[",\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x).join(',')).join('\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}
