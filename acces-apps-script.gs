/**
 * Contrôle des codes d’accès de l’appli « Marées de Saint-Malo ».
 * Feuille « Codes » : A Code | B Nom | C Actif (OUI/NON) | D Expiration (date, facultatif)
 *                     E Appareil (rempli tout seul) | F Dernière connexion | G Nb connexions
 * Un code se verrouille sur le premier téléphone qui l’utilise (colonne E).
 * Pour changer de téléphone : efface la case E de la personne. Pour couper un accès : mets NON en C.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var res = verifier_(String(p.code || ""), String(p.device || ""));
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function verifier_(code, device) {
  code = code.trim().toUpperCase();
  if (!code || !device) return {ok: false, message: "Code manquant."};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Codes");
    if (!sh) return {ok: false, message: "Feuille « Codes » introuvable."};
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() !== code) continue;
      var actif = String(rows[i][2]).trim().toUpperCase();
      if (actif !== "OUI") return {ok: false, message: "Ce code est désactivé."};
      var exp = rows[i][3];
      if (exp instanceof Date && exp.getTime() < new Date().setHours(0, 0, 0, 0)) return {ok: false, message: "Ce code a expiré."};
      var app = String(rows[i][4]).trim();
      if (app && app !== device) return {ok: false, message: "Ce code est déjà utilisé sur un autre appareil."};
      var r = i + 1;
      if (!app) sh.getRange(r, 5).setValue(device);
      sh.getRange(r, 6).setValue(new Date());
      sh.getRange(r, 7).setValue((Number(rows[i][6]) || 0) + 1);
      return {ok: true, nom: String(rows[i][1])};
    }
    return {ok: false, message: "Code inconnu."};
  } finally {
    lock.releaseLock();
  }
}
