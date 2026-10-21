/**
 * EQUACIONS V2 · 1r BATXILLERAT
 * Enganxa TOT aquest bloc AL FINAL DEL MATEIX FITXER on hi ha const JOCS,
 * dins el projecte d'Apps Script que ja fas servir amb Potències.
 * No substitueixis el codi existent i no creïs un projecte nou.
 * Després executa preparaEquacionsV2 i actualitza la implementació existent:
 * Implementa > Gestiona les implementacions > llapis > Nova versió > Implementa.
 * Es conserva la mateixa URL /exec i el mateix FULL_ID.
 */
JOCS['equacions-v2-1batx'] = {
  name: 'Equacions v2 · 1r Batxillerat',
  sheet: 'Equacions v2 1r Batx',
  ids: [
    'linear-0', 'linear-1', 'linear-2',
    'quadratic-0', 'quadratic-1', 'quadratic-2',
    'biquadratic-0', 'biquadratic-1', 'biquadratic-2',
    'polynomial-0', 'polynomial-1', 'polynomial-2',
    'radical-0', 'radical-1', 'radical-2',
    'rational-0', 'rational-1', 'rational-2',
    'logarithmic-0', 'logarithmic-1', 'logarithmic-2',
    'exponential-0', 'exponential-1', 'exponential-2'
  ],
  titles: [
    '1r grau · Fàcil', '1r grau · Mitjana', '1r grau · Difícil',
    '2n grau · Fàcil', '2n grau · Mitjana', '2n grau · Difícil',
    'Biquadrades · Fàcil', 'Biquadrades · Mitjana', 'Biquadrades · Difícil',
    'Polinòmiques · Fàcil', 'Polinòmiques · Mitjana', 'Polinòmiques · Difícil',
    'Amb radicals · Fàcil', 'Amb radicals · Mitjana', 'Amb radicals · Difícil',
    'Racionals · Fàcil', 'Racionals · Mitjana', 'Racionals · Difícil',
    'Logarítmiques · Fàcil', 'Logarítmiques · Mitjana', 'Logarítmiques · Difícil',
    'Exponencials · Fàcil', 'Exponencials · Mitjana', 'Exponencials · Difícil'
  ]
};
function preparaEquacionsV2() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = full();
    const cfg = JOCS['equacions-v2-1batx'];
    let sh = ss.getSheetByName(cfg.sheet);
    // 24 categories + correu, joc, mitjana i data = 28 columnes.
    // Un full nou acostuma a tenir-ne 26: les ampliem abans d'escriure.
    if (!sh) sh = ss.insertSheet(cfg.sheet);
    if (sh.getMaxColumns() < cfg.ids.length + 4) {
      sh.insertColumnsAfter(sh.getMaxColumns(), cfg.ids.length + 4 - sh.getMaxColumns());
    }
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Correu', 'Joc', ...cfg.titles, 'Mitjana (màxim 100)', 'Actualitzat']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, cfg.ids.length + 4)
        .setBackground('#203350').setFontColor('white').setFontWeight('bold');
      sh.getRange(2, 3, Math.max(1, sh.getMaxRows() - 1), cfg.ids.length)
        .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1000000).setAllowInvalid(false).build())
        .setBackground('#e8f5e9');
    }
    const headers = sh.getRange(1, 1, 1, cfg.ids.length + 4).getValues()[0];
    if (headers[0] !== 'Correu' || headers[1] !== 'Joc' || cfg.titles.some((t, i) => headers[i + 2] !== t)) {
      throw Error('Les capçaleres d’Equacions v2 no coincideixen. No s’han modificat les puntuacions.');
    }
    preparaPerfils(ss);
    let games = ss.getSheetByName('Jocs');
    if (!games) {
      games = ss.insertSheet('Jocs');
      games.appendRow(['ID', 'Títol']);
      games.setFrozenRows(1);
    }
    if (!games.getDataRange().getValues().some(r => r[0] === 'equacions-v2-1batx')) {
      games.appendRow(['equacions-v2-1batx', cfg.name]);
    }
    console.log('Equacions v2 preparat al mateix full que Potències: ' + ss.getUrl());
  } finally {
    lock.releaseLock();
  }
}
