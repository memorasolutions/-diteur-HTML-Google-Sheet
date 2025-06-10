// Code.gs - Script principal Google Apps Script

/**
 * Fonction appelée lors de l'ouverture de la feuille
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Éditeur HTML')
    .addItem('Éditer la cellule active', 'editActiveCellDialog')
    .addToUi();
}

/**
 * Fonction appelée lors de la sélection d'une cellule
 */
function onSelectionChange(e) {
  
  const range = e.range;
  if (!range || range.getNumColumns() !== 1 || range.getNumRows() !== 1) return;
  
  const value = range.getValue();
  if (!value || typeof value !== 'string') return;
  
  // Vérifier si la cellule contient du HTML
  if (value.includes('<') && value.includes('>')) {
    openEditor(range);
  }
}

/**
 * Ouvre l'éditeur HTML dans une boîte de dialogue
 */
function openEditor(cell) {
  cell = cell || SpreadsheetApp.getActiveSheet().getActiveCell();
  const cellData = {
    content: cell.getValue() || '',
    row: cell.getRow(),
    col: cell.getColumn(),
  };

  const template = HtmlService.createTemplateFromFile('editor');
  template.cellData = cellData;
  const html = template.evaluate()
    .setTitle('Éditeur HTML');
    
  const ui = SpreadsheetApp.getUi();
  html.setWidth(900).setHeight(700);
  ui.showModalDialog(html, 'Éditeur HTML');
}

/**
 * Ouvre l'éditeur dans une boîte de dialogue
 */
function editActiveCellDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell.getValue();
  
  if (typeof value !== 'string') {
    SpreadsheetApp.getUi().alert('La cellule active ne contient pas de texte.');
    return;
  }
  
  openEditor(cell);
}

/**
 * Récupère le contenu de la cellule active
 */
/**
 * Sauvegarde le contenu dans la cellule
 */
function saveCellContent(content, row, col) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(row, col).setValue(content);
  return { success: true };
}

/**
 * Nettoie le HTML pour éviter les injections XSS
 */
function sanitizeHtml(html) {
  // Liste des balises autorisées
  const allowedTags = ['p', 'br', 'strong', 'b', 'em', 'i', 'u', 'span', 'div', 
                      'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 
                      'a', 'img', 'blockquote', 'code', 'pre', 'table', 'tr', 
                      'td', 'th', 'thead', 'tbody'];
  
  // Supprime les scripts et styles
  html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
  
  // Supprime les attributs dangereux
  html = html.replace(/\son\w+\s*=/gi, ' ');
  html = html.replace(/javascript:/gi, '');
  
  return html;
}
