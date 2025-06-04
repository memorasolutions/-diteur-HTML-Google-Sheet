// Code.gs - Script principal Google Apps Script

/**
 * Fonction appelée lors de l'ouverture de la feuille
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Éditeur HTML')
    .addItem('Éditer dans la barre latérale', 'editActiveCell')
    .addItem('Éditer dans une fenêtre (plus large)', 'editActiveCellDialog')
    .addToUi();
}

/**
 * Édite la cellule active si elle contient du HTML
 */
function editActiveCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell.getValue();
  
  if (!value || typeof value !== 'string') {
    SpreadsheetApp.getUi().alert('La cellule active est vide ou ne contient pas de texte.');
    return;
  }
  
  if (value.includes('<') && value.includes('>')) {
    openEditor();
  } else {
    // Proposer d'ouvrir même sans HTML
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Pas de HTML détecté', 
      'La cellule ne semble pas contenir de HTML. Voulez-vous l\'éditer quand même ?',
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      openEditor();
    }
  }
}

/**
 * Fonction appelée lors de la sélection d'une cellule
 */
function onSelectionChange(e) {
  const isActive = PropertiesService.getUserProperties().getProperty('editorActive');
  if (isActive !== 'true') return;
  
  const range = e.range;
  if (!range || range.getNumColumns() !== 1 || range.getNumRows() !== 1) return;
  
  const value = range.getValue();
  if (!value || typeof value !== 'string') return;
  
  // Vérifier si la cellule contient du HTML
  if (value.includes('<') && value.includes('>')) {
    openEditor();
  }
}

/**
 * Ouvre l'éditeur HTML dans une boîte de dialogue ou sidebar
 */
function openEditor(useDialog = false) {
  const html = HtmlService.createHtmlOutputFromFile('editor')
    .setTitle('Éditeur HTML');
    
  const ui = SpreadsheetApp.getUi();
  
  if (useDialog) {
    // Boîte de dialogue modale (peut être plus large)
    html.setWidth(900).setHeight(700);
    ui.showModalDialog(html, 'Éditeur HTML');
  } else {
    // Sidebar (max 600px de large)
    html.setWidth(600);
    ui.showSidebar(html);
  }
}

/**
 * Ouvre l'éditeur dans une boîte de dialogue
 */
function editActiveCellDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell.getValue();
  
  if (!value || typeof value !== 'string') {
    SpreadsheetApp.getUi().alert('La cellule active est vide ou ne contient pas de texte.');
    return;
  }
  
  openEditor(true); // Ouvre en mode dialogue
}

/**
 * Récupère le contenu de la cellule active
 */
function getCellContent() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  return {
    content: cell.getValue() || '',
    row: cell.getRow(),
    col: cell.getColumn()
  };
}

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
