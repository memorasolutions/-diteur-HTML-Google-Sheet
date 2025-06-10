// Code.gs - Script principal Google Apps Script optimisé

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
 * Ouvre l'éditeur HTML dans une boîte de dialogue - OPTIMISÉ
 */
function openEditor(cell) {
  cell = cell || SpreadsheetApp.getActiveSheet().getActiveCell();
  
  // Utiliser getDisplayValue() pour une lecture plus rapide
  const cellData = {
    content: cell.getDisplayValue() || '',
    row: cell.getRow(),
    col: cell.getColumn(),
  };

  // Créer le template
  const template = HtmlService.createTemplateFromFile('editor');
  template.cellData = cellData;
  
  // Évaluer avec les optimisations
  const html = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
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
 * Sauvegarde le contenu dans la cellule - OPTIMISÉ
 */
function saveCellContent(content, row, col) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange(row, col);
    
    // Désactiver temporairement les recalculs pour améliorer les performances
    const recalcEnabled = sheet.isRecalculationEnabled();
    if (recalcEnabled) {
      sheet.setRecalculationEnabled(false);
    }
    
    // Nettoyer le contenu avant sauvegarde
    const cleanContent = sanitizeHtml(content);
    
    // Sauvegarder la valeur
    range.setValue(cleanContent);
    
    // Flush pour forcer l'écriture immédiate
    SpreadsheetApp.flush();
    
    // Réactiver les recalculs
    if (recalcEnabled) {
      sheet.setRecalculationEnabled(true);
    }
    
    return { success: true };
  } catch (error) {
    // En cas d'erreur, s'assurer que les recalculs sont réactivés
    try {
      SpreadsheetApp.getActiveSheet().setRecalculationEnabled(true);
    } catch (e) {}
    
    throw error;
  }
}

/**
 * Nettoie le HTML pour éviter les injections XSS - OPTIMISÉ
 */
function sanitizeHtml(html) {
  // Retour rapide si pas de HTML
  if (!html || typeof html !== 'string') return '';
  if (!html.includes('<') && !html.includes('>')) return html;
  
  // Liste des balises autorisées
  const allowedTags = ['p', 'br', 'strong', 'b', 'em', 'i', 'u', 'span', 'div', 
                      'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 
                      'a', 'img', 'blockquote', 'code', 'pre', 'table', 'tr', 
                      'td', 'th', 'thead', 'tbody'];
  
  // Supprime les scripts et styles en une seule passe
  html = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/\son\w+\s*=/gi, ' ')
    .replace(/javascript:/gi, '');
  
  return html;
}

/**
 * Insère un bouton d'accès direct à l'éditeur HTML.
 * Le bouton est ajouté en haut à gauche de la feuille et
 * exécute la fonction editActiveCellDialog lors d'un clic.
 */
function insertEditorButton() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var imageUrl = 'https://www.gstatic.com/images/icons/material/system/1x/edit_black_24dp.png';
  var blob = UrlFetchApp.fetch(imageUrl).getBlob();
  var button = sheet.insertImage(blob, 1, 1);
  if (button.assignScript) {
    button.assignScript('editActiveCellDialog');
  }
  button.setAltTextDescription('Éditer la cellule active');
}

/**
 * Optimisation : Récupère le contenu de plusieurs cellules en une fois
 * Utile si vous avez besoin de traiter plusieurs cellules
 */
function getBatchCellContents(ranges) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const results = [];
  
  ranges.forEach(range => {
    const cell = sheet.getRange(range.row, range.col);
    results.push({
      content: cell.getDisplayValue() || '',
      row: range.row,
      col: range.col
    });
  });
  
  return results;
}

/**
 * Optimisation : Sauvegarde plusieurs cellules en une fois
 */
function saveBatchCellContents(updates) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const recalcEnabled = sheet.isRecalculationEnabled();
  
  try {
    // Désactiver les recalculs
    if (recalcEnabled) {
      sheet.setRecalculationEnabled(false);
    }
    
    // Appliquer toutes les mises à jour
    updates.forEach(update => {
      const range = sheet.getRange(update.row, update.col);
      const cleanContent = sanitizeHtml(update.content);
      range.setValue(cleanContent);
    });
    
    // Forcer l'écriture
    SpreadsheetApp.flush();
    
    // Réactiver les recalculs
    if (recalcEnabled) {
      sheet.setRecalculationEnabled(true);
    }
    
    return { success: true, count: updates.length };
  } catch (error) {
    // S'assurer que les recalculs sont réactivés en cas d'erreur
    try {
      if (recalcEnabled) {
        sheet.setRecalculationEnabled(true);
      }
    } catch (e) {}
    
    throw error;
  }
}
