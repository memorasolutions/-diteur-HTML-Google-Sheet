// Code.gs - Script principal Google Apps Script optimisé

/**
 * Note: Les méthodes isRecalculationEnabled() et setRecalculationEnabled()
 * n'existent pas dans l'API Google Apps Script actuelle.
 * Les optimisations utilisent d'autres techniques comme SpreadsheetApp.flush()
 * pour améliorer les performances.
 */

// Variable globale pour le template en cache
let TEMPLATE_CACHE = null;

/**
 * Cache simple pour éviter de recharger le même contenu
 */
const CACHE = {
  content: new Map(),
  maxSize: 50,
  
  get: function(row, col) {
    const key = `${row},${col}`;
    return this.content.get(key);
  },
  
  set: function(row, col, value) {
    const key = `${row},${col}`;
    
    // Limiter la taille du cache
    if (this.content.size >= this.maxSize) {
      const firstKey = this.content.keys().next().value;
      this.content.delete(firstKey);
    }
    
    this.content.set(key, value);
  },
  
  clear: function() {
    this.content.clear();
  }
};

/**
 * Pré-charge le template pour améliorer les performances
 */
function preloadTemplate() {
  try {
    TEMPLATE_CACHE = HtmlService.createTemplateFromFile('editor');
  } catch (e) {
    console.error('Erreur lors du préchargement du template:', e);
  }
}

/**
 * Fonction appelée lors de l'ouverture de la feuille
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Éditeur HTML')
    .addItem('Éditer la cellule active', 'editActiveCellDialog')
    .addToUi();
  
  // Précharger le template en arrière-plan
  preloadTemplate();
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
  
  // Utiliser getValue() avec une vérification rapide
  const value = cell.getValue();
  const cleanValue = sanitizeHtml(value);
  const cellData = {
    content: cleanValue || '',
    row: cell.getRow(),
    col: cell.getColumn(),
  };

  // Utiliser le template en cache si disponible, sinon le créer
  let template;
  try {
    template = TEMPLATE_CACHE || HtmlService.createTemplateFromFile('editor');
  } catch (e) {
    template = HtmlService.createTemplateFromFile('editor');
  }
  
  template.cellData = cellData;
  
  // Évaluer avec les optimisations
  const html = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Éditeur HTML');
    
  const ui = SpreadsheetApp.getUi();
  html.setWidth(900).setHeight(750);
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
    
    // Nettoyer le contenu avant sauvegarde
    const cleanContent = sanitizeHtml(content);
    
    // Utiliser getRange et setValue en une seule opération
    sheet.getRange(row, col).setValue(cleanContent);
    
    // Invalider le cache pour cette cellule
    if (typeof CACHE !== 'undefined' && CACHE && CACHE.content) {
      const key = `${row},${col}`;
      CACHE.content.delete(key);
    }
    
    // Flush pour forcer l'écriture immédiate
    SpreadsheetApp.flush();
    
    return { success: true };
  } catch (error) {
    console.error('Erreur lors de la sauvegarde:', error);
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
      content: cell.getValue() || '',
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
  
  try {
    // Pour optimiser, grouper les cellules contiguës
    const values = [];
    const ranges = [];
    
    updates.forEach(update => {
      const cleanContent = sanitizeHtml(update.content);
      values.push([cleanContent]);
      ranges.push(sheet.getRange(update.row, update.col));
    });
    
    // Utiliser setValues pour chaque range
    ranges.forEach((range, index) => {
      range.setValue(values[index][0]);
    });
    
    // Forcer l'écriture
    SpreadsheetApp.flush();
    
    return { success: true, count: updates.length };
  } catch (error) {
    console.error('Erreur lors de la sauvegarde batch:', error);
    throw error;
  }
}

/**
 * Version alternative optimisée pour les cellules contiguës
 */
function saveCellContentsFast(startRow, startCol, contents) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Nettoyer tous les contenus
    const cleanContents = contents.map(content => [sanitizeHtml(content)]);
    
    // Obtenir la plage de cellules
    const range = sheet.getRange(startRow, startCol, cleanContents.length, 1);
    
    // Écrire toutes les valeurs en une fois
    range.setValues(cleanContents);
    
    // Forcer l'écriture
    SpreadsheetApp.flush();
    
    return { success: true, count: cleanContents.length };
  } catch (error) {
    console.error('Erreur lors de la sauvegarde rapide:', error);
    throw error;
  }
}

/**
 * Fonction utilitaire pour mesurer les performances
 */
function measurePerformance(functionName, ...args) {
  const startTime = new Date().getTime();
  
  try {
    const result = this[functionName].apply(this, args);
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    console.log(`${functionName} a pris ${duration}ms`);
    return result;
  } catch (error) {
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    console.error(`${functionName} a échoué après ${duration}ms:`, error);
    throw error;
  }
}
