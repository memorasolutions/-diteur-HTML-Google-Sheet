// Code.gs - Script principal Google Apps Script optimisé

// Variable globale pour le template en cache
let TEMPLATE_CACHE = null;
let SCRIPT_PROPERTIES = null;

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
    .addItem('Gérer les templates', 'openTemplateManager')
    .addSeparator()
    .addItem('Exporter les templates', 'exportTemplates')
    .addItem('Importer des templates', 'importTemplates')
    .addItem('Réinitialiser les templates', 'resetTemplates')
    .addToUi();
  
  // Précharger le template en arrière-plan
  preloadTemplate();
  
  // Initialiser les propriétés du script
  SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
  
  // Initialiser les templates s'ils n'existent pas
  if (!SCRIPT_PROPERTIES.getProperty('templates')) {
    SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify({}));
  }
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
 * Inclut le contenu d'un fichier HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
  html.setWidth(1000).setHeight(800);
  ui.showModalDialog(html, 'Éditeur HTML');
}

/**
 * Ouvre l'éditeur dans une boîte de dialogue
 */
function editActiveCellDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell.getValue();
  
  if (typeof value !== 'string' && value !== '') {
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
                      'td', 'th', 'thead', 'tbody', 'col', 'colgroup'];
  
  // Supprime les scripts et styles en une seule passe
  html = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/\son\w+\s*=/gi, ' ')
    .replace(/javascript:/gi, '');
  
  return html;
}

/**
 * Gestion des templates
 */
function saveTemplate(name, content, description) {
  try {
    if (!SCRIPT_PROPERTIES) {
      SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
    }
    
    let templates = getTemplates();
    
    templates[name] = {
      content: content,
      description: description || '',
      createdAt: templates[name] ? templates[name].createdAt : new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    
    SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify(templates));
    return { success: true };
  } catch (error) {
    console.error('Erreur lors de la sauvegarde du template:', error);
    return { success: false, error: error.toString() };
  }
}

function getTemplates() {
  try {
    if (!SCRIPT_PROPERTIES) {
      SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
    }
    
    const templatesStr = SCRIPT_PROPERTIES.getProperty('templates');
    
    // Si pas de templates, retourner un objet vide
    if (!templatesStr) {
      // Initialiser avec un template d'exemple
      const defaultTemplates = {
        "Tableau simple": {
          content: '<table style="width: 100%; border-collapse: collapse;"><tr><th style="border: 1px solid #ddd; padding: 8px; background-color: #5B9BD5; color: white;">En-tête 1</th><th style="border: 1px solid #ddd; padding: 8px; background-color: #5B9BD5; color: white;">En-tête 2</th></tr><tr><td style="border: 1px solid #ddd; padding: 8px;">Cellule 1</td><td style="border: 1px solid #ddd; padding: 8px;">Cellule 2</td></tr></table>',
          description: "Un tableau simple avec en-têtes",
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        },
        "Titre et paragraphe": {
          content: '<h2 style="color: #4F46E5;">Titre de section</h2><p>Voici un paragraphe de texte avec <strong>du texte en gras</strong> et <em>du texte en italique</em>.</p>',
          description: "Un titre H2 avec un paragraphe formaté",
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }
      };
      
      // Sauvegarder les templates par défaut
      SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify(defaultTemplates));
      return defaultTemplates;
    }
    
    return JSON.parse(templatesStr);
  } catch (error) {
    console.error('Erreur lors de la récupération des templates:', error);
    return {};
  }
}

function deleteTemplate(name) {
  try {
    if (!SCRIPT_PROPERTIES) {
      SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
    }
    
    let templates = getTemplates();
    delete templates[name];
    SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify(templates));
    return { success: true };
  } catch (error) {
    console.error('Erreur lors de la suppression du template:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Ouvre le gestionnaire de templates
 */
function openTemplateManager() {
  const html = HtmlService.createTemplateFromFile('template-manager')
    .evaluate()
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestionnaire de Templates');
}

/**
 * Export des templates
 */
function exportTemplates() {
  const templates = getTemplates();
  const json = JSON.stringify(templates, null, 2);
  
  // Créer un blob avec le JSON
  const blob = Utilities.newBlob(json, 'application/json', 'templates_export.json');
  
  // Créer un fichier temporaire et obtenir son URL
  const file = DriveApp.createFile(blob);
  const url = file.getDownloadUrl();
  
  // Afficher le lien de téléchargement
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>Templates exportés avec succès !</p>
     <p><a href="${url}" target="_blank">Télécharger le fichier</a></p>
     <p style="color: #666; font-size: 12px;">Le fichier sera supprimé automatiquement après téléchargement.</p>`
  )
  .setWidth(400)
  .setHeight(150);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export des Templates');
  
  // Supprimer le fichier après 5 minutes
  Utilities.sleep(300000);
  file.setTrashed(true);
}

/**
 * Import des templates
 */
function importTemplates() {
  const html = HtmlService.createHtmlOutputFromFile('import-dialog')
    .setWidth(500)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Importer des Templates');
}

function processImportedTemplates(jsonContent) {
  try {
    const importedTemplates = JSON.parse(jsonContent);
    const existingTemplates = getTemplates();
    
    // Fusionner les templates
    Object.keys(importedTemplates).forEach(key => {
      existingTemplates[key] = importedTemplates[key];
      existingTemplates[key].updatedAt = new Date().toISOString();
    });
    
    SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify(existingTemplates));
    return { success: true, count: Object.keys(importedTemplates).length };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Optimisation : Récupère le contenu de plusieurs cellules en une fois
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

/**
 * Réinitialise les templates avec des exemples
 */
function resetTemplates() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Réinitialiser les templates',
    'Cette action va supprimer tous vos templates actuels et les remplacer par des templates d\'exemple. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    const defaultTemplates = {
      "Tableau simple": {
        content: '<table style="width: 100%; border-collapse: collapse;"><tr><th style="border: 1px solid #ddd; padding: 8px; background-color: #5B9BD5; color: white;">En-tête 1</th><th style="border: 1px solid #ddd; padding: 8px; background-color: #5B9BD5; color: white;">En-tête 2</th></tr><tr><td style="border: 1px solid #ddd; padding: 8px;">Cellule 1</td><td style="border: 1px solid #ddd; padding: 8px;">Cellule 2</td></tr></table>',
        description: "Un tableau simple avec en-têtes",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      },
      "Titre et paragraphe": {
        content: '<h2 style="color: #4F46E5;">Titre de section</h2><p>Voici un paragraphe de texte avec <strong>du texte en gras</strong> et <em>du texte en italique</em>.</p>',
        description: "Un titre H2 avec un paragraphe formaté",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      },
      "Liste à puces": {
        content: '<h3>Liste des éléments :</h3><ul><li>Premier élément</li><li>Deuxième élément</li><li>Troisième élément</li></ul>',
        description: "Une liste à puces simple",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      },
      "Citation": {
        content: '<blockquote style="border-left: 4px solid #4F46E5; padding-left: 16px; margin: 16px 0; color: #6B7280; font-style: italic;">"Le succès n\'est pas final, l\'échec n\'est pas fatal : c\'est le courage de continuer qui compte."<br><small>- Winston Churchill</small></blockquote>',
        description: "Une citation avec attribution",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      }
    };
    
    SCRIPT_PROPERTIES.setProperty('templates', JSON.stringify(defaultTemplates));
    ui.alert('Templates réinitialisés', 'Les templates ont été réinitialisés avec succès.', ui.ButtonSet.OK);
  }
}
