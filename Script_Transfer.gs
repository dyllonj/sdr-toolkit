// Global variables to store document references
let targetDoc;
let sourceSheet;

function onOpen() {
  // Create a custom menu in Google Sheets
  SpreadsheetApp.getUi()
    .createMenu('Auto-Fill Docs')
    .addItem('Setup Documents', 'setupDocuments')
    .addItem('Update Doc from Selected Row (Ctrl+Shift+U)', 'updateDocFromSelection')
    .addItem('Revert to Last Version (Ctrl+Shift+R)', 'revertToLastVersion')
    .addToUi();
    
  // Install triggers for keyboard shortcuts
  installTriggers();
}

function installTriggers() {
  // Remove any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Add trigger for keyboard shortcuts
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

function onEdit(e) {
  // Check if event object exists and contains the necessary information
  if (!e || !e.source || !e.range) return;
  
  // Get the active spreadsheet
  const ss = e.source;
  const range = e.range;
  
  // Check for keyboard shortcuts
  const activeSheet = ss.getActiveSheet();
  const activeRange = ss.getActiveRange();
  
  // Get the keys pressed from the event object
  const keyEvent = e.keyCode;
  
  // Check for Ctrl+Shift+U (Update Doc)
  if (keyEvent === 85 && e.ctrlKey && e.shiftKey) { // 85 is keycode for 'U'
    updateDocFromSelection();
    return;
  }
  
  // Check for Ctrl+Shift+R (Revert Doc)
  if (keyEvent === 82 && e.ctrlKey && e.shiftKey) { // 82 is keycode for 'R'
    revertToLastVersion();
    return;
  }
}

function setupDocuments() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for the target Google Doc URL
  const docResponse = ui.prompt(
    'Setup',
    'Please paste the URL of the target Google Doc:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (docResponse.getSelectedButton() == ui.Button.OK) {
    try {
      // Extract document ID from URL and store in Script Properties
      const docUrl = docResponse.getResponseText();
      const docId = docUrl.match(/[-\w]{25,}/);
      if (!docId) {
        throw new Error('Invalid document URL');
      }
      PropertiesService.getScriptProperties().setProperty('TARGET_DOC_ID', docId[0]);
      ui.alert('Setup Complete', 'Documents have been linked successfully!\n\nKeyboard shortcuts:\nCtrl+Shift+U: Update from selected row\nCtrl+Shift+R: Revert to last version', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', 'Invalid document URL. Please try again.', ui.ButtonSet.OK);
    }
  }
}

function revertToLastVersion() {
  try {
    const docId = PropertiesService.getScriptProperties().getProperty('TARGET_DOC_ID');
    if (!docId) {
      throw new Error('Please run Setup Documents first');
    }
    
    // Get the document's revision history
    const file = DriveApp.getFileById(docId);
    const revisions = Drive.Revisions.list(docId);
    
    if (revisions.items && revisions.items.length > 1) {
      // Get the previous version (second most recent)
      const previousVersion = revisions.items[revisions.items.length - 2];
      
      // Revert to the previous version
      Drive.Revisions.update({}, docId, previousVersion.id);
      
      SpreadsheetApp.getUi().alert('Success', 'Document has been reverted to the previous version!', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      throw new Error('No previous versions found');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function updateDocFromSelection() {
  try {
    // Get the active spreadsheet and selected range
    const sheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = sheet.getActiveRange();
    
    // Ensure only one row is selected
    if (selectedRange.getNumRows() != 1) {
      throw new Error('Please select a single row');
    }
    
    // Get the data from the selected row
    const rowData = selectedRange.getValues()[0];
    
    // Get the target document
    const docId = PropertiesService.getScriptProperties().getProperty('TARGET_DOC_ID');
    if (!docId) {
      throw new Error('Please run Setup Documents first');
    }
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    // Update the document with the row data
    body.replaceText('{{account_name}}', rowData[0] || ''); // Assuming account name is in first column
    body.replaceText('{{church}}', rowData[1] || '');       // Assuming church is in second column
    body.replaceText('{{prospect}}', rowData[2] || '');     // Assuming prospect is in third column
    
    // Save the document
    doc.saveAndClose();
    
    SpreadsheetApp.getUi().alert('Success', 'Document has been updated!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
