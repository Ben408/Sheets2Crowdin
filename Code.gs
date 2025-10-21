/**
 * Crowdin Enterprise Integration for Google Sheets
 * Syncs translations between row-based Google Sheets and Crowdin Enterprise
 * 
 * Organization: strava.crowdin.com
 */

// Configuration keys
const CONFIG_KEYS = {
  CROWDIN_TOKEN: 'CROWDIN_TOKEN',
  CROWDIN_PROJECT_ID: 'CROWDIN_PROJECT_ID',
  CROWDIN_ORGANIZATION: 'strava'
};

/**
 * Get the correct Crowdin Enterprise API base URL
 * IMPORTANT: Use Enterprise org subdomain for API, not dashboard hostname
 */
function getCrowdinBaseUrl() {
  const org = CONFIG_KEYS.CROWDIN_ORGANIZATION;
  return 'https://' + org + '.api.crowdin.com/api/v2';
}

/**
 * Get authentication headers
 */
function getAuthHeaders() {
  const token = PropertiesService.getUserProperties().getProperty(CONFIG_KEYS.CROWDIN_TOKEN);
  if (!token) throw new Error('No Crowdin token saved. Open Config and save it.');
  return {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json'
  };
}

/**
 * Get the correct branch ID for the project (handles projects with no branches)
 */
function getBranchId(config) {
  try {
    const projectId = config.projectId;
    const branchesUrl = getCrowdinBaseUrl() + '/projects/' + projectId + '/branches';
    
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    const response = UrlFetchApp.fetch(branchesUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      if (data.data && data.data.length > 0) {
        // Use the first branch ID
        var branchId = data.data[0].data.id;
        Logger.log('Found branch ID: ' + branchId);
        return branchId;
      }
    }
    
    Logger.log('No branches found - will create strings without branchId');
    return 0; // No branches found
    
  } catch (error) {
    Logger.log('Error getting branch ID: ' + error.message);
    return 0; // Fallback to no branch
  }
}

/**
 * Get or create a file ID for the project (required for string creation)
 */
function getFileId(config) {
  try {
    const projectId = config.projectId;
    const filesUrl = getCrowdinBaseUrl() + '/projects/' + projectId + '/files';
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    const response = UrlFetchApp.fetch(filesUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      if (data.data && data.data.length > 0) {
        // Use the first file ID
        var fileId = data.data[0].data.id;
        Logger.log('Found existing file ID: ' + fileId);
        return fileId;
      }
    }
    
    // No files found - create a new one
    Logger.log('No files found - creating a new file');
    return createFile(config);
    
  } catch (error) {
    Logger.log('Error getting file ID: ' + error.message);
    return createFile(config);
  }
}

/**
 * Create a new file for the project
 */
function createFile(config) {
  try {
    const projectId = config.projectId;
    const createFileUrl = getCrowdinBaseUrl() + '/projects/' + projectId + '/files';
    
    // Create a simple JSON file with a placeholder string
    var fileContent = JSON.stringify({
      "placeholder": "This is a placeholder file for Google Sheets strings"
    }, null, 2);
    
    // For now, let's try to create a file using the storage API
    // First, we need to upload the file content to storage
    var storageUrl = getCrowdinBaseUrl() + '/storages';
    
    var storageOptions = {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + PropertiesService.getUserProperties().getProperty(CONFIG_KEYS.CROWDIN_TOKEN),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        fileName: 'google_sheets_strings.json',
        content: fileContent
      }),
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    var storageResponse = UrlFetchApp.fetch(storageUrl, storageOptions);
    var storageCode = storageResponse.getResponseCode();
    var storageText = storageResponse.getContentText();
    
    if (storageCode >= 200 && storageCode < 300) {
      var storageData = JSON.parse(storageText);
      var storageId = storageData.data.id;
      
      Logger.log('Created storage ID: ' + storageId);
      
      // Now create the file using the storage ID
      var fileOptions = {
        method: 'POST',
        headers: getAuthHeaders(),
        payload: JSON.stringify({
          storageId: storageId,
          name: 'google_sheets_strings.json',
          title: 'Google Sheets Strings'
        }),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      var fileResponse = UrlFetchApp.fetch(createFileUrl, fileOptions);
      var fileCode = fileResponse.getResponseCode();
      var fileText = fileResponse.getContentText();
      
      if (fileCode >= 200 && fileCode < 300) {
        var fileData = JSON.parse(fileText);
        var fileId = fileData.data.id;
        Logger.log('Created file ID: ' + fileId);
        return fileId;
      }
    }
    
    Logger.log('Could not create file - using fallback approach');
    return 'fallback'; // Fallback value
    
  } catch (error) {
    Logger.log('Error creating file: ' + error.message);
    return 'fallback'; // Fallback value
  }
}

// Language mapping (Sheet Label -> Crowdin locale)
const LANGUAGE_MAP = {
  'English (US)': 'en',
  'French': 'fr-FR',
  'German': 'de-DE',
  'Japanese': 'ja-JP',
  'Spain Spanish': 'es-ES',
  'Dutch': 'nl-NL',
  'Italian': 'it-IT',
  'Portuguese': 'pt-PT',
  'Russian': 'ru-RU',
  'LATAM Spanish': 'es-419',
  'Traditional Chinese': 'zh-TW',
  'Simplified Chinese': 'zh-CN',
  'Indonesian': 'id-ID'
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌍 Crowdin Sync')
    .addItem('📤 Push Selected Cells to Crowdin', 'pushSelectedCells')
    .addItem('📤 Push Current Sheet Only', 'pushCurrentSheet')
    .addSeparator()
    .addItem('📥 Pull Current Sheet Only (Complete)', 'pullCurrentSheetComplete')
    .addItem('📥 Pull Selected Strings Only', 'pullSelectedStrings')
    .addSeparator()
    .addItem('⚙️ Configure Settings', 'showConfigDialog')
    .addItem('🔍 Test Connection', 'testConnection')
    .addItem('🧪 Test API Endpoints', 'testApiEndpoints')
    .addItem('📊 Test Spreadsheet Access', 'testSpreadsheetAccess')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * Shows configuration dialog
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Config')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Crowdin Enterprise Configuration');
}

/**
 * Shows about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Crowdin Enterprise Sync',
    'Version 1.0\n\n' +
    'This tool syncs translations between your row-based Google Sheet and Crowdin Enterprise.\n\n' +
    'Setup:\n' +
    '1. Go to Crowdin Sync → Configure Settings\n' +
    '2. Enter your Personal Access Token and Project ID\n' +
    '3. Use Push/Pull to sync content\n\n' +
    'Organization: strava.crowdin.com',
    ui.ButtonSet.OK
  );
}

/**
 * Saves configuration from dialog
 */
function saveConfig(token, projectId) {
  try {
    const props = PropertiesService.getUserProperties();
    
    // Validate inputs
    if (!token || token.trim() === '') {
      throw new Error('Personal Access Token is required');
    }
    
    if (!projectId || projectId.trim() === '') {
      throw new Error('Project ID is required');
    }
    
    // Ensure project ID is numeric
    if (isNaN(projectId)) {
      throw new Error('Project ID must be a number');
    }
    
    // Save properties
    props.setProperty(CONFIG_KEYS.CROWDIN_TOKEN, token.trim());
    props.setProperty(CONFIG_KEYS.CROWDIN_PROJECT_ID, projectId.trim());
    
    // Verify save worked
    const savedToken = props.getProperty(CONFIG_KEYS.CROWDIN_TOKEN);
    const savedProjectId = props.getProperty(CONFIG_KEYS.CROWDIN_PROJECT_ID);
    
    if (!savedToken || !savedProjectId) {
      throw new Error('Failed to save configuration - check permissions');
    }
    
    return 'Configuration saved successfully!';
  } catch (error) {
    Logger.log('Save config error: ' + error.message);
    throw new Error('Failed to save configuration: ' + error.message);
  }
}

/**
 * Gets current configuration
 */
function getConfig() {
  const props = PropertiesService.getUserProperties();
  return {
    token: props.getProperty(CONFIG_KEYS.CROWDIN_TOKEN) || '',
    projectId: props.getProperty(CONFIG_KEYS.CROWDIN_PROJECT_ID) || ''
  };
}

/**
 * Validates configuration
 */
function validateConfig() {
  const config = getConfig();
  if (!config.token || !config.projectId) {
    throw new Error('Please configure Crowdin settings first (Crowdin Sync → Configure Settings)');
  }
  return config;
}

/**
 * Push selected cells to Crowdin (NEW FEATURE)
 */
function pushSelectedCells() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get the active selection
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const selection = sheet.getSelection();
    
    if (!selection) {
      ui.alert('No Selection', 'Please select some cells first, then run this function.', ui.ButtonSet.OK);
      return;
    }
    
    const ranges = selection.getActiveRangeList().getRanges();
    if (ranges.length === 0) {
      ui.alert('No Selection', 'Please select some cells first, then run this function.', ui.ButtonSet.OK);
      return;
    }
    
    // Get all selected cells
    const selectedCells = [];
    ranges.forEach(range => {
      const values = range.getValues();
      const rowStart = range.getRow();
      const colStart = range.getColumn();
      
      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const cellValue = values[row][col];
          if (cellValue && cellValue.toString().trim() !== '') {
            const actualRow = rowStart + row;
            const actualCol = colStart + col;
            const columnLetter = columnToLetter(actualCol);
            const maxvalue = findCharMaxValue(actualCol)
          
            selectedCells.push({
              text: cellValue.toString().trim(),
              identifier: sheet.getName() + '_R' + actualRow + 'C' + columnLetter,
              context: 'Sheet: ' + sheet.getName() + ', Cell: ' + columnLetter + actualRow,
              max_char: maxvalue
            });
          }
        }
      }
    });
    
    if (selectedCells.length === 0) {
      ui.alert('No Content', 'Selected cells are empty. Please select cells with text content.', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert(
      'Ready to Push Selected Cells', 
      'Found ' + selectedCells.length + ' cells with content:\n\n' +
      selectedCells.slice(0, 3).map(cell => '"' + cell.text.substring(0, 50) + (cell.text.length > 50 ? '...' : '') + '"').join('\n') +
      (selectedCells.length > 3 ? '\n... and ' + (selectedCells.length - 3) + ' more' : '') +
      '\n\nPush these to Crowdin?',
      ui.ButtonSet.YES_NO
    );
    
    // Upload selected cells
    let processed = 0;
    let errors = 0;
    
    for (const cell of selectedCells) {
      try {
        const result = uploadSingleString(cell, config);
        if (result.success) {
          processed++;
          Logger.log('✅ Uploaded: ' + cell.identifier);
        } else {
          errors++;
          Logger.log('❌ Failed: ' + cell.identifier + ' - ' + result.error);
        }
        
        Utilities.sleep(100);
        
      } catch (error) {
        errors++;
        Logger.log('❌ Error with ' + cell.identifier + ': ' + error.message);
      }
    }
    
    if (processed > 0) {
      ui.alert(
        'Push Complete! ✅',
        'Successfully pushed ' + processed + ' selected cells to Crowdin.\n\n' +
        'Errors: ' + errors + '\n' +
        'Strings are now available for translation in Crowdin.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Push Failed',
        'Could not push any selected cells. Errors: ' + errors,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', 'Failed to push selected cells:\n' + error.message, ui.ButtonSet.OK);
    Logger.log('Push selected cells error: ' + error.stack);
  }
}

/**
 * Main function to push source strings to Crowdin
 */
function pushToCrowdin() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get spreadsheet with better error handling
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (error) {
      throw new Error('Cannot access spreadsheet. Please ensure you have edit permissions and try running this from the sheet menu.');
    }
    
    const sheets = ss.getSheets();
    
    ui.alert('Starting push to Crowdin...', ui.ButtonSet.OK);
    
    let totalStrings = 0;
    let processedSheets = 0;
    
    // Process each sheet
    for (const sheet of sheets) {
      const result = processSheetForPush(sheet, config);
      if (result) {
        totalStrings += result.stringCount;
        processedSheets++;
      }
    }
    
    if (processedSheets === 0) {
      ui.alert(
        'No data to push',
        'No sheets found with "English (US)" row.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Push Complete! ✅',
        `Successfully pushed ${totalStrings} strings from ${processedSheets} sheet(s) to Crowdin.\n\n` +
        `Strings are now available for translation in Crowdin.`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to push to Crowdin:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Push error: ' + error.stack);
  }
}

/**
 * Processes a single sheet for pushing to Crowdin (goes to Column Z)
 */
function processSheetForPush(sheet, config) {
  const data = sheet.getDataRange().getValues();
  
  // Find the "English (US)" row
  var englishRowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes('English (US)')) {
      englishRowIndex = i;
      break;
    }
  }
  
  if (englishRowIndex === -1) {
    return null; // No English row found in this sheet
  }
  
  const englishRow = data[englishRowIndex];
  
  // Create strings in Crowdin (now goes to Column Z)
  var strings = [];
  var maxColumns = Math.min(englishRow.length, 26); // Columns A-Z (index 0-25)
  
  for (var col = 3; col < maxColumns; col++) { // Start from column D (index 3) to Z
    var sourceText = englishRow[col];
    if (sourceText && sourceText.toString().trim() !== '') {
      // Create a unique identifier for this string
      var columnLetter = columnToLetter(col + 1);
      var stringId = sheet.getName() + '_' + columnLetter;
      
      strings.push({
        identifier: stringId,
        text: sourceText.toString(),
        context: 'Sheet: ' + sheet.getName() + ', Column: ' + columnLetter
      });
    }
  }
  
  if (strings.length > 0) {
    uploadStringsToCrowdin(strings, config);
  }
  
  return { stringCount: strings.length };
}

/**
 * Uploads strings to Crowdin using String-based projects API
 */
function uploadStringsToCrowdin(strings, config) {
  const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
  const projectId = config.projectId;
  
  Logger.log(`Starting upload of ${strings.length} strings to project ${projectId}`);
  
  // Try the String-based API first
  try {
    uploadStringsDirectly(strings, config);
  } catch (error) {
    Logger.log(`String-based upload failed: ${error.message}`);
    Logger.log(`Trying alternative approach...`);
    
    // Fallback: Try creating a JSON file and uploading it
    try {
      uploadStringsAsFile(strings, config);
    } catch (fileError) {
      Logger.log(`File-based upload also failed: ${fileError.message}`);
      throw new Error(`Both upload methods failed. String API: ${error.message}, File API: ${fileError.message}`);
    }
  }
}

/**
 * Upload strings directly using String-based API (optimized for speed)
 */
function uploadStringsDirectly(strings, config) {
  const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
  const projectId = config.projectId;
  
  Logger.log(`Starting fast upload of ${strings.length} strings`);
  
  // Get the main branch ID
  let branchId = 0;
  try {
    const branchesUrl = `${baseUrl}/projects/${projectId}/branches`;
    const branchesResponse = makeApiRequest(branchesUrl, 'GET', null, config.token);
    if (branchesResponse.data && branchesResponse.data.length > 0) {
      // Use the first branch (usually "main")
      branchId = branchesResponse.data[0].data.id;
      Logger.log(`Using branch ID: ${branchId}`);
    }
  } catch (error) {
    Logger.log(`Could not fetch branches: ${error.message}`);
  }
  
  // Get all existing strings in one call
  let existingStrings = [];
  try {
    const listUrl = `${baseUrl}/projects/${projectId}/strings?branchId=${branchId}&limit=500`;
    const response = makeApiRequest(listUrl, 'GET', null, config.token);
    existingStrings = response.data || [];
    Logger.log(`Found ${existingStrings.length} existing strings`);
  } catch (error) {
    Logger.log(`Could not fetch existing strings: ${error.message}`);
  }
  
  // Create a map of existing strings by identifier for fast lookup
  const existingMap = {};
  existingStrings.forEach(item => {
    if (item.data && item.data.identifier) {
      existingMap[item.data.identifier] = item.data;
    }
  });
  
  // Process strings quickly
  let processed = 0;
  let errors = 0;
  
  for (const string of strings) {
    try {
      if (existingMap[string.identifier]) {
        // Update existing string
        const updateUrl = `${baseUrl}/projects/${projectId}/strings/${existingMap[string.identifier].id}`;
        const updatePayload = {
          text: string.text,
          context: string.context
        };
        
        makeApiRequest(updateUrl, 'PATCH', updatePayload, config.token);
        Logger.log(`Updated: ${string.identifier}`);
      } else {
        // Create new string
        const createUrl = `${baseUrl}/projects/${projectId}/strings`;
        const createPayload = {
          text: string.text,
          identifier: string.identifier,
          context: string.context,
          branchId: branchId  // Use the actual branch ID
        };
        
        makeApiRequest(createUrl, 'POST', createPayload, config.token);
        Logger.log(`Created: ${string.identifier}`);
      }
      
      processed++;
      
      // Only add delay every 10 strings to speed up
      if (processed % 10 === 0) {
        Utilities.sleep(50); // Reduced delay
      }
      
    } catch (error) {
      errors++;
      Logger.log(`Error with ${string.identifier}: ${error.message}`);
    }
  }
  
  Logger.log(`Fast upload complete. Processed: ${processed}, Errors: ${errors}`);
}

/**
 * Upload strings as a JSON file (fallback method)
 */
function uploadStringsAsFile(strings, config) {
  const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
  const projectId = config.projectId;
  
  Logger.log(`Trying file-based upload for ${strings.length} strings`);
  
  // Create a JSON structure for the strings
  const jsonData = {};
  strings.forEach(string => {
    jsonData[string.identifier] = string.text;
  });
  
  // Create a simple JSON file content
  const fileContent = JSON.stringify(jsonData, null, 2);
  
  // Upload as a file to Crowdin
  const uploadUrl = `${baseUrl}/projects/${projectId}/files`;
  
  // Create a simple JSON file upload
  const formData = {
    'storageId': 'temp', // This might need adjustment
    'name': 'strings.json',
    'title': 'Google Sheets Strings'
  };
  
  // For now, let's try a simpler approach - just log what we would upload
  Logger.log(`Would upload file with content: ${fileContent}`);
  
  // This is a placeholder - we'd need to implement proper file upload
  throw new Error('File-based upload not yet implemented - checking logs for string-based errors first');
}

/**
 * Finds a string in Crowdin by identifier
 */
function findStringByIdentifier(identifier, config) {
  const baseUrl = getCrowdinBaseUrl();
  const projectId = config.projectId;
  const branchId = getBranchId(config);
  
  var url = baseUrl + '/projects/' + projectId + '/strings?filter=' + encodeURIComponent(identifier);
  if (branchId > 0) {
    url += '&branchId=' + branchId;
  }
  
  try {
    Logger.log('Searching for string with identifier: ' + identifier);
    Logger.log('Using URL: ' + url);
    
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('String search response code: ' + responseCode);
    Logger.log('String search response: ' + (responseText ? responseText.slice(0, 500) : ''));
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      if (data.data && data.data.length > 0) {
        Logger.log('Found string: ' + data.data[0].data.identifier + ' with ID: ' + data.data[0].data.id);
        return data.data[0].data;
      } else {
        Logger.log('No strings found for identifier: ' + identifier);
      }
    } else {
      Logger.log('String search failed: ' + responseCode + ' - ' + responseText);
    }
  } catch (error) {
    Logger.log('Error finding string: ' + error.message);
  }
  
  return null;
}

/**
 * Push only the current active sheet to Crowdin (fixed fast version)
 */
function pushCurrentSheet() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get current active sheet
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (error) {
      throw new Error('Cannot access spreadsheet. Please ensure you have edit permissions and try running this from the sheet menu.');
    }
    
    const activeSheet = ss.getActiveSheet();
    
    ui.alert('Starting fast push to Crowdin...', `Pushing sheet: ${activeSheet.getName()}`, ui.ButtonSet.OK);
    
    // Get strings from sheet using fast extraction
    const strings = extractStringsFromSheet(activeSheet);
    
    if (strings.length === 0) {
      ui.alert(
        'No data to push',
        `No "English (US)" row found in sheet "${activeSheet.getName()}".`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    Logger.log(`Found ${strings.length} strings to upload`);
    // Upload strings one by one with minimal delays
    let processed = 0;
    let errors = 0;
    
    for (const string of strings) {
      try {
          const result = uploadSingleString(string, config);
          if (result.success) {
            processed++;
            Logger.log(`✅ Uploaded: ${string.identifier}`);
          } else {
            errors++;
            Logger.log(`❌ Failed: ${string.identifier} - ${result.error}`);
          }

          // Very small delay between strings
          Utilities.sleep(100);
        
      } catch (error) {
        errors++;
        Logger.log(`❌ Error with ${string.identifier}: ${error.message}`);
      }
    }
    
    if (processed > 0) {
      ui.alert(
        'Push Complete! ✅',
        `Successfully pushed ${processed} strings from "${activeSheet.getName()}" to Crowdin.\n\n` +
        `Errors: ${errors}\n` +
        `Strings are now available for translation in Crowdin.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Push Failed',
        `Could not push any strings. Errors: ${errors}`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to push to Crowdin:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Push error: ' + error.stack);
  }
}


/**
 * Extract strings from a sheet (fixed fast version - goes to Column Z)
 */
function extractStringsFromSheet(sheet) {
  try {
    Logger.log('Fast extraction from: ' + sheet.getName());
    
    // Use a more targeted approach - get specific rows instead of large ranges
    const maxRows = 20; // Only check first 20 rows
    const maxCols = 26; // Check all columns A-Z (index 0-25)
    
    // Get data in a more efficient way
    const data = [];
    for (var row = 1; row <= maxRows; row++) {
      var rowData = sheet.getRange(row, 1, 1, maxCols).getValues()[0];
      data.push(rowData);
      console.log(rowData)
      // Check if we found the English row
      if (rowData[0] && rowData[0].toString().includes('English (US)')) {
        Logger.log('Found English row at: ' + row);
        
        // Extract strings from this row
        var strings = [];
        for (var col = 3; col < maxCols; col++) { // Columns D-Z (index 3-25)
          console.log(rowData[col].toString())
          console.log("entrada: " + rowData[col].toString() + " es o no es numero?: " + isNaN(rowData[col].toString()))
          if (rowData[col] && rowData[col].toString().trim() !== '' && isNaN(rowData[col].toString())) {
            var columnLetter = columnToLetter(col + 1);
            const maxvalue = findCharMaxValue(col + 1)
            
            strings.push({
              text: rowData[col].toString().trim(),
              identifier: sheet.getName() + '_R' + row + 'C' + columnLetter,
              context: 'Sheet: ' + sheet.getName() + ', Cell: ' + columnLetter + row,
              max_char : maxvalue
            });
            
            Logger.log('Added string ' + columnLetter + ': ' + rowData[col].toString().substring(0, 50) + '...');
          }
        }
        
        Logger.log('Extracted ' + strings.length + ' strings');
        return strings;
      }
    }
    
    Logger.log('No English (US) row found in first ' + maxRows + ' rows');
    return [];
    
  } catch (error) {
    Logger.log('Fast extraction failed: ' + error.message);
    return [];
  }
}

/**
 * Upload a small batch of strings
 */
function uploadStringBatch(strings, config) {
  const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
  const projectId = config.projectId;
  
  // Get branch ID
  let branchId = 228360; // Use the known branch ID
  try {
    const branchesUrl = `${baseUrl}/projects/${projectId}/branches`;
    const branchesResponse = makeApiRequest(branchesUrl, 'GET', null, config.token);
    if (branchesResponse.data && branchesResponse.data.length > 0) {
      branchId = branchesResponse.data[0].data.id;
    }
  } catch (error) {
    Logger.log(`Using default branch ID: ${branchId}`);
  }
  
  // Upload each string in the batch
  for (const string of strings) {
    try {
      const createUrl = `${baseUrl}/projects/${projectId}/strings`;
      const createPayload = {
        text: string.text,
        identifier: string.identifier,
        context: string.context,
        branchId: branchId
      };
      
      makeApiRequest(createUrl, 'POST', createPayload, config.token);
      Logger.log(`Created: ${string.identifier}`);
      
      // Small delay between individual strings
      Utilities.sleep(100);
      
    } catch (error) {
      Logger.log(`Error creating ${string.identifier}: ${error.message}`);
      throw error;
    }
  }
}

/**
 * Push just ONE string to test the connection (guaranteed to work)
 */
function pushSingleString() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get current active sheet
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (error) {
      throw new Error('Cannot access spreadsheet. Please ensure you have edit permissions and try running this from the sheet menu.');
    }
    
    const activeSheet = ss.getActiveSheet();
    
    // Get the first string from the sheet
    const strings = extractStringsFromSheet(activeSheet);
    
    if (strings.length === 0) {
      ui.alert(
        'No data to push',
        `No "English (US)" row found in sheet "${activeSheet.getName()}".`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Take only the first string
    const testString = strings[0];
    
    ui.alert('Testing single string push...', `Pushing: "${testString.text}"`, ui.ButtonSet.OK);
    
    // Upload just this one string
    const result = uploadSingleString(testString, config);
    
    if (result.success) {
      ui.alert(
        'Single String Push Complete! ✅',
        `Successfully pushed: "${testString.text}"\n` +
        `String ID: ${result.stringId}\n` +
        `Identifier: ${testString.identifier}\n\n` +
        `Check Crowdin to see if it appears!`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Single String Push Failed',
        `Error: ${result.error}`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to push single string:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Single string push error: ' + error.stack);
  }
}

/**
 * Upload a single string (ultra-simple version)
 */
function uploadSingleString(string, config) {
  try {
    Logger.log(`Starting ultra-simple upload: ${string.identifier}`);
    
    const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
    const projectId = config.projectId;
    const branchId = getBranchId(config);

    const url = `${baseUrl}/projects/${projectId}/strings`;

    let payload = null
    if (string.max_char == 0){
    payload = {
      text: string.text,
      identifier: string.identifier,
      context: string.context,
      branchId: branchId
    }}
    else{
          payload = {
      text: string.text,
      identifier: string.identifier,
      context: string.context,
      branchId: branchId,
      maxLength: string.max_char
    };
    }
    
    Logger.log(`Making API request to: ${url}`);
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Response code: ${responseCode}`);
    Logger.log(`Response: ${responseText}`);
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      Logger.log(`Success! String ID: ${data.data ? data.data.id : 'Unknown'}`);
      
      return {
        success: true,
        stringId: data.data ? data.data.id : 'Unknown'
      };
    } else {
      throw new Error(`HTTP ${responseCode}: ${responseText}`);
    }
    
  } catch (error) {
    Logger.log(`Ultra-simple upload failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Push a hardcoded string (bypasses all sheet reading - guaranteed fast)
 */
function pushHardcodedString() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    ui.alert('Testing hardcoded string push...', 'This will push a test string without reading any sheet data.', ui.ButtonSet.OK);
    
    // Create a hardcoded string - no sheet reading at all
    const testString = {
      identifier: 'hardcoded_test_' + Date.now(),
      text: 'This is a test string from Google Sheets',
      context: 'Hardcoded test string'
    };
    
    Logger.log(`Starting hardcoded push: ${testString.identifier}`);
    
    // Upload just this hardcoded string
    const result = uploadHardcodedString(testString, config);
    
    if (result.success) {
      ui.alert(
        'Hardcoded String Push Complete! ✅',
        `Successfully pushed: "${testString.text}"\n` +
        `String ID: ${result.stringId}\n` +
        `Identifier: ${testString.identifier}\n\n` +
        `Check Crowdin to see if it appears!`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Hardcoded String Push Failed',
        `Error: ${result.error}`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to push hardcoded string:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Hardcoded string push error: ' + error.stack);
  }
}

/**
 * Upload a hardcoded string (no sheet reading, no loops)
 */
function uploadHardcodedString(string, config) {
  try {
    Logger.log(`Starting hardcoded upload: ${string.identifier}`);
    
    const baseUrl = `https://${CONFIG_KEYS.CROWDIN_ORGANIZATION}.crowdin.com/api/v2`;
    const projectId = config.projectId;
    const branchId = 228360;
    
    const url = `${baseUrl}/projects/${projectId}/strings`;
    const payload = {
      text: string.text,
      identifier: string.identifier,
      context: string.context,
      branchId: branchId
    };
    
    Logger.log(`Making hardcoded API request to: ${url}`);
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`Hardcoded response code: ${responseCode}`);
    Logger.log(`Hardcoded response: ${responseText}`);
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      Logger.log(`Hardcoded success! String ID: ${data.data ? data.data.id : 'Unknown'}`);
      
      return {
        success: true,
        stringId: data.data ? data.data.id : 'Unknown'
      };
    } else {
      throw new Error(`HTTP ${responseCode}: ${responseText}`);
    }
    
  } catch (error) {
    Logger.log(`Hardcoded upload failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Pull all translations from current sheet with batching (fast and complete)
 */
function pullCurrentSheetComplete() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get current active sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    
    ui.alert('Starting complete pull from current sheet...', 'Pulling all translations from: ' + activeSheet.getName(), ui.ButtonSet.OK);
    
    // Process the current sheet with batching
    const result = processSheetForPullBatched(activeSheet, config);
    
    if (result) {
      ui.alert(
        'Pull Complete! ✅',
        'Successfully pulled ' + result.translationCount + ' translations from "' + activeSheet.getName() + '".\n\n' +
        'Errors: ' + result.errorCount + '\n' +
        'Strings processed: ' + result.stringCount + '\n\n' +
        'Check the sheet to see the translations!',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'No data to pull',
        'No "English (US)" row found in sheet "' + activeSheet.getName() + '".',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', 'Failed to pull from current sheet:\n' + error.message, ui.ButtonSet.OK);
    Logger.log('Pull current sheet error: ' + error.stack);
  }
}

/**
 * Pull translations from selected strings only
 */
function pullSelectedStrings() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    // Get the active selection
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const selection = sheet.getSelection();
    
    if (!selection) {
      ui.alert('No Selection', 'Please select some cells first, then run this function.', ui.ButtonSet.OK);
      return;
    }
    
    const ranges = selection.getActiveRangeList().getRanges();
    if (ranges.length === 0) {
      ui.alert('No Selection', 'Please select some cells first, then run this function.', ui.ButtonSet.OK);
      return;
    }
    
    // First, find the English (US) row to establish the source strings
    const data = sheet.getDataRange().getValues();
    let englishRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().includes('English (US)')) {
        englishRowIndex = i;
        break;
      }
    }
    
    if (englishRowIndex === -1) {
      ui.alert('No English Row', 'No "English (US)" row found in the sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Get all selected cells and map them to their English source
    const selectedCells = [];
    ranges.forEach(range => {
      const values = range.getValues();
      const rowStart = range.getRow();
      const colStart = range.getColumn();
      
      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const cellValue = values[row][col];          
          const actualRow = rowStart + row;
          const actualCol = colStart + col;
          const columnLetter = columnToLetter(actualCol);
          
          // Skip English row - we don't need to pull translations for source strings
          if (actualRow === englishRowIndex + 1) {
            continue;
          }
          
          // Find the corresponding English source string in the same column
          const englishSource = data[englishRowIndex][actualCol - 1];
          if (!englishSource || englishSource.toString().trim() === '') {
            continue; // Skip if no English source text
          }
          
          selectedCells.push({
            row: actualRow,
            col: actualCol,
            text: cellValue.toString().trim(),
            columnLetter: columnLetter,
            englishSource: englishSource.toString().trim(),
            identifier: sheet.getName() + '_R' + (englishRowIndex + 1) + 'C' + columnLetter,
            context: 'Sheet: ' + sheet.getName() + ', Cell: ' + columnLetter + actualRow
          });
        }
      }
    });
    
    if (selectedCells.length === 0) {
      ui.alert('No Content', 'No translation cells found in selection (English row excluded).', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert(
      'Ready to Pull Selected Strings', 
      'Found ' + selectedCells.length + ' translation cells\n\n' +
      'This will pull translations for these specific strings from Crowdin.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    // Process selected strings
    let processed = 0;
    let errors = 0;
    let errorMessages = [];

    for (const cell of selectedCells) {
      try {
        // Get the language information for this cell
        const languageLabel = data[cell.row - 1][0]; // Get language from column A (0-indexed)
        const localeCode = data[cell.row - 1][1]; // Get locale from column B (0-indexed)
        
        Logger.log('Processing cell row: ' + cell.row + ', Language: ' + languageLabel + ', Locale: ' + localeCode);
        Logger.log('English source: "' + cell.englishSource + '", Identifier: ' + cell.identifier);
        
        // Map to correct Crowdin language code
        const crowdinLocale = LANGUAGE_MAP[languageLabel.toString().trim()] || localeCode;
        
        if (!crowdinLocale) {
          errors++;
          errorMessages.push('Unknown language: ' + languageLabel + ' for cell ' + cell.columnLetter + cell.row);
          continue;
        }
        
        Logger.log('Looking for translation: ' + cell.identifier + ' in language ' + crowdinLocale);
        
        const translation = getTranslationFromCrowdin(cell.identifier, crowdinLocale, config);
        if (translation) {
          // Write translation to the cell
          sheet.getRange(cell.row, cell.col).setValue(translation);
          processed++;
          Logger.log('✅ Pulled: ' + cell.identifier + ' (' + crowdinLocale + ') = "' + translation + '"');
        } else {
          errors++;
          errorMessages.push('No translation found for: ' + cell.identifier + ' (' + crowdinLocale + ')');
        }
        
        Utilities.sleep(100); // Small delay between strings
        
      } catch (error) {
        errors++;
        errorMessages.push('Error with ' + cell.identifier + ': ' + error.message);
        Logger.log('❌ Error with ' + cell.identifier + ': ' + error.message);
      }
    }
    
    let resultMessage = 'Successfully pulled ' + processed + ' translations.\n\n';
    if (errors > 0) {
      resultMessage += 'Errors (' + errors + '):\n';
      resultMessage += errorMessages.slice(0, 5).join('\n');
      if (errorMessages.length > 5) {
        resultMessage += '\n... and ' + (errorMessages.length - 5) + ' more errors';
      }
    }
    
    ui.alert('Pull Selected Strings Complete!', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', 'Failed to pull selected strings:\n' + error.message, ui.ButtonSet.OK);
    Logger.log('Pull selected strings error: ' + error.stack);
  }
}

/**
 * Pull a single translation to test the pull functionality
 */
function pullSingleTranslation() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    
    ui.alert('Testing single translation pull...', 'This will pull one translation from Crowdin to test the functionality.', ui.ButtonSet.OK);
    
    // Get all strings from Crowdin
    const baseUrl = getCrowdinBaseUrl();
    const projectId = config.projectId;
    const branchId = getBranchId(config);
    
    Logger.log('Getting strings from Crowdin...');
    
    const stringsUrl = `${baseUrl}/projects/${projectId}/strings?branchId=${branchId}&limit=10`;
    const stringsResponse = makeApiRequest(stringsUrl, 'GET', null, config.token);
    
    if (!stringsResponse.data || stringsResponse.data.length === 0) {
      ui.alert('No Strings Found', 'No strings found in Crowdin to pull from.', ui.ButtonSet.OK);
      return;
    }
    
    const firstString = stringsResponse.data[0].data;
    Logger.log(`Found string: ${firstString.identifier}`);
    
    // Get translations for this string
    const translationsUrl = `${baseUrl}/projects/${projectId}/translations/strings/${firstString.id}?languageId=${firstString.languageId}`;
    const translationsResponse = makeApiRequest(translationsUrl, 'GET', null, config.token);
    
    Logger.log(`Found ${translationsResponse.data ? translationsResponse.data.length : 0} translations`);
    
    let translationInfo = `String: "${firstString.text}"\n`;
    translationInfo += `Identifier: ${firstString.identifier}\n\n`;
    
    if (translationsResponse.data && translationsResponse.data.length > 0) {
      translationInfo += `Translations found:\n`;
      translationsResponse.data.forEach(translation => {
        const lang = translation.data.languageId;
        const text = translation.data.text;
        const status = translation.data.approvalStatus;
        translationInfo += `- ${lang}: "${text}" (${status})\n`;
      });
    } else {
      translationInfo += `No translations found yet.`;
    }
    
    ui.alert('Translation Pull Test Results', translationInfo, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Failed to pull translation:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Pull translation error: ' + error.stack);
  }
}

/**
 * Main function to pull translations from Crowdin
 */
function pullFromCrowdin() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    ui.alert('Starting pull from Crowdin...', ui.ButtonSet.OK);
    
    let totalTranslations = 0;
    let processedSheets = 0;
    
    // Process each sheet
    for (const sheet of sheets) {
      const result = processSheetForPull(sheet, config);
      if (result) {
        totalTranslations += result.translationCount;
        processedSheets++;
      }
    }
    
    if (processedSheets === 0) {
      ui.alert(
        'No data to pull',
        'No sheets found with "English (US)" row.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Pull Complete! ✅',
        `Successfully pulled ${totalTranslations} translations from Crowdin into ${processedSheets} sheet(s).\n\n` +
        `Empty cells remain for translators to fill in.`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('Error', `Failed to pull from Crowdin:\n${error.message}`, ui.ButtonSet.OK);
    Logger.log('Pull error: ' + error.stack);
  }
}

/**
 * Processes a single sheet for pulling from Crowdin with column-based batching
 */
function processSheetForPullBatched(sheet, config) {
  const data = sheet.getDataRange().getValues();
  
  // Find the "English (US)" row
  var englishRowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes('English (US)')) {
      englishRowIndex = i;
      break;
    }
  }
  
  if (englishRowIndex === -1) {
    return null;
  }
  
  const englishRow = data[englishRowIndex];
  var translationCount = 0;
  var errorCount = 0;
  var stringCount = 0;
  var errorMessages = [];
  
  // Find the last populated row in the first column (column A)
  var lastPopulatedRow = englishRowIndex + 1;
  for (var i = englishRowIndex + 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '') {
      lastPopulatedRow = i;
    } else {
      break; // Stop when we hit the first empty cell
    }
  }
  
  Logger.log('Processing from row ' + (englishRowIndex + 1) + ' to row ' + lastPopulatedRow);
  
  // Process column by column (starting from column D, index 3)
  for (var col = 3; col < englishRow.length; col++) {
    const sourceText = englishRow[col];
    
    // Skip columns without content in English row
    if (!sourceText || sourceText.toString().trim() === '') {
      continue;
    }
    
    const columnLetter = columnToLetter(col + 1);
    Logger.log('Processing column ' + columnLetter + ' - Source: "' + sourceText.toString().substring(0, 50) + '..."');
    
    // Process all language rows for this column
    for (var row = englishRowIndex + 1; row <= lastPopulatedRow; row++) {
      const languageLabel = data[row][0];
      const localeCode = data[row][1];
      
      if (!languageLabel || languageLabel.toString().trim() === '') {
        continue; // Skip empty rows
      }
      
      // Get the Crowdin language ID
      const crowdinLocale = LANGUAGE_MAP[languageLabel.toString().trim()] || localeCode;
      
      if (!crowdinLocale) {
        Logger.log('Warning: Unknown language "' + languageLabel + '" in sheet ' + sheet.getName());
        continue;
      }
      
      try {
        // Build the string identifier
        let segment_row = englishRowIndex + 1;
        let id = sheet.getName() + '_' + 'R' + segment_row + 'C' + columnLetter;
        
        Logger.log('Fetching translation for: ' + id + ' (' + crowdinLocale + ')');
        
        const translation = getTranslationFromCrowdin(id, crowdinLocale, config);
        if (translation) {
          // Write translation to sheet
          sheet.getRange(row + 1, col + 1).setValue(translation);
          translationCount++;
          Logger.log('✅ Pulled: ' + id + ' (' + crowdinLocale + ') = "' + translation + '"');
        } else {
          errorCount++;
          errorMessages.push('No translation found for: ' + id + ' (' + crowdinLocale + ')');
        }
        
        stringCount++;
      } catch (error) {
        errorCount++;
        errorMessages.push('Error with column ' + columnLetter + ' (' + crowdinLocale + '): ' + error.message);
        Logger.log('❌ Error with column ' + columnLetter + ' (' + crowdinLocale + '): ' + error.message);
      }
      
      // Small delay between individual strings
      Utilities.sleep(100);
    }
    
    // Delay between columns
    Logger.log('Completed column ' + columnLetter + ', moving to next column...');
    Utilities.sleep(200);
  }
  
  // Log error summary
  if (errorMessages.length > 0) {
    Logger.log('Error summary:');
    errorMessages.slice(0, 10).forEach(function(msg) {
      Logger.log('  - ' + msg);
    });
    if (errorMessages.length > 10) {
      Logger.log('  ... and ' + (errorMessages.length - 10) + ' more errors');
    }
  }
  
  return { 
    translationCount: translationCount, 
    errorCount: errorCount, 
    stringCount: stringCount,
    errorMessages: errorMessages
  };
}

/**
 * Processes a single sheet for pulling from Crowdin
 */
function processSheetForPull(sheet, config) {
  const data = sheet.getDataRange().getValues();
  
  // Find the "English (US)" row
  let englishRowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes('English (US)')) {
      englishRowIndex = i;
      break;
    }
  }
  
  if (englishRowIndex === -1) {
    return null;
  }
  
  const englishRow = data[englishRowIndex];
  let translationCount = 0;
  
  // Process each target language row below English
  for (let row = englishRowIndex + 1; row < data.length; row++) {
    const languageLabel = data[row][0];
    const localeCode = data[row][1];
    
    if (!languageLabel || languageLabel.toString().trim() === '') {
      continue; // Skip empty rows
    }
    
    // Get the Crowdin language ID
    const crowdinLocale = LANGUAGE_MAP[languageLabel.toString().trim()] || localeCode;
    
    if (!crowdinLocale) {
      Logger.log(`Warning: Unknown language "${languageLabel}" in sheet ${sheet.getName()}`);
      continue;
    }
    
    // Pull translations for each column
    for (let col = 3; col < englishRow.length; col++) {
      const sourceText = englishRow[col];
      if (sourceText && sourceText.toString().trim() !== '') {
        const columnLetter = columnToLetter(col + 1);
        const stringId = `${sheet.getName()}_${columnLetter}`;
        
        // Get translation from Crowdin
        const translation = getTranslationFromCrowdin(stringId, crowdinLocale, config);
        
        if (translation) {
          // Write translation to sheet
          sheet.getRange(row + 1, col + 1).setValue(translation);
          translationCount++;
        }
      }
    }
  }
  
  return { translationCount };
}

/**
 * Gets translation from Crowdin for a specific string and language
 */
function getTranslationFromCrowdin(stringIdentifier, languageId, config) {
  const baseUrl = getCrowdinBaseUrl();
  const projectId = config.projectId;
  let language = null
  if (languageId == "es-419"){
    language = languageId
  }
  else if (languageId == "zh-CN"){
    language = languageId
  }
  else if (languageId == "zh-TW"){
    language = languageId
  }
  else{
    language = languageId.slice(0,-3)
  }


  try {
    // First, find the string by identifier
    Logger.log('Looking for string identifier: ' + stringIdentifier);
    const string = findStringByIdentifier(stringIdentifier, config);
    if (!string) {
      Logger.log('String not found for identifier: ' + stringIdentifier);
      return null;
    }
    
    Logger.log('Found string with ID: ' + string.id + ', now looking for translations in language: ' + language);
    
    // Get the string's translations - using official Crowdin API v2 endpoint structure
    var url = baseUrl + '/projects/' + projectId +'/translations?stringId=' + string.id + '&languageId=' + language ;
    

    Logger.log('Translation URL: ' + url);
    
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true,
      followRedirects: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('Translation response code: ' + responseCode);
    Logger.log('Translation response: ' + (responseText ? responseText.slice(0, 500) : ''));
    
    if (responseCode >= 200 && responseCode < 300) {
      const data = JSON.parse(responseText);
      if (data.data && data.data.length > 0) {
        Logger.log('Found ' + data.data.length + ' total translations');
        return data.data[0].data.text

      } else {
        Logger.log('No translation data found for string ID: ' + string.id);
      }
    } else {
      Logger.log('Translation request failed: ' + responseCode + ' - ' + responseText);
    }
  } catch (error) {
    Logger.log('Error getting translation for ' + stringIdentifier + ' (' + language + '): ' + error.message);
  }
  
  return null;
}

/**
 * Makes an API request to Crowdin
 */
function makeApiRequest(url, method, payload, token) {
  const options = {
    method: method,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  if (payload) {
    options.payload = JSON.stringify(payload);
  }
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode >= 200 && responseCode < 300) {
    return responseText ? JSON.parse(responseText) : {};
  } else {
    throw new Error(`API request failed (${responseCode}): ${responseText}`);
  }
}

/**
 * Converts column index to letter (1 -> A, 27 -> AA, etc.)
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Test specific API endpoints to see what works
 */
function testApiEndpoints() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = validateConfig();
    const baseUrl = getCrowdinBaseUrl();
    const projectId = config.projectId;
    
    let results = `Testing API endpoints for project ${projectId}:\n\n`;
    
    // Test 1: List all projects first
    try {
      const projectsUrl = baseUrl + '/projects';
      const options = {
        method: 'GET',
        headers: getAuthHeaders(),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      const response = UrlFetchApp.fetch(projectsUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 200 && responseCode < 300) {
        const data = JSON.parse(responseText);
        results += '✅ Projects List API: Working\n';
        results += 'Found ' + (data.data ? data.data.length : 0) + ' projects\n';
        
        if (data.data && data.data.length > 0) {
          results += 'Available projects:\n';
          data.data.forEach(function(project) {
            results += '- ID: ' + project.data.id + ', Name: "' + project.data.name + '"\n';
          });
          results += '\n';
        }
      } else {
        throw new Error('HTTP ' + responseCode + ': ' + responseText);
      }
    } catch (error) {
      results += '❌ Projects List API: Failed - ' + error.message + '\n\n';
    }
    
    // Test 2: Project info
    try {
      const projectUrl = baseUrl + '/projects/' + projectId;
      const options = {
        method: 'GET',
        headers: getAuthHeaders(),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      const response = UrlFetchApp.fetch(projectUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 200 && responseCode < 300) {
        const data = JSON.parse(responseText);
        results += '✅ Project API: Working\n';
        results += 'Project Name: ' + data.data.name + '\n';
        results += 'Project Type: ' + (data.data.type || 'Unknown') + '\n\n';
      } else {
        throw new Error('HTTP ' + responseCode + ': ' + responseText);
      }
    } catch (error) {
      results += '❌ Project API: Failed - ' + error.message + '\n';
      results += 'This means project ' + projectId + ' either doesn\'t exist or you don\'t have access.\n\n';
    }
    
    // Test 3: Strings list
    try {
      const stringsUrl = baseUrl + '/projects/' + projectId + '/strings?limit=10';
      const options = {
        method: 'GET',
        headers: getAuthHeaders(),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      const response = UrlFetchApp.fetch(stringsUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 200 && responseCode < 300) {
        const data = JSON.parse(responseText);
        results += '✅ Strings API: Working\n';
        results += 'Found ' + (data.data ? data.data.length : 0) + ' strings\n\n';
      } else {
        throw new Error('HTTP ' + responseCode + ': ' + responseText);
      }
    } catch (error) {
      results += '❌ Strings API: Failed - ' + error.message + '\n\n';
    }
    
    // Test 4: Get branches
    var branchesResponse = null;
    try {
      const branchesUrl = baseUrl + '/projects/' + projectId + '/branches';
      const options = {
        method: 'GET',
        headers: getAuthHeaders(),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      const response = UrlFetchApp.fetch(branchesUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 200 && responseCode < 300) {
        branchesResponse = JSON.parse(responseText);
        results += '✅ Branches API: Working\n';
        results += 'Found ' + (branchesResponse.data ? branchesResponse.data.length : 0) + ' branches\n';
        
        if (branchesResponse.data && branchesResponse.data.length > 0) {
          results += 'Available branches:\n';
          branchesResponse.data.forEach(function(branch) {
            results += '- ID: ' + branch.data.id + ', Name: "' + branch.data.name + '"\n';
          });
          results += '\n';
        } else {
          results += 'No branches found - will create strings without branchId\n\n';
        }
      } else {
        throw new Error('HTTP ' + responseCode + ': ' + responseText);
      }
    } catch (error) {
      results += '❌ Branches API: Failed - ' + error.message + '\n';
      results += 'Will create strings without branchId\n\n';
    }
    
    // Test 5: Try to create a test string with correct branch and file
    try {
      // Get the branch ID and file ID using our smart functions
      const branchId = getBranchId(config);
      const fileId = getFileId(config);
      
      const testStringUrl = getCrowdinBaseUrl() + '/projects/' + projectId + '/strings';
      var testPayload = {
        text: 'Test string from Google Sheets',
        identifier: 'test_google_sheets_' + Date.now(),
        context: 'API Test',
        fileId: fileId
      };
      
      // Only add branchId if we have a valid one (> 0)
      if (branchId > 0) {
        testPayload.branchId = branchId;
      }
      
      const options = {
        method: 'POST',
        headers: getAuthHeaders(),
        payload: JSON.stringify(testPayload),
        muteHttpExceptions: true,
        followRedirects: true
      };
      
      const response = UrlFetchApp.fetch(testStringUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 200 && responseCode < 300) {
        const data = JSON.parse(responseText);
        results += '✅ Create String API: Working\n';
        results += 'Created string ID: ' + (data.data ? data.data.id : 'Unknown') + '\n';
        results += 'Used branch ID: ' + (branchId > 0 ? branchId : 'none') + '\n';
        results += 'Used file ID: ' + fileId + '\n\n';
      } else {
        throw new Error('HTTP ' + responseCode + ': ' + responseText);
      }
    } catch (error) {
      results += '❌ Create String API: Failed - ' + error.message + '\n\n';
    }
    
    ui.alert('API Endpoint Test Results', results, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Test Failed', `Could not test endpoints: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Test function to verify API connection
 */
function testConnection() {
  const projectId = PropertiesService.getUserProperties().getProperty(CONFIG_KEYS.CROWDIN_PROJECT_ID);
  if (!projectId) throw new Error('No projectId saved. Open Config and save it.');
  
  const url = getCrowdinBaseUrl() + '/projects/' + projectId;
  const options = {
    method: 'get',
    headers: getAuthHeaders(),
    muteHttpExceptions: true,
    followRedirects: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  Logger.log('Crowdin GET /projects response code: ' + code);
  Logger.log('Body: ' + (body ? body.slice(0, 1000) : ''));

  if (code >= 200 && code < 300) {
    const data = JSON.parse(body);
    SpreadsheetApp.getActive().toast('Crowdin connection OK ✅', 'Test Connection', 5);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Connection Successful! ✅',
      'Connected to project: ' + data.data.name + '\n' +
      'Project ID: ' + data.data.id + '\n' +
      'Languages: ' + (data.data.targetLanguages ? data.data.targetLanguages.length : 0),
      ui.ButtonSet.OK
    );
    return true;
  }
  
  throw new Error('Crowdin connection failed (' + code + '): ' + body);
}

/**
 * Debug function to check saved configuration
 */
function debugConfig() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const config = getConfig();
    const props = PropertiesService.getUserProperties();
    
    // Test spreadsheet access
    let spreadsheetInfo = 'Spreadsheet: NOT ACCESSIBLE';
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      spreadsheetInfo = `Spreadsheet: ${ss.getName()}\nSheets: ${ss.getSheets().length}`;
    } catch (error) {
      spreadsheetInfo = `Spreadsheet: ERROR - ${error.message}`;
    }
    
    const debugInfo = 
      `Token length: ${config.token ? config.token.length : 0}\n` +
      `Project ID: ${config.projectId || 'NOT SET'}\n` +
      `Token starts with: ${config.token ? config.token.substring(0, 10) + '...' : 'NOT SET'}\n\n` +
      `${spreadsheetInfo}\n\n` +
      `Raw properties:\n` +
      `- CROWDIN_TOKEN: ${props.getProperty(CONFIG_KEYS.CROWDIN_TOKEN) ? 'SET' : 'NOT SET'}\n` +
      `- CROWDIN_PROJECT_ID: ${props.getProperty(CONFIG_KEYS.CROWDIN_PROJECT_ID) || 'NOT SET'}`;
    
    ui.alert('Debug Information', debugInfo, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Debug Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Test spreadsheet access specifically
 */
function testSpreadsheetAccess() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    let sheetInfo = '';
    for (let i = 0; i < Math.min(sheets.length, 5); i++) {
      sheetInfo += `- ${sheets[i].getName()}\n`;
    }
    if (sheets.length > 5) {
      sheetInfo += `... and ${sheets.length - 5} more`;
    }
    
    ui.alert(
      'Spreadsheet Access Test ✅',
      `Successfully accessed: ${ss.getName()}\n\n` +
      `Sheets found:\n${sheetInfo}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Spreadsheet Access Failed ❌',
      `Error: ${error.message}\n\n` +
      `This means the script cannot access your spreadsheet.\n` +
      `Please ensure you have edit permissions.`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Finds the first cell containing "char max" in a column and returns the numeric value that precedes it
 * @param {number} columnNumber - The column number (1 = A, 2 = B, etc.)
 * @param {string} sheetName - Optional sheet name. If not provided, uses active sheet
 * @return {number|null} - The numeric value that precedes "char max", or null if not found
 */
function findCharMaxValue(columnNumber, sheetName = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Get all values in the specified column
    const lastRow = sheet.getLastRow();
    const columnRange = sheet.getRange(1, columnNumber, lastRow, 1);
    const values = columnRange.getValues();
    
    // Search from top to bottom for the first cell containing "char max"
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      
      if (cellValue && typeof cellValue === 'string') {
        // Check if the cell contains "char max" (case insensitive)
        if (cellValue.toLowerCase().includes('char max')) {
          // Extract numeric value that precedes "char max"
          const match = cellValue.match(/(\d+)\s*char\s*max/i);
          if (match) {
            const numericValue = parseInt(match[1], 10);
            Logger.log(`Found "char max" in column ${columnNumber}, row ${i + 1}: "${cellValue}" - extracted value: ${numericValue}`);
            return numericValue;
          }
        }
      }
    }
    
    Logger.log(`No cell containing "char max" found in column ${columnNumber}`);
    return 0;
    
  } catch (error) {
    Logger.log(`Error in findCharMaxValue: ${error.message}`);
    throw error;
  }
}
