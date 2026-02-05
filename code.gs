/* -------------------------------------------------------------------------
   CONFIGURATION
   ------------------------------------------------------------------------- */
const FOLDER_ID = '11BafLeb9gNMS6zPe-5LUmz1Xb5YfF86f'; 
const SHEET_NAME_TRANS = 'DB_Transactions';
const SHEET_NAME_LOGS = 'DB_Logs';
const SHEET_NAME_SETTINGS = 'DB_Settings';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Apartment Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* -------------------------------------------------------------------------
   API: GET DATA
   ------------------------------------------------------------------------- */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);

  if (!transSheet || !settingsSheet) {
    return JSON.stringify({ error: true, message: "Missing Tabs 'DB_Transactions' or 'DB_Settings'." });
  }
  
  // 1. Get Transactions
  const transRange = transSheet.getDataRange();
  let transactions = [];
  
  if (!transRange.isBlank()) {
    const transData = transRange.getValues();
    const transHeaders = transData.shift(); 
    
    transactions = transData.map(row => {
      let obj = {};
      transHeaders.forEach((header, i) => {
        if(header) {
          let value = row[i];
          // Handle Dates safely
          if ((header === 'Date' || header === 'Method_Date') && value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          if ((header === 'Date' || header === 'Method_Date') && !value) value = ""; 
          obj[header] = value;
        }
      });
      return obj;
    }).reverse();
  }

  // 2. Get Settings
  const settings = {};
  const setRange = settingsSheet.getDataRange();
  if (!setRange.isBlank()) {
    const setVals = setRange.getValues();
    setVals.shift(); 
    setVals.forEach(row => {
      const type = row[0];
      const val = row[1];
      if (type && val) {
        if (!settings[type]) settings[type] = [];
        settings[type].push(val);
      }
    });
  }

  return JSON.stringify({ transactions: transactions, settings: settings });
}

/* -------------------------------------------------------------------------
   API: HANDLE FORM SUBMIT
   ------------------------------------------------------------------------- */
function processForm(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
  
  // Auto-Save New Categories
  if (formObject.category) {
    const existingCats = settingsSheet.getDataRange().getValues()
      .filter(r => r[0] === 'Category')
      .map(r => r[1]);
    if (!existingCats.includes(formObject.category)) {
      settingsSheet.appendRow(['Category', formObject.category]);
    }
  }

  // --- MULTI-FILE UPLOAD LOGIC ---
  // 1. Start with existing files (that weren't deleted by the user)
  let finalUrls = formObject.existingFileUrl ? formObject.existingFileUrl.split(',') : [];
  let finalIds = formObject.existingFileId ? formObject.existingFileId.split(',') : [];
  
  // 2. Process NEW uploads (Array of objects)
  // formObject.files will be an array of { data: "base64...", name: "filename" }
  if (formObject.files && Array.isArray(formObject.files)) {
    formObject.files.forEach(fileObj => {
      const result = saveFileToDrive(fileObj.data, fileObj.name);
      finalUrls.push(result.url);
      finalIds.push(result.id);
    });
  }

  // 3. Join back to strings
  const strUrls = finalUrls.join(',');
  const strIds = finalIds.join(',');
  // -------------------------------

  const timestamp = new Date();
  
  if (formObject.recordId) {
    // Edit Mode
    handleEdit(formObject.recordId, formObject, strUrls, strIds, timestamp);
    return { success: true, message: "Record Updated Successfully" };
  } else {
    // New Mode
    const newId = Utilities.getUuid();
    const newRow = [
      newId,
      timestamp,
      formObject.date,
      formObject.user,
      formObject.type,
      formObject.scope,
      formObject.category,
      formObject.amount,
      formObject.description,
      strUrls, // Saved as comma-separated string
      strIds,  // Saved as comma-separated string
      formObject.user,
      formObject.paymentMethod,
      formObject.methodDate,
      formObject.installments
    ];
    transSheet.appendRow(newRow);
    return { success: true, message: "Record Added Successfully" };
  }
}

/* -------------------------------------------------------------------------
   HELPER: HANDLE EDIT
   ------------------------------------------------------------------------- */
function handleEdit(id, form, fileUrl, fileId, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  const logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  
  const data = transSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] == id);
  
  if (rowIndex === -1) throw new Error("Record not found");
  
  const realRowIndex = rowIndex + 1;
  const originalRow = data[rowIndex];
  
  // Archive to Logs (Dynamic Length)
  const logRow = [
    Utilities.getUuid(),
    id,
    timestamp,
    form.user, 
    form.editReason || "No reason provided",
    ...originalRow.slice(2) // Copy all original data columns safely
  ];
  logSheet.appendRow(logRow);
  
  // Update Transaction Sheet
  // Map fields to specific column indices (1-based)
  // Col 3=Date ... Col 12=User, Col 13=Method, Col 14=MethodDate, Col 15=Installments
  const range = transSheet.getRange(realRowIndex, 3, 1, 13); // Grab Date -> Installments
  const vals = [[
    form.date,
    form.user,
    form.type,
    form.scope,
    form.category,
    form.amount,
    form.description,
    fileUrl, // Now contains comma-separated list
    fileId,  // Now contains comma-separated list
    form.user, 
    form.paymentMethod,
    form.methodDate,
    form.installments
  ]];
  
  range.setValues(vals);
}

function saveFileToDrive(base64Data, fileName) {
  try {
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: file.getUrl(), id: file.getId() };
  } catch (e) {
    throw new Error("File Upload Failed: " + e.toString());
  }
}
