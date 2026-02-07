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
   API: HANDLE BATCH SUBMIT (New)
   ------------------------------------------------------------------------- */
function processBatch(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
  
  const timestamp = new Date();
  const newIds = [];

  // payload.transactions is an array of objects
  payload.transactions.forEach(t => {
    
    // 1. Save Category if new (only for Expenses)
    if (t.category && t.type === 'Expense') {
      const existingCats = settingsSheet.getDataRange().getValues()
        .filter(r => r[0] === 'Category')
        .map(r => r[1]);
      if (!existingCats.includes(t.category) && t.category !== 'Fee Settlement') {
        settingsSheet.appendRow(['Category', t.category]);
      }
    }

    // 2. Handle File Uploads (Base64 -> Drive)
    let strUrls = t.existingFileUrl || "";
    let strIds = t.existingFileId || "";
    
    if (t.files && Array.isArray(t.files)) {
      const newUrls = [];
      const newIdsArr = [];
      t.files.forEach(f => {
        const result = saveFileToDrive(f.data, f.name);
        newUrls.push(result.url);
        newIdsArr.push(result.id);
      });
      // Append to existing
      strUrls = strUrls ? strUrls + "," + newUrls.join(',') : newUrls.join(',');
      strIds = strIds ? strIds + "," + newIdsArr.join(',') : newIdsArr.join(',');
    }

    // 3. Create Main Record
    const mainId = Utilities.getUuid();
    newIds.push(mainId);
    
    const rowData = [
      mainId,
      timestamp,
      t.date,
      t.user,
      t.type,          // 'Expense' or 'Payment'
      t.scope,         // 'Joint Balance' or 'Owner Only'
      t.category,
      t.amount,
      t.description,
      strUrls,
      strIds,
      t.user, // Creator
      t.paymentMethod || "",
      t.methodDate || "",
      t.installments || ""
    ];
    transSheet.appendRow(rowData);

    // 4. LOGIC: Owner Only Dual Transaction
    // If it's an Expense marked "Owner Only", we immediately inject a "Payment" to settle it.
    if (t.type === 'Expense' && t.scope === 'Owner Only') {
      const shadowId = Utilities.getUuid();
      const shadowRow = [
        shadowId,
        timestamp,
        t.date,
        t.user,
        'Payment',       // Flip to Payment
        'Owner Only',    // Keep Scope
        'Owner Cover',   // System Category
        t.amount,        // Same Amount
        '[Auto] Offset for: ' + t.description,
        '', '',          // No files needed for shadow
        t.user,
        t.paymentMethod, // Use the method selected (e.g. Cash/Instapay)
        t.methodDate || "",
        t.installments || ""
      ];
      transSheet.appendRow(shadowRow);
    }
  });

  return { success: true, message: "Saved " + newIds.length + " record(s)." };
}

/* -------------------------------------------------------------------------
   API: HANDLE EDIT (Legacy/Single)
   ------------------------------------------------------------------------- */
function processForm(formObject) {
  // Wrapper to redirect single edits to the batch processor logic or handle separately
  // For safety, we keep the original logic for EDITS only.
  
  if (!formObject.recordId) return { success: false, message: "Use processBatch for new records" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  // ... (Reuse handleEdit logic) ...
  
  // Re-process files for Edit
  let strUrls = formObject.existingFileUrl || "";
  let strIds = formObject.existingFileId || "";
  
  if (formObject.files && Array.isArray(formObject.files)) {
      const newUrls = [];
      const newIdsArr = [];
      formObject.files.forEach(f => {
        const res = saveFileToDrive(f.data, f.name);
        newUrls.push(res.url);
        newIdsArr.push(res.id);
      });
      strUrls = strUrls ? strUrls + "," + newUrls.join(',') : newUrls.join(',');
      strIds = strIds ? strIds + "," + newIdsArr.join(',') : newIdsArr.join(',');
  }

  handleEdit(formObject.recordId, formObject, strUrls, strIds, new Date());
  return { success: true, message: "Record Updated" };
}

/* -------------------------------------------------------------------------
   CORE: HANDLE EDIT (Cell-Level Update & Highlight)
   ------------------------------------------------------------------------- */
function handleEdit(recordId, form, strUrls, strIds, editTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAME_TRANS);
  const logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  
  const data = transSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] === recordId); // Find row by ID
  
  if (rowIndex === -1) throw new Error("Record not found");
  
  const sheetRowIndex = rowIndex + 1; // 1-based row for Sheet operations
  const originalRow = data[rowIndex];
  
  // Define Column Mapping (0-based Index in Array -> Field Name in Form)
  const map = {
    2: 'date',
    3: 'user', 
    6: 'category',
    7: 'amount',
    8: 'description',
    9: 'fileUrl', 
    10: 'fileId', 
    12: 'paymentMethod',
    13: 'methodDate',
    14: 'installments'
  };

  const updates = []; 
  
  // 1. Check mapped fields
  for (const [colIdxStr, fieldKey] of Object.entries(map)) {
    const colIdx = parseInt(colIdxStr);
    const oldVal = originalRow[colIdx];
    let newVal;

    if (fieldKey === 'fileUrl') newVal = strUrls;
    else if (fieldKey === 'fileId') newVal = strIds;
    else if (fieldKey === 'user') newVal = form.user || oldVal; 
    else newVal = form[fieldKey]; 

    let isDifferent = false;
    
    // Normalize & Compare
    if (newVal instanceof Date || oldVal instanceof Date || fieldKey.toLowerCase().includes('date')) {
       const d1 = newVal ? new Date(newVal).setHours(0,0,0,0) : 'null';
       const d2 = oldVal ? new Date(oldVal).setHours(0,0,0,0) : 'null';
       if (d1 !== d2) isDifferent = true;
       if(isDifferent && newVal) newVal = Utilities.formatDate(new Date(newVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
    } 
    else if (fieldKey === 'amount' || fieldKey === 'installments') {
       if (parseFloat(newVal || 0) != parseFloat(oldVal || 0)) isDifferent = true;
    }
    else {
       if (String(newVal || "") !== String(oldVal || "")) isDifferent = true;
    }

    if (isDifferent && newVal !== undefined) {
      updates.push({ col: colIdx + 1, val: newVal }); // +1 for Sheet Column
    }
  }

  // 2. Execute Updates
  if (updates.length > 0) {
    
    // A. LOGGING
    // We use .slice(2) to skip [ID, Timestamp] and align with headers [Date, User, Type...]
    const logRow = [
      Utilities.getUuid(),
      recordId,
      editTime,
      form.user,
      form.editReason || "Edit",
      ...originalRow.slice(2) 
    ];
    
    logSheet.appendRow(logRow);
    const lastLogRaw = logSheet.getLastRow();
    
    // B. APPLY UPDATES & HIGHLIGHTS
    updates.forEach(u => {
       // 1. Update Transaction Sheet
       transSheet.getRange(sheetRowIndex, u.col).setValue(u.val).setBackground("#ffff00");
       
       // 2. Highlight Old Value in Log Sheet
       // Logic: Log Data starts at Col 6 (Index 1). 
       // Date is Sheet Col 3. We want Date to land in Log Col 6.
       // Formula: 3 + SheetCol = LogCol (e.g., 3 + 3 = 6)
       logSheet.getRange(lastLogRaw, 3 + u.col).setBackground("#ffff00");
    });

    // C. Update Timestamp
    transSheet.getRange(sheetRowIndex, 2).setValue(editTime);
  }
}

/* -------------------------------------------------------------------------
   API: GET HISTORY
   ------------------------------------------------------------------------- */
function getRecordHistory(recordId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!logSheet) return JSON.stringify([]);

  const data = logSheet.getDataRange().getValues();
  // Log Structure: [LogUUID, RecordID, EditTime, Editor, EditReason, ...OriginalData]
  // OriginalData starts from the specific columns sliced in handleEdit
  
  const history = data
    .filter(row => row[1] === recordId) // Match Record ID
    .map(row => {
      // Map relevant original values (Indices based on handleEdit slice logic)
      return {
        editDate: Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
        editor: row[3],
        reason: row[4],
        data: {
          date: Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          user: row[6],
          type: row[7],
          category: row[9],
          amount: row[10],
          description: row[11],
          files: row[12]
        }
      };
    })
    .reverse(); // Newest edits first

  return JSON.stringify(history);
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
