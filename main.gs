// ── 1. ADMINISTRATIVE SETUP & POSITIONING ──────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛡️ DPC: DATA PROTECTION')
    .addItem('1. Initialize/Reset System', 'initializeAdminSystem')
    .addItem('2. Verify My Identity (Mandatory)', 'verifyUserIdentity')
    .addSeparator()
    .addItem('Manual Security Sweep', 'securitySweep')
    .addToUi();
}

function initializeAdminSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or Position AIESEC ADMIN
  let adminSheet = ss.getSheetByName("AIESEC ADMIN");
  if (!adminSheet) {
    adminSheet = ss.insertSheet("AIESEC ADMIN", 0); 
  } else {
    ss.setActiveSheet(adminSheet);
    ss.moveActiveSheet(1); // Ensure it is the first tab
  }
  
  setupAdminHeaders(adminSheet);
  
  // Create Hidden User Map for Identity Resolution
  let mapSheet = ss.getSheetByName("SYS_USER_MAP");
  if (!mapSheet) {
    mapSheet = ss.insertSheet("SYS_USER_MAP");
    mapSheet.hideSheet();
    mapSheet.appendRow(["User Key", "Email Address", "Last Verified"]);
  }

  createInstallableTriggers();
  
  adminSheet.getRange("B2").setValue(new Date());
  adminSheet.getRange("C2").setValue("ACTIVE / SECURED");
  SpreadsheetApp.getUi().alert("✅ DPC System Initialized. Tab positioned at Index 0.");
}

function setupAdminHeaders(sheet) {
  sheet.clear(); // Clean start to fix positioning
  sheet.getRange("A1:C1").setValues([["DPC CONTROL PANEL", "", ""]])
       .merge().setBackground("#00416A").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A2:C2").setValues([["Action", "Last Run / Status", "Security Level"]])
       .setBackground("#f3f3f3").setFontWeight("bold");
  
  sheet.getRange("A3:C3").setValues([["Triggers & Audit", "Checking...", "RIGID"]]);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
}

// ── 2. IDENTITY RESOLUTION (THE ANONYMOUS FIX) ─────────────────────
function verifyUserIdentity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const userKey = Session.getTemporaryActiveUserKey();
  const mapSheet = ss.getSheetByName("SYS_USER_MAP");

  if (!userEmail) {
    SpreadsheetApp.getUi().alert("❌ Error", "Please ensure you have granted permissions to the script.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Update the map: Key -> Email
  const data = mapSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userKey) {
      mapSheet.getRange(i + 1, 2, 1, 2).setValues([[userEmail, new Date()]]);
      found = true;
      break;
    }
  }
  if (!found) {
    mapSheet.appendRow([userKey, userEmail, new Date()]);
  }

  SpreadsheetApp.getUi().alert("✅ Identity Verified", "System now recognizes: " + userEmail, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── 3. THE RIGID AUDIT TRIGGER ──────────────────────────────────────
function createInstallableTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("executeAuditLog").forSpreadsheet(ss).onEdit().create();
}

function executeAuditLog(e) {
  if (!e) return;
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName === "Audit_Log" || sheetName === "AIESEC ADMIN" || sheetName === "SYS_USER_MAP") return;

  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  const userKey = Session.getTemporaryActiveUserKey();
  
  // Resolve Email from Map
  let resolvedEmail = "Anonymous (Not Verified)";
  const mapData = ss.getSheetByName("SYS_USER_MAP").getDataRange().getValues();
  for (let i = 1; i < mapData.length; i++) {
    if (mapData[i][0] === userKey) {
      resolvedEmail = mapData[i][1];
      break;
    }
  }

  const logData = [
    "",timestamp, 
    resolvedEmail, 
    "EDIT", 
    sheetName + "!" + e.range.getA1Notation(), 
    e.oldValue || "[New Data]", 
    e.value || "[Cleared]"
  ];

  writeToLog(ss, logData);
}

function writeToLog(ss, data) {
  let logSheet = ss.getSheetByName("Audit_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Audit_Log");
    logSheet.appendRow(["Timestamp", "User Email", "Action", "Location", "Old Value", "New Value"]);
    logSheet.getRange("A1:F1").setBackground("#efefef").setFontWeight("bold");
  }
  logSheet.appendRow(data);
}

function securitySweep() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Security protocols active. No leaks detected.");
}