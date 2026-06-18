// ============================================================
// Clarity & Reach — Google Sheets Receiver
// ============================================================
//
// SETUP INSTRUCTIONS:
//
// 1. Open Google Sheets and create a new spreadsheet.
//    Name it something like "Clarity & Reach Responses".
//
// 2. In the spreadsheet, go to Extensions > Apps Script.
//
// 3. Delete any existing code and paste this entire file in.
//
// 4. Click Save (Ctrl+S / Cmd+S).
//
// 5. Click Deploy > New deployment.
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
//    Click Deploy and authorize when prompted.
//
// 6. Copy the Web App URL that appears after deployment.
//
// 7. In clarity-reach-diagnostic.html, find this line near the bottom:
//      var SHEETS_URL = "YOUR_APPS_SCRIPT_URL_HERE";
//    Replace the placeholder with your copied URL.
//
// 8. Push the updated HTML to your GitHub repo. Done.
//
// The sheet will auto-create a header row on first submission.
// Each submission adds one new row.
// ============================================================

var SHEET_NAME = "Responses";

var ITEM_IDS = [
  "industry-insights",
  "audience-definition",
  "competitive-strategy",
  "growth-gameplan",
  "brand-strategy",
  "compelling-messaging",
  "distinctive-visuals",
  "storytelling-strategy",
  "brand-advertising",
  "key-brand-assets",
  "integrated-content",
  "memorable-moments",
  "performance-advertising",
  "attractive-offers",
  "digital-discovery",
  "compelling-evidence"
];

var ITEM_LABELS = {
  "industry-insights":       "Industry Insights",
  "audience-definition":     "Audience Definition",
  "competitive-strategy":    "Competitive Strategy",
  "growth-gameplan":         "Growth Gameplan",
  "brand-strategy":          "Brand Strategy",
  "compelling-messaging":    "Compelling Messaging",
  "distinctive-visuals":     "Distinctive Visuals",
  "storytelling-strategy":   "Storytelling Strategy",
  "brand-advertising":       "Brand Advertising",
  "key-brand-assets":        "Key Brand Assets",
  "integrated-content":      "Integrated Content",
  "memorable-moments":       "Memorable Moments",
  "performance-advertising": "Performance Advertising",
  "attractive-offers":       "Attractive Offers",
  "digital-discovery":       "Digital Discovery",
  "compelling-evidence":     "Compelling Evidence"
};

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    var data = JSON.parse(e.postData.contents);

    // Build header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = ["Name", "Organization", "Submitted At"];
      ITEM_IDS.forEach(function(id) {
        headers.push(ITEM_LABELS[id] || id);
      });
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Build data row
    var row = [
      data.userName || "",
      data.orgName || "",
      data.submittedAt || new Date().toISOString()
    ];
    ITEM_IDS.forEach(function(id) {
      row.push(data[id] !== undefined ? data[id] : "");
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test this function manually in the Apps Script editor
// to confirm it can write to your sheet before deploying.
function testWrite() {
  var testData = {
    userName: "Test User",
    orgName: "Test Org",
    submittedAt: new Date().toISOString(),
    "industry-insights": "4",
    "audience-definition": "3",
    "competitive-strategy": "N/A",
    "growth-gameplan": "5",
    "brand-strategy": "2",
    "compelling-messaging": "3",
    "distinctive-visuals": "4",
    "storytelling-strategy": "1",
    "brand-advertising": "3",
    "key-brand-assets": "4",
    "integrated-content": "2",
    "memorable-moments": "N/A",
    "performance-advertising": "3",
    "attractive-offers": "4",
    "digital-discovery": "5",
    "compelling-evidence": "3"
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    var headers = ["Name", "Organization", "Submitted At"];
    Object.keys(ITEM_LABELS).forEach(function(id) {
      headers.push(ITEM_LABELS[id]);
    });
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }

  var row = [testData.userName, testData.orgName, testData.submittedAt];
  ITEM_IDS.forEach(function(id) { row.push(testData[id] || ""); });
  sheet.appendRow(row);
  Logger.log("Test row written successfully.");
}
