
var SHEET_NAME = "Responses";

var ITEM_IDS = [
  "industry-insights",
  "audience-definition",
  "competitive-strategy",
  "business-outlook",
  "brand-strategy",
  "compelling-messaging",
  "distinctive-visuals",
  "storytelling-strategy",
  "brand-advertising",
  "key-brand-assets",
  "integrated-content",
  "memorable-experiences",
  "performance-advertising",
  "attractive-offers",
  "digital-discovery",
  "compelling-evidence"
];

var ITEM_LABELS = {
  "industry-insights":       "Industry Insights",
  "audience-definition":     "Audience Definition",
  "competitive-strategy":    "Competitive Strategy",
  "business-outlook":        "Business Outlook",
  "brand-strategy":          "Brand Strategy",
  "compelling-messaging":    "Compelling Messaging",
  "distinctive-visuals":     "Distinctive Visuals",
  "storytelling-strategy":   "Storytelling Strategy",
  "brand-advertising":       "Brand Advertising",
  "key-brand-assets":        "Key Brand Assets",
  "integrated-content":      "Integrated Content",
  "memorable-experiences":   "Memorable Experiences",
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

    // Body arrives as text/plain in the form "data=<url-encoded-json>"
    // (this avoids a FormData cloning issue in some sandboxed browser contexts).
    var rawBody = e.postData ? e.postData.contents : "";
    var jsonString;

    if (e.parameter && e.parameter.data) {
      // Form-encoded request (FormData or x-www-form-urlencoded)
      jsonString = e.parameter.data;
    } else if (rawBody.indexOf("data=") === 0) {
      // Plain-text body: strip the "data=" prefix and decode
      jsonString = decodeURIComponent(rawBody.substring(5));
    } else {
      // Fallback: assume the whole body is raw JSON
      jsonString = rawBody;
    }

    var data = JSON.parse(jsonString);

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
    "business-outlook": "5",
    "brand-strategy": "2",
    "compelling-messaging": "3",
    "distinctive-visuals": "4",
    "storytelling-strategy": "1",
    "brand-advertising": "3",
    "key-brand-assets": "4",
    "integrated-content": "2",
    "memorable-experiences": "N/A",
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
