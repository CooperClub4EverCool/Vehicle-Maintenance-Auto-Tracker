// Google Apps Script - Deploy as Web App
// 1. Go to script.google.com
// 2. Copy this code
// 3. Deploy > New deployment > Web app > Execute as [your email]
// 4. Allow permissions and copy the deployment URL

const FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID"; // Replace with your folder ID
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID"; // Replace with your sheet ID (optional)

// Process incoming receipt uploads
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const base64 = data.file;
    const fileName = data.fileName;

    // Decode base64 and create file
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      "image/jpeg",
      fileName
    );

    // Upload to Google Drive folder
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);

    // Process with Claude AI (via Anthropic API)
    const receiptData = processReceiptWithClaude(file);

    // Save to Google Sheet (optional)
    if (SHEET_ID) {
      saveToSheet(receiptData, file.getUrl());
    }

    return ContentService.createTextOutput(
      JSON.stringify(receiptData)
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Process receipt with Claude AI
function processReceiptWithClaude(file) {
  const CLAUDE_API_KEY = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  const payload = {
    model: "claude-3-5-sonnet-20241022",
    max_tokens: 1024,
    messages: [
      {
        role: "user",
        content: [
          {
            type: "image",
            source: {
              type: "base64",
              media_type: "image/jpeg",
              data: base64
            }
          },
          {
            type: "text",
            text: `Extract the following information from this receipt in JSON format:
{
  "vendor": "store name",
  "date": "YYYY-MM-DD",
  "amount": "numeric amount",
  "category": "vehicle maintenance category (oil change, tire, repair, fuel, wash, other)",
  "items": ["item1", "item2"],
  "confidence": 0.95
}

Be precise and extract only what you see. If unsure, set confidence lower.`
          }
        ]
      }
    ]
  };

  const options = {
    method: "post",
    headers: {
      "x-api-key": CLAUDE_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
  const result = JSON.parse(response.getContentText());

  if (result.content && result.content[0]) {
    const text = result.content[0].text;
    return JSON.parse(text);
  }

  return { error: "Failed to process receipt", confidence: 0 };
}

// Save receipt data to Google Sheet
function saveToSheet(receiptData, fileUrl) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  
  sheet.appendRow([
    new Date(),
    receiptData.vendor,
    receiptData.date,
    receiptData.amount,
    receiptData.category,
    (receiptData.items || []).join(", "),
    receiptData.confidence,
    fileUrl
  ]);
}

// Monitor Google Drive folder for new files (run as trigger)
function processNewRecepts() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType(MimeType.JPEG);

  while (files.hasNext()) {
    const file = files.next();
    
    // Skip if already processed (check metadata)
    if (file.getDescription().includes("PROCESSED")) {
      continue;
    }

    // Process file
    const receiptData = processReceiptWithClaude(file);

    // Save to sheet
    if (SHEET_ID) {
      saveToSheet(receiptData, file.getUrl());
    }

    // Mark as processed
    file.setDescription("PROCESSED");

    Logger.log(`Processed: ${file.getName()}`);
  }
}
