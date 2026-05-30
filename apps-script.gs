// Google Apps Script - Deploy as Web App
// 1. Go to script.google.com
// 2. Copy this code
// 3. Deploy > New deployment > Web app > Execute as [your email]
// 4. Allow permissions and copy the deployment URL

const FOLDER_ID = "1RHqADYpsbyUDBVGKWbOtfoaNOPpHXZMH"; // Your Vehicle Maintenance folder
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID"; // Optional: replace with your sheet ID to log receipts

// Process incoming receipt uploads
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const base64 = data.file;
    const fileName = data.fileName;
    const mimeType = data.mimeType || "image/jpeg";

    // Decode base64 and create file
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      mimeType,
      fileName
    );

    // Upload to Google Drive folder
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);

    // Process receipt (handles both images and PDFs)
    const receiptData = processReceipt(file);

    // Save to Google Sheet (optional)
    if (SHEET_ID && SHEET_ID !== "YOUR_GOOGLE_SHEET_ID") {
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

// Process receipt (image or PDF)
function processReceipt(file) {
  const fileName = file.getName().toLowerCase();
  const isPDF = fileName.endsWith('.pdf');
  
  if (isPDF) {
    return processReceiptPDF(file);
  } else {
    return processReceiptImage(file);
  }
}

// Process image receipt with Claude AI
function processReceiptImage(file) {
  const CLAUDE_API_KEY = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  
  if (!CLAUDE_API_KEY) {
    throw new Error("CLAUDE_API_KEY not set. Add it in Project Settings > Script Properties");
  }
  
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

// Process PDF receipt with Claude AI
function processReceiptPDF(file) {
  const CLAUDE_API_KEY = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  
  if (!CLAUDE_API_KEY) {
    throw new Error("CLAUDE_API_KEY not set. Add it in Project Settings > Script Properties");
  }
  
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
            type: "document",
            source: {
              type: "base64",
              media_type: "application/pdf",
              data: base64
            }
          },
          {
            type: "text",
            text: `Extract the following information from this PDF receipt in JSON format:
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
function processNewReceipts() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  // Get both JPEG and PDF files
  const jpegFiles = folder.getFilesByType(MimeType.JPEG);
  const pdfFiles = folder.getFilesByType(MimeType.PDF);
  
  let processed = 0;

  // Process JPEG files
  while (jpegFiles.hasNext()) {
    const file = jpegFiles.next();
    processed += processFileIfNew(file);
  }

  // Process PDF files
  while (pdfFiles.hasNext()) {
    const file = pdfFiles.next();
    processed += processFileIfNew(file);
  }

  Logger.log(`Total receipts processed: ${processed}`);
}

// Helper function to process file if not already processed
function processFileIfNew(file) {
  // Skip if already processed (check description)
  if (file.getDescription() && file.getDescription().includes("PROCESSED")) {
    return 0;
  }

  try {
    // Process file
    const receiptData = processReceipt(file);

    // Save to sheet
    if (SHEET_ID && SHEET_ID !== "YOUR_GOOGLE_SHEET_ID") {
      saveToSheet(receiptData, file.getUrl());
    }

    // Mark as processed
    file.setDescription("PROCESSED - " + new Date());

    Logger.log(`Processed: ${file.getName()}`);
    return 1;
  } catch (error) {
    Logger.log(`Error processing ${file.getName()}: ${error}`);
    return 0;
  }
}
