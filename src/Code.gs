// LEADGEN AI
// Intelligent Lead Management & Customer Engagement

// CONFIGURATION

const CONFIG = {
  SHEET_NAME: "Form Responses 1",
  GEMINI_MODEL: "gemini-3.8-flash",
  API_KEY_PROPERTY: "GEMINI_API_KEY",

  PRIORITY_THRESHOLDS: {
    HIGH: 75,
    MEDIUM: 45
  }
};

// TEST GMAIL CONNECTION

function verifyMailService() {

  const testEmail = PropertiesService
    .getScriptProperties()
    .getProperty("TEST_EMAIL_ADDRESS");

  if (!testEmail) {
    throw new Error("TEST_EMAIL_ADDRESS is missing.");
  }

  const quota = MailApp.getRemainingDailyQuota();

  Logger.log("GMAIL SERVICE AVAILABLE");

  Logger.log("Test recipient configured.");

  Logger.log("Remaining daily email quota: " + quota);

}


//1: DATABASE CONNECTION

function testConnection() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("Spreadsheet connection failed.");
  }

  Logger.log("LEADGEN AI DATABASE CONNECTED");

  Logger.log("Database: " + spreadsheet.getName());

  Logger.log("Sheets: " + spreadsheet.getSheets().length);

  spreadsheet.getSheets().forEach(function(sheet) {
    Logger.log("Sheet: " + sheet.getName());
  });

}

// 2: HEADER NORMALISATION

function normalizeHeader(value) {

  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

}


function getHeaderMap(sheet) {

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  const map = {};

  headers.forEach(function(header, index) {

    map[normalizeHeader(header)] = index + 1;

  });

  return map;

}


/**
 * Finds an earlier lead with the same email address.
 * Returns the original lead ID and row number, or null if none exists.
 */
function findExistingLeadByEmail(sheet, currentRow, email, headerMap) {

  const normalizedEmail = String(email || "").trim().toLowerCase();

  if (!normalizedEmail || currentRow <= 2) {
    return null;
  }

  const emailColumn = headerMap[normalizeHeader("Email Address")];
  const leadIdColumn = headerMap[normalizeHeader("Lead ID")];
  const duplicateOfColumn = headerMap[normalizeHeader("Duplicate Of")];

  if (!emailColumn || !leadIdColumn) {
    throw new Error("Email Address or Lead ID column is missing.");
  }

  const numberOfPreviousRows = currentRow - 2;

  const previousEmails = sheet
    .getRange(2, emailColumn, numberOfPreviousRows, 1)
    .getDisplayValues();

  const previousLeadIds = sheet
    .getRange(2, leadIdColumn, numberOfPreviousRows, 1)
    .getDisplayValues();

  const previousDuplicateOf = duplicateOfColumn
    ? sheet.getRange(2, duplicateOfColumn, numberOfPreviousRows, 1).getDisplayValues()
    : [];

  // Scan from oldest to newest so the original lead is retained.
  for (let index = 0; index < numberOfPreviousRows; index++) {

    const previousEmail = String(previousEmails[index][0] || "")
      .trim()
      .toLowerCase();

    if (previousEmail !== normalizedEmail) {
      continue;
    }

    const existingLeadId = String(previousLeadIds[index][0] || "").trim();
    const originalLeadId = existingLeadId ||
      (previousDuplicateOf.length
        ? String(previousDuplicateOf[index][0] || "").trim()
        : "");

    if (originalLeadId) {
      return {
        row: index + 2,
        leadId: originalLeadId
      };
    }
  }

  return null;
}

// 3: AUTOMATION SETUP

function setupAutomation() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + CONFIG.SHEET_NAME
    );
  }

  const requiredHeaders = [

    "Lead ID",
    "AI Summary",
    "Lead Score",
    "Lead Priority",
    "AI Reasoning",
    "Email Subject",
    "AI Response",
    "Email Status",
    "Processing Status",
    "Processed At",
    "Duplicate Status",
    "Duplicate Of"

  ];

  let headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  let normalizedHeaders = headers.map(normalizeHeader);

  requiredHeaders.forEach(function(header) {

    if (!normalizedHeaders.includes(normalizeHeader(header))) {

      sheet.getRange(
        1,
        sheet.getLastColumn() + 1
      ).setValue(header);

      normalizedHeaders.push(normalizeHeader(header));

    }

  });

  // Remove duplicate triggers belonging to this handler.

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {

    if (
      trigger.getHandlerFunction() === "handleFormSubmit"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

  // Create exactly one submission trigger.

  ScriptApp.newTrigger("handleFormSubmit")
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  Logger.log("================================");

  Logger.log("LEADGEN AI AUTOMATION READY");

  Logger.log("Database: " + spreadsheet.getName());

  Logger.log("Sheet: " + CONFIG.SHEET_NAME);

  Logger.log("Trigger installed successfully.");

  Logger.log("================================");

}

// 4: AUTOMATIC FORM SUBMISSION

function handleFormSubmit(e) {

  if (!e || !e.range) {

    throw new Error(
      "Please execute this function through a form submission."
    );

  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {

    throw new Error(
      "Another lead is currently being processed."
    );

  }

  try {

    const sheet = e.range.getSheet();

    const row = e.range.getRow();

    if (sheet.getName() !== CONFIG.SHEET_NAME) {

      throw new Error("Unexpected response sheet.");

    }

    const headerMap = getHeaderMap(sheet);

    const rowData = sheet
      .getRange(row, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // Verify required system columns.

    const requiredSystemHeaders = [
      "Lead ID",
      "Processing Status",
      "Email Status"
    ];

    requiredSystemHeaders.forEach(function(header) {

      if (!headerMap[normalizeHeader(header)]) {

        throw new Error(
          "Missing spreadsheet column: " + header
        );

      }

    });

    function getField(name) {

      const column =
        headerMap[normalizeHeader(name)];

      return column
        ? String(rowData[column - 1] || "").trim()
        : "";

    }

    function updateField(name, value) {

      const column =
        headerMap[normalizeHeader(name)];

      if (column) {

        sheet.getRange(row, column).setValue(value);

      }

    }

    // Prevent duplicate processing.

    const currentStatus = getField("Processing Status");

    if (
      currentStatus === "Completed" ||
      currentStatus === "Processing"
    ) {

      Logger.log("Lead already processed or in progress.");

      return;

    }

    updateField("Processing Status", "Processing");

    // Extract customer information.

    const lead = {

      name: getField("Full Name"),

      email: getField("Email Address"),

      phone: getField("Phone Number"),

      company: getField("Company Name"),

      service: getField("Service Required"),

      budget: getField("Estimated Budget"),

      requirements:
        getField("Describe Your Requirements"),

      contactMethod:
        getField("Preferred Contact Method")

    };

    // Check for an earlier submission using the same email address.
    const duplicateLead = findExistingLeadByEmail(
      sheet,
      row,
      lead.email,
      headerMap
    );

    if (duplicateLead) {

      updateField("Duplicate Status", "Duplicate");
      updateField("Duplicate Of", duplicateLead.leadId);
      updateField("Processing Status", "Duplicate");
      updateField("Email Status", "Skipped - Duplicate");

      Logger.log("Duplicate submission detected.");
      Logger.log("Duplicate row: " + row);
      Logger.log("Original lead row: " + duplicateLead.row);
      Logger.log("Original Lead ID: " + duplicateLead.leadId);
      Logger.log("Gemini analysis and email delivery skipped.");

      return;
    }

    // Create a Lead ID only after confirming this is a new lead.
    let leadId = getField("Lead ID");

    if (!leadId) {

      leadId =
        "LEAD-" + Utilities.getUuid().substring(0, 8).toUpperCase();

      updateField("Lead ID", leadId);

    }

    updateField("Duplicate Status", "Unique");
    updateField("Duplicate Of", "");

    Logger.log("Processing Lead: " + leadId);

    // Debug data extraction without logging personal details.

    Logger.log("Customer Data Extraction:");

    Object.keys(lead).forEach(function(key) {

      Logger.log(
        key + ": " +
        (lead[key] ? "Received" : "Missing")
      );

    });

    // Send information to Gemini.

    const analysis = analyzeLeadWithGemini(lead);

    // Save AI-generated results.

    updateField("AI Summary", analysis.summary);

    updateField("Lead Score", analysis.leadScore);

    updateField("Lead Priority", analysis.priority);

    updateField("AI Reasoning", analysis.reasoning);

    updateField("Email Subject", analysis.emailSubject);

    updateField("AI Response", analysis.emailBody);

    sendLeadEmail(lead, analysis, updateField);

    updateField("Processing Status", "Completed");

    updateField("Processed At", new Date());

    Logger.log("================================");

    Logger.log("LEAD PROCESSING SUCCESSFUL");

    Logger.log("Lead ID: " + leadId);

    Logger.log("Lead Score: " + analysis.leadScore);

    Logger.log("Priority: " + analysis.priority);

    Logger.log("================================");

  }

  catch (error) {

    const sheet = e.range.getSheet();

    const headerMap = getHeaderMap(sheet);

    const statusColumn =
      headerMap[normalizeHeader("Processing Status")];

    if (statusColumn) {

      sheet.getRange(e.range.getRow(), statusColumn)
        .setValue(
          "Failed: " + String(error.message).substring(0, 150)
        );

    }

    Logger.log("AUTOMATION ERROR: " + error.message);

    throw error;

  }

  finally {

    lock.releaseLock();

  }

}

// MODULE 5: GEMINI AI CONNECTION

function testGeminiConnection() {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.API_KEY_PROPERTY);

  if (!apiKey) {

    throw new Error("Gemini API key not configured.");

  }

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent";

  const payload = {

    contents: [{
      parts: [{
        text:
          "Introduce yourself as the AI engine of LeadGen AI."
      }]
    }]

  };

  const response = UrlFetchApp.fetch(url, {

    method: "post",

    contentType: "application/json",

    headers: {
      "x-goog-api-key": apiKey
    },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true

  });

  Logger.log("HTTP STATUS: " + response.getResponseCode());

  if (response.getResponseCode() !== 200) {

    throw new Error(response.getContentText());

  }

  const result = JSON.parse(response.getContentText());

  Logger.log(
    result.candidates[0].content.parts[0].text
  );

}

// 6: AI LEAD INTELLIGENCE ENGINE

function analyzeLeadWithGemini(lead) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.API_KEY_PROPERTY);

  if (!apiKey) {

    throw new Error("Gemini API key missing.");

  }

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL +
    ":generateContent";

  const prompt = `

You are LeadGen AI, an intelligent business
lead management assistant.

Analyse the customer enquiry below.

CUSTOMER INFORMATION:

Name: ${lead.name || "Not provided"}

Email: ${lead.email || "Not provided"}

Company: ${lead.company || "Not provided"}

Service: ${lead.service || "Not provided"}

Budget: ${lead.budget || "Not provided"}

Requirements: ${lead.requirements || "Not provided"}

Preferred Contact: ${lead.contactMethod || "Not provided"}


YOUR TASK:

1. Summarise the customer's requirements.

2. Assign a lead score from 0 to 100.

3. Classify the lead:
   High: 75-100
   Medium: 45-74
   Low: 0-44

4. Explain the scoring decision.

5. Generate a personalised professional email.

SCORING CRITERIA:

Budget suitability: 25 points

Requirement clarity: 25 points

Service relevance: 20 points

Business potential: 20 points

Readiness to proceed: 10 points

IMPORTANT RULES:

- Never invent customer information.
- Do not assume missing information.
- Do not treat missing information as evidence of low customer quality.
- Explain uncertainty where information is insufficient.
- Keep the response professional.
- Do not promise services, prices or delivery dates.
- Generate a helpful response based on the actual enquiry.
- The email must address the customer by name when available.

RETURN ONLY VALID JSON:

{
  "summary": "Customer requirement summary",
  "leadScore": 70,
  "reasoning": "Explanation of scoring",
  "emailSubject": "Email subject",
  "emailBody": "Complete personalised email"
}

`;

  const payload = {

    contents: [{
      parts: [{
        text: prompt
      }]
    }],

    generationConfig: {

      temperature: 0.3,

      responseMimeType: "application/json"

    }

  };

  // Retry temporary server errors.

  let response;

  let status;

  for (let attempt = 1; attempt <= 4; attempt++) {

    response = UrlFetchApp.fetch(url, {

      method: "post",

      contentType: "application/json",

      headers: {
        "x-goog-api-key": apiKey
      },

      payload: JSON.stringify(payload),

      muteHttpExceptions: true

    });

    status = response.getResponseCode();

    Logger.log("Gemini Attempt: " + attempt);

    Logger.log("HTTP Status: " + status);

    if (status === 200) {

      break;

    }

    // Retry only temporary server errors.

    if (status !== 500 && status !== 503) {

      break;

    }

    if (attempt < 4) {

      Utilities.sleep(
        2000 * Math.pow(2, attempt - 1)
      );

    }

  }

  if (status !== 200) {

    throw new Error(
      "Gemini API Error: " + response.getContentText()
    );

  }

  const result = JSON.parse(response.getContentText());

  if (
    !result.candidates ||
    !result.candidates[0] ||
    !result.candidates[0].content ||
    !result.candidates[0].content.parts
  ) {

    throw new Error("Gemini returned an empty response.");

  }

  const aiText =
    result.candidates[0].content.parts
      .map(function(part) {
        return part.text || "";
      })
      .join("");

  const analysis = JSON.parse(aiText);

  // Validate the AI-generated score.

  const score = Number(analysis.leadScore);

  if (
    !Number.isFinite(score) ||
    score < 0 ||
    score > 100
  ) {

    throw new Error("Invalid AI lead score.");

  }

  // Determine priority using fixed business rules.

  let priority;

  if (score >= CONFIG.PRIORITY_THRESHOLDS.HIGH) {

    priority = "High";

  }

  else if (score >= CONFIG.PRIORITY_THRESHOLDS.MEDIUM) {

    priority = "Medium";

  }

  else {

    priority = "Low";

  }

  // Validate generated content.

  if (
    !analysis.summary ||
    !analysis.emailSubject ||
    !analysis.emailBody
  ) {

    throw new Error("Incomplete AI response.");

  }

  return {

    summary: analysis.summary,

    leadScore: score,

    priority: priority,

    reasoning: analysis.reasoning || "No explanation provided.",

    emailSubject: analysis.emailSubject,

    emailBody: analysis.emailBody

  };

}

// 7: TEST AI ANALYSIS

function testLeadAnalysis() {

  const sampleLead = {

    name: "Priya Mehta",

    email: "priya@example.com",

    phone: "9876543210",

    company: "BrightTech Solutions",

    service: "AI Automation",

    budget: "₹50,000–₹1,00,000",

    requirements:
      "We need an AI-powered customer support system " +
      "that can automatically answer customer enquiries.",

    contactMethod: "Email"

  };

  const result = analyzeLeadWithGemini(sampleLead);

  Logger.log("================================");

  Logger.log("AI LEAD ANALYSIS RESULTS");

  Logger.log("================================");

  Logger.log("Summary: " + result.summary);

  Logger.log("Score: " + result.leadScore);

  Logger.log("Priority: " + result.priority);

  Logger.log("Reasoning: " + result.reasoning);

  Logger.log("Email Subject: " + result.emailSubject);

  Logger.log("Email Body: " + result.emailBody);

  Logger.log("================================");

}

// 8: AUTOMATED EMAIL DELIVERY

function sendLeadEmail(lead, analysis, updateField) {

  const properties = PropertiesService
    .getScriptProperties();

  // Testing mode: emails go only to your inbox.
  const testMode = true;

  const testEmail = properties
    .getProperty("TEST_EMAIL_ADDRESS");

  const recipient = testMode
    ? testEmail
    : lead.email;

  // Validate recipient
  if (!recipient || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
    throw new Error("Invalid or missing recipient email.");
  }

  // Check available daily quota
  if (MailApp.getRemainingDailyQuota() < 1) {
    throw new Error("Daily email sending quota exhausted.");
  }

  // Prevent duplicate sending
  updateField("Email Status", "Sending");

  try {

    MailApp.sendEmail({

      to: recipient,

      subject: analysis.emailSubject,

      body: analysis.emailBody,

      name: "LeadGen AI"

    });

    updateField(
      "Email Status",
      testMode ? "Test Sent" : "Sent"
    );

    Logger.log("EMAIL SENT SUCCESSFULLY");

    Logger.log("Recipient: " + recipient);

  }

  catch (error) {

    updateField(
      "Email Status",
      "Failed: " + error.message.substring(0, 100)
    );

    throw error;

  }

}


