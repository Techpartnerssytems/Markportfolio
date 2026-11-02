/**
 * ============================================================================
 * MARK DALUMPINES — AI AUTOMATION SPECIALIST & GHL EXPERT
 * CONTACT FORM INGESTION & DUAL AUTOMATION DISPATCH SYSTEM
 * ============================================================================
 * Lead Systems Architect: Mark Dalumpines
 * Title: AI Automation Specialist & GHL Expert
 * Primary Notification Email: Dtpsystems2026@gmail.com
 * Direct Email: dalumpinesmark8@gmail.com
 *
 * Description:
 *  1. Captures prospect inquiries from index.html (Enhanced Contact Form Column 1).
 *  2. Appends complete lead records into Google Sheets (Database).
 *  3. Dispatches automated branded HTML email notifications to BOTH parties:
 *     - Admin Notification -> Dtpsystems2026@gmail.com (ReceivedEmailMark.html)
 *     - Lead Confirmation -> Prospect Client Email (SendEmailLead.html)
 * ============================================================================
 */

// GLOBAL CONFIGURATION SETTINGS
const CONFIG = {
  // If container-bound (script attached to a Google Sheet), leave SPREADSHEET_ID empty.
  // If standalone script, paste your Google Sheet ID inside the quotes:
  SPREADSHEET_ID: '',

  // Target sheet tab name:
  SHEET_NAME: 'Inquiries',

  // Admin Notification Email for Mark Dalumpines:
  ADMIN_EMAIL: 'Dtpsystems2026@gmail.com',

  // Sender Branding Names:
  ADMIN_NAME: 'Mark Dalumpines',
  ADMIN_TITLE: 'AI Automation Specialist & GHL Expert',
  SENDER_DISPLAY_NAME: 'Mark Dalumpines | AI Automation Specialist & GHL Expert',

  // Timezone for logging timestamps:
  TIMEZONE: 'Asia/Manila'
};

/**
 * Handle HTTP POST requests from the website contact form
 */
function doPost(e) {
  try {
    let data = {};

    // 1. Parse incoming payload (handles JSON text/plain from fetch or standard URL encoded form)
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // 2. Normalize and sanitize data fields
    const sanitizedData = {
      name: (data.name || 'Not Provided').trim(),
      email: (data.email || '').trim().toLowerCase(),
      phone: (data.phone || 'N/A').trim(),
      businessName: (data.businessName || 'Not Provided').trim(),
      businessIndustry: (data.businessIndustry || 'Not Specified').trim(),
      serviceNeeded: (data.serviceNeeded || 'General Consultation').trim(),
      challenges: (data.challenges || 'N/A').trim(),
      budget: (data.budget || 'Flexible / Let\'s Discuss').trim(),
      message: (data.message || 'No additional project details provided.').trim(),
      source: (data.source || 'Website Contact Form').trim(),
      timestamp: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
    };

    // 3. Save lead data into Google Sheets database
    const sheetResult = appendLeadToSheet(sanitizedData);

    // 4. Send Email Notification to Mark Dalumpines (Receiver)
    let adminEmailSent = false;
    try {
      sendAdminNotification(sanitizedData);
      adminEmailSent = true;
    } catch (adminErr) {
      Logger.log('[Warning] Error dispatching Admin email: ' + adminErr.toString());
    }

    // 5. Send Auto-Responder Email to the Lead / Client
    let leadEmailSent = false;
    if (sanitizedData.email && sanitizedData.email.indexOf('@') !== -1) {
      try {
        sendLeadConfirmation(sanitizedData);
        leadEmailSent = true;
      } catch (leadErr) {
        Logger.log('[Warning] Error dispatching Lead email: ' + leadErr.toString());
      }
    }

    // 6. Return standard JSON response
    const responsePayload = {
      status: 'success',
      message: 'Inquiry successfully saved to database and automated notifications dispatched.',
      sheetUpdated: sheetResult,
      adminNotified: adminEmailSent,
      leadNotified: leadEmailSent,
      timestamp: sanitizedData.timestamp
    };

    return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('[Critical Error] doPost failure: ' + err.toString());
    const errorPayload = {
      status: 'error',
      message: err.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle HTTP GET requests for deployment health verification
 */
function doGet(e) {
  const info = {
    status: 'active',
    system: 'Mark Dalumpines Portfolio Lead Automation Engine',
    specialist: 'Mark Dalumpines — AI Automation Specialist & GHL Expert',
    adminEmail: CONFIG.ADMIN_EMAIL,
    timezone: CONFIG.TIMEZONE,
    timestamp: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
  };

  return ContentService.createTextOutput(JSON.stringify(info))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Appends the sanitized lead data to the target Google Sheet
 */
function appendLeadToSheet(data) {
  let ss;
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== '') {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('Unable to connect to Google Spreadsheet. Please verify SPREADSHEET_ID or bind script to sheet.');
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // If sheet is completely new or has no rows, create the formatted header
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Timestamp (GMT+8)',
      'Client Name',
      'Email Address',
      'Phone Number',
      'Business Name',
      'Business Industry',
      'Service Needed',
      'Budget Range',
      'Business Challenges',
      'Project Details / Scope',
      'Lead Status',
      'Traffic Source'
    ];

    sheet.appendRow(headers);

    // Style Header Row (Dark Blue background, bold white text, frozen row)
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0b1e33');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontFamily('Arial');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  const rowData = [
    data.timestamp,
    data.name,
    data.email,
    data.phone,
    data.businessName,
    data.businessIndustry,
    data.serviceNeeded,
    data.budget,
    data.challenges,
    data.message,
    'New Lead',
    data.source
  ];

  sheet.appendRow(rowData);

  // Format the newly added row
  const lastRow = sheet.getLastRow();
  const newRowRange = sheet.getRange(lastRow, 1, 1, rowData.length);
  newRowRange.setVerticalAlignment('middle');
  newRowRange.setFontFamily('Arial');
  newRowRange.setFontSize(10);

  // Auto-resize columns on periodic basis
  try {
    for (let c = 1; c <= rowData.length; c++) {
      if (c !== 9 && c !== 10) { // Keep challenges and message wrapped
        sheet.autoResizeColumn(c);
      }
    }
    sheet.getRange(lastRow, 9, 1, 2).setWrap(true);
  } catch (resizeErr) {}

  return true;
}

/**
 * Dispatch Admin Notification to Mark Dalumpines using ReceivedEmailMark.html
 */
function sendAdminNotification(data) {
  const template = HtmlService.createTemplateFromFile('ReceivedEmailMark');

  // Bind template variables
  template.name = data.name;
  template.email = data.email;
  template.phone = data.phone;
  template.businessName = data.businessName;
  template.businessIndustry = data.businessIndustry;
  template.serviceNeeded = data.serviceNeeded;
  template.budget = data.budget;
  template.challenges = data.challenges;
  template.message = data.message;
  template.timestamp = data.timestamp;
  template.source = data.source;
  template.markName = CONFIG.ADMIN_NAME;
  template.markRole = CONFIG.ADMIN_TITLE;

  const htmlBody = template.evaluate().getContent();
  const subject = `⚡ New Project Inquiry: ${data.businessName} — ${data.serviceNeeded} (${data.name})`;

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, '', {
    htmlBody: htmlBody,
    name: 'Portfolio Inbound System',
    replyTo: data.email
  });
}

/**
 * Dispatch Auto-Responder Confirmation to Lead using SendEmailLead.html
 */
function sendLeadConfirmation(data) {
  const template = HtmlService.createTemplateFromFile('SendEmailLead');

  // Bind template variables
  template.name = data.name;
  template.email = data.email;
  template.phone = data.phone;
  template.businessName = data.businessName;
  template.businessIndustry = data.businessIndustry;
  template.serviceNeeded = data.serviceNeeded;
  template.budget = data.budget;
  template.challenges = data.challenges;
  template.message = data.message;
  template.timestamp = data.timestamp;
  template.markName = CONFIG.ADMIN_NAME;
  template.markRole = CONFIG.ADMIN_TITLE;
  template.adminEmail = CONFIG.ADMIN_EMAIL;

  const htmlBody = template.evaluate().getContent();
  const subject = `Thank you for reaching out, ${data.name}! I've received your inquiry`;

  GmailApp.sendEmail(data.email, subject, '', {
    htmlBody: htmlBody,
    name: CONFIG.SENDER_DISPLAY_NAME,
    replyTo: CONFIG.ADMIN_EMAIL
  });
}

/**
 * Test function: run directly inside the Google Apps Script Editor to verify setup
 */
function testSubmission() {
  const sampleEvent = {
    postData: {
      contents: JSON.stringify({
        name: 'Alex Mercer',
        email: CONFIG.ADMIN_EMAIL,
        phone: '+63 917 123 4567',
        businessName: 'Apex Horizons Consulting',
        businessIndustry: 'E-Commerce & Digital Marketing',
        serviceNeeded: 'Funnels-CRM-Automation-Go High Level Platform',
        challenges: 'Manual lead follow-up taking 4 hours daily, lack of automated pipeline nurturing.',
        budget: '₱30,000 - ₱60,000 (~$600 - $1,200)',
        message: 'We need an automated GHL funnel that captures leads from Meta Ads, sends SMS and email follow-ups, and integrates with our Google Sheet database.',
        source: 'Test Runner in Apps Script'
      })
    }
  };

  const result = doPost(sampleEvent);
  Logger.log('Test result: ' + result.getContent());
}

