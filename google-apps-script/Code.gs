/**
 * Google Apps Script for Skillversity IT Support & Employee Complaint Ticketing System
 */

const TARGET_EMAIL = "skillversitycomplaints@gmail.com"; // Default notification recipient

const HEADERS = [
  "Ticket ID",
  "Date & Time",
  "Employee Name",
  "Employee Email",
  "Contact Number",
  "Department",
  "Category",
  "Subcategory",
  "Subject",
  "Description",
  "Impact",
  "Urgency",
  "Priority",
  "Status",
  "Assigned To",
  "Vendor Ticket Ref",
  "Troubleshooting Notes",
  "Last Updated",
  "Uploaded Files / Attachments"
];

/**
 * Run this function directly from the Apps Script Editor menu (Select 'setupSheetHeaders' -> Click 'Run')
 * to immediately create or update Row 1 column headers in your existing Google Sheet.
 */
function setupSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#2563eb")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  Logger.log("Headers updated successfully to 19 standard columns!");
  return "Headers updated successfully to 19 standard columns!";
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "Skillversity IT Ticketing Apps Script Endpoint is Active",
    headersCount: HEADERS.length,
    headers: HEADERS
  })).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheetHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    setupSheetHeaders();
    return;
  }
  
  // Check if current row 1 matches expected headers
  const existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  let needsUpdate = false;
  
  if (!existingHeaders || existingHeaders.length < HEADERS.length) {
    needsUpdate = true;
  } else {
    // Check if Contact Number (index 4) or standard headers are missing
    if (existingHeaders[4] !== "Contact Number" || existingHeaders[0] !== "Ticket ID") {
      needsUpdate = true;
    }
  }

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#2563eb")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Ensure column headers are up to date (19 columns)
    ensureSheetHeaders(sheet);

    const action = data.action || "CREATE";

    if (action === "SETUP_HEADERS") {
      setupSheetHeaders();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "Headers updated successfully in Google Sheet"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "CREATE") {
      const row = [
        data.ticketId || "",
        data.date || new Date().toLocaleString(),
        data.userName || "",
        data.userEmail || "",
        data.contactNumber || data.userPhone || "N/A",
        data.department || "",
        data.category || "",
        data.subCategory || "",
        data.subject || "",
        data.description || "",
        data.impact || "Medium",
        data.urgency || "Medium",
        data.priority || "Medium",
        data.status || "Open",
        data.assignedTo || "Unassigned",
        data.vendorTicketRef || "N/A",
        data.troubleshootingNotes || "",
        new Date().toLocaleString(),
        data.attachmentUrls || "None"
      ];

      sheet.appendRow(row);

      // Send parallel email notification
      sendNotificationEmail(data);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "Ticket created and synced to Skillversity Google Sheet + Email sent successfully",
        ticketId: data.ticketId
      })).setMimeType(ContentService.MimeType.JSON);
    } 
    
    else if (action === "UPDATE") {
      const ticketId = data.ticketId;
      const values = sheet.getDataRange().getValues();
      let rowIndex = -1;

      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === ticketId) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        if (data.status) sheet.getRange(rowIndex, 14).setValue(data.status);
        if (data.assignedTo) sheet.getRange(rowIndex, 15).setValue(data.assignedTo);
        if (data.vendorTicketRef !== undefined) sheet.getRange(rowIndex, 16).setValue(data.vendorTicketRef);
        if (data.troubleshootingNotes) sheet.getRange(rowIndex, 17).setValue(data.troubleshootingNotes);
        sheet.getRange(rowIndex, 18).setValue(new Date().toLocaleString());

        if (data.status === "Resolved" && data.userEmail) {
          sendResolutionEmail(data);
        }

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          message: "Ticket updated in Google Sheet",
          ticketId: ticketId
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Ticket ID not found in sheet"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    else if (action === "DELETE") {
      const ticketId = data.ticketId;
      const values = sheet.getDataRange().getValues();
      let rowIndex = -1;

      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === ticketId) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        sheet.deleteRow(rowIndex);
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          message: "Ticket row deleted from Google Sheet",
          ticketId: ticketId
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Ticket ID not found in sheet"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotificationEmail(data) {
  const recipient = data.targetEmail || TARGET_EMAIL;
  const subject = `[Skillversity Ticket #${data.ticketId}] Priority: ${data.priority} - ${data.subject}`;
  const body = `
New Skillversity Support Ticket / Employee Complaint Registered:

Ticket ID: ${data.ticketId}
Date: ${data.date}
Employee Name: ${data.userName}
Employee Email: ${data.userEmail}
Contact Number: ${data.contactNumber || data.userPhone || "N/A"}
Department: ${data.department}
Category: ${data.category} (Subcategory: ${data.subCategory || "N/A"})
Priority: ${data.priority} (Impact: ${data.impact}, Urgency: ${data.urgency})

Subject: ${data.subject}

Description:
${data.description}

Attached Files:
${data.attachmentUrls || "None"}

------------------------------------------------
Skillversity IT Support & Complaint Management System
`;

  try {
    MailApp.sendEmail(recipient, subject, body);
    
    if (data.userEmail) {
      const userSubject = `Skillversity Ticket Received [#${data.ticketId}] - ${data.subject}`;
      const userBody = `Dear ${data.userName},\n\nYour support ticket #${data.ticketId} has been successfully logged with Skillversity IT Support.\n\nSubject: ${data.subject}\nCategory: ${data.category} (${data.subCategory || "General"})\nPriority: ${data.priority}\nStatus: Open\nAttachments: ${data.attachmentUrls || "None"}\n\nOur team will review your request shortly.\n\nThank you,\nSkillversity IT Support Team`;
      MailApp.sendEmail(data.userEmail, userSubject, userBody);
    }
  } catch (e) {
    Logger.log("Email Error: " + e.toString());
  }
}

function sendResolutionEmail(data) {
  const subject = `[Skillversity Ticket Resolved #${data.ticketId}] ${data.subject}`;
  const body = `Dear ${data.userName || "User"},\n\nYour support ticket #${data.ticketId} has been marked as RESOLVED by Skillversity IT Support.\n\nResolution Details & Actions Taken:\n${data.troubleshootingNotes || "Issue resolved by IT Support."}\n\nPlease confirm resolution on the portal if your issue is fixed.\n\nThank you,\nSkillversity IT Support Team`;
  try {
    MailApp.sendEmail(data.userEmail, subject, body);
  } catch (e) {
    Logger.log("Resolution Email Error: " + e.toString());
  }
}
