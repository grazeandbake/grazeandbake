/**
 * GRAZE & BAKE — Website Enquiry & Job Sheet Receiver
 * ----------------------------------------------------
 * Handles TWO forms from the website:
 *  1. The public enquiry form  → logs to "Enquiries" tab + emails you
 *  2. The private job-details.html page (shared with confirmed clients
 *     only) → logs to "Job Sheets" tab + emails you a print-ready summary
 *
 * SETUP:
 * 1. Go to script.google.com → New project (standalone script)
 * 2. Paste this entire file in, replacing the placeholder code
 * 3. Open your Booking Tracker sheet, copy the ID from its URL:
 *    docs.google.com/spreadsheets/d/PASTE-THIS-BIT-HERE/edit
 * 4. Paste it into SHEET_ID below
 * 5. Deploy → New deployment → Web app → Execute as: Me →
 *    Who has access: Anyone → Deploy → authorise → copy the URL
 * 6. That same URL goes into BOTH index.html (SHEETS_WEBAPP_URL) and
 *    job-details.html (SHEETS_WEBAPP_URL) — one script, one URL, two forms.
 *
 * Emails go to grazeandbake@hotmail.co.uk, cc joshscottlaurentyson@hotmail.com.
 */

var SHEET_ID = ""; // paste your spreadsheet ID between the quotes

function getSpreadsheet_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getEnquiriesSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName("Enquiries");
  if (!sheet) {
    sheet = ss.insertSheet("Enquiries");
    var headers = [
      "Received", "Name", "Phone", "Email", "Event Date", "Event Type",
      "Guests", "Services Interested In", "Allergies / Dietary", "Message", "Status"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold").setBackground("#5C3D1E").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 140);
    sheet.setColumnWidth(8, 260);
    sheet.setColumnWidth(10, 300);
  }
  return sheet;
}

function getJobSheetsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName("Job Sheets");
  if (!sheet) {
    sheet = ss.insertSheet("Job Sheets");
    var headers = [
      "Received", "Name", "Phone", "Event Date", "Event Type",
      "Adults", "Children", "Address", "Delivery/Collection", "Setup Time",
      "Event Start", "Collection Time", "Access Notes", "Items Booked",
      "Theme/Styling", "Allergies", "Venue Contact", "Notes"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold").setBackground("#6E4227").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
    sheet.setColumnWidth(8, 260);
    sheet.setColumnWidth(13, 260);
    sheet.setColumnWidth(14, 260);
    sheet.setColumnWidth(18, 260);
  }
  return sheet;
}

function doPost(e) {
  try {
    var p = e.parameter;
    if (p.formType === "jobsheet") {
      return handleJobSheet_(p);
    }
    return handleEnquiry_(p);
  } catch (err) {
    return ContentService.createTextOutput("ERROR: " + err);
  }
}

function handleEnquiry_(p) {
  var sheet = getEnquiriesSheet_();
  sheet.appendRow([
    new Date(), p.name || "", p.phone || "", p.email || "", p.event_date || "",
    p.event_type || "", p.guests || "", p.services || "", p.allergies || "",
    p.message || "", "NEW"
  ]);

  MailApp.sendEmail({
    to: "grazeandbake@hotmail.co.uk",
    cc: "joshscottlaurentyson@hotmail.com",
    subject: "New website enquiry — " + (p.name || "customer") + (p.event_date ? " (" + p.event_date + ")" : ""),
    body:
      "New enquiry from the website:\n\n" +
      "Name: " + (p.name || "") + "\n" +
      "Phone: " + (p.phone || "") + "\n" +
      "Email: " + (p.email || "") + "\n" +
      "Event date: " + (p.event_date || "") + "\n" +
      "Event type: " + (p.event_type || "") + "\n" +
      "Guests: " + (p.guests || "") + "\n" +
      "Services interested in: " + (p.services || "") + "\n" +
      "Allergies / dietary: " + (p.allergies || "") + "\n" +
      "Message: " + (p.message || "") + "\n\n" +
      "This enquiry has also been added to the Enquiries tab of your Booking Tracker."
  });

  return ContentService.createTextOutput("OK");
}

function handleJobSheet_(p) {
  var sheet = getJobSheetsSheet_();
  sheet.appendRow([
    new Date(), p.name || "", p.phone || "", p.event_date || "", p.event_type || "",
    p.guests_adults || "", p.guests_children || "", p.address || "",
    p.delivery || "", p.setup_time || "", p.event_start || "", p.collection_time || "",
    p.access || "", p.items || "", p.theme || "", p.allergies || "",
    p.venue_contact || "", p.notes || ""
  ]);

  MailApp.sendEmail({
    to: "grazeandbake@hotmail.co.uk",
    cc: "joshscottlaurentyson@hotmail.com",
    subject: "JOB SHEET — " + (p.name || "client") + (p.event_date ? " (" + p.event_date + ")" : ""),
    body:
      "═══════════════════════════════\n" +
      "JOB SHEET — READY TO PRINT\n" +
      "═══════════════════════════════\n\n" +
      "CLIENT: " + (p.name || "") + "\n" +
      "CONTACT: " + (p.phone || "") + "\n" +
      "EVENT DATE: " + (p.event_date || "") + "\n" +
      "EVENT TYPE: " + (p.event_type || "") + "\n" +
      "GUESTS: " + (p.guests_adults || "0") + " adults, " + (p.guests_children || "0") + " children\n\n" +
      "── VENUE & TIMINGS ──\n" +
      "Address: " + (p.address || "") + "\n" +
      "Delivery or collection: " + (p.delivery || "") + "\n" +
      "Setup needed by: " + (p.setup_time || "") + "\n" +
      "Event start: " + (p.event_start || "") + "\n" +
      "Collection time: " + (p.collection_time || "") + "\n" +
      "Parking & access notes: " + (p.access || "") + "\n\n" +
      "── ORDER ──\n" +
      "Items booked: " + (p.items || "") + "\n" +
      "Theme / styling: " + (p.theme || "") + "\n\n" +
      "── IMPORTANT ──\n" +
      "Allergies / dietary requirements: " + (p.allergies || "") + "\n" +
      "Venue contact: " + (p.venue_contact || "") + "\n" +
      "Additional notes: " + (p.notes || "") + "\n\n" +
      "═══════════════════════════════\n" +
      "Also saved to the Job Sheets tab of your Booking Tracker."
  });

  return ContentService.createTextOutput("OK");
}

// Run manually once (select from the dropdown next to Run, click Run) to
// authorise and confirm both tabs are reachable before going live.
function testSetup() {
  getEnquiriesSheet_();
  getJobSheetsSheet_();
  Logger.log("Both sheets connected OK");
}
