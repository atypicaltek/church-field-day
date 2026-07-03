/**
 * FIELD DAY SIGNUP — CONFIRMATION / UPDATE / CANCEL SYSTEM
 * =========================================================
 * Attach this script to the MASTER SPREADSHEET (Extensions > Apps Script).
 *
 * WHAT IT DOES
 * 1. When someone submits (or edits) any of the 11 signup forms, it emails
 *    them a receipt listing every option they selected, plus a personal link
 *    to UPDATE or CANCEL their signup at any time.
 * 2. When they use that link to change answers and resubmit, this script
 *    fires again, updates their row in the Sheet (marks old row Superseded,
 *    new row Active), and sends a fresh "Here's your CURRENT signup" email.
 * 3. When they check "Cancel my signup" and resubmit, it marks their row
 *    Cancelled, excludes them from live counts, and confirms by email.
 * 4. After every submit/update/cancel, it rebuilds the Dashboard tab showing
 *    live Active-only counts by Site × Activity.
 *
 * SETUP STEPS
 * 1. Run `createAllFieldDayForms()` in FieldDay_CreateForms.gs first.
 *    That script creates all 11 forms with Email and Cancel fields already
 *    added, and prints Form IDs in the "Form Links & Embed Codes" tab of
 *    the master spreadsheet (column B).
 * 2. Copy each Form ID from column B of that tab into the FORM_IDS object
 *    below, replacing the "PASTE_…" placeholders.
 * 3. Run `setupTriggers()` ONCE from this script (select it in the function
 *    dropdown → Run → authorize when prompted).
 *    That's it — the system is live.
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────

// Paste each form's ID here (from "Form Links & Embed Codes" tab, column B).
const FORM_IDS = {
  "Food":             "PASTE_FOOD_FORM_ID_HERE",
  "Drinks":           "PASTE_DRINKS_FORM_ID_HERE",
  "Games":            "PASTE_GAMES_FORM_ID_HERE",
  "DJ":               "PASTE_DJ_FORM_ID_HERE",
  "Karaoke":          "PASTE_KARAOKE_FORM_ID_HERE",
  "Basketball":       "PASTE_BASKETBALL_FORM_ID_HERE",
  "Kid Games":        "PASTE_KIDGAMES_FORM_ID_HERE",
  "Parking":          "PASTE_PARKING_FORM_ID_HERE",
  "Site Setup":       "PASTE_SITESETUP_FORM_ID_HERE",
  "Popcorn Machine":  "PASTE_POPCORN_FORM_ID_HERE",
  "Pretzels":         "PASTE_PRETZELS_FORM_ID_HERE",
};

// Maps each FORM_IDS key → the exact Google Sheets tab name created by FieldDay_CreateForms.gs
const SHEET_TAB_NAMES = {
  "Food":             "Food Signups",
  "Drinks":           "Drinks Signups",
  "Games":            "Games Signups",
  "DJ":               "DJ Signups",
  "Karaoke":          "Karaoke Signups",
  "Basketball":       "Basketball Signups",
  "Kid Games":        "Kid Games Signups",
  "Parking":          "Parking Signups",
  "Site Setup":       "Site Setup Signups",
  "Popcorn Machine":  "Popcorn Signups",
  "Pretzels":         "Pretzels Signups",
};

// Must match addEmailAndCancel() in FieldDay_CreateForms.gs exactly
const CANCEL_QUESTION_TITLE = "Cancel my signup (check to withdraw)";
const EMAIL_QUESTION_TITLE  = "Email Address";

const DASHBOARD_SHEET_NAME = "Dashboard";

const SITE_LIST = [
  "The Cathedral",
  "New Beginnings",
  "Newport News",
  "Downtown Hampton",
  "Kecoughtan Road",
  "Ingleside Road",
  "Port Norfolk",
  "Downtown Franklin",
];


// ── ONE-TIME SETUP ────────────────────────────────────────────────────────────

/**
 * Run this ONCE after filling in FORM_IDS above.
 * Creates one onFormSubmit trigger per form.
 */
function setupTriggers() {
  // Remove any existing triggers this script created (safe to re-run)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onFormSubmitHandler") {
      ScriptApp.deleteTrigger(t);
    }
  });

  let created = 0;
  Object.entries(FORM_IDS).forEach(([activity, formId]) => {
    if (formId.startsWith("PASTE_")) {
      Logger.log("⚠️  Skipping " + activity + " — form ID not filled in yet.");
      return;
    }
    const form = FormApp.openById(formId);
    ScriptApp.newTrigger("onFormSubmitHandler")
      .forForm(form)
      .onFormSubmit()
      .create();
    Logger.log("✅ Trigger created for: " + activity);
    created++;
  });

  Logger.log("\nSetup complete. Triggers created: " + created + " / " + Object.keys(FORM_IDS).length);
}


// ── MAIN HANDLER ─────────────────────────────────────────────────────────────

/**
 * Fires on every new submit AND every edit-link resubmit.
 * Google passes the form and response via the event object `e`.
 */
function onFormSubmitHandler(e) {
  const formResponse = e.response;
  const form         = e.source;
  const activityKey  = getActivityKey(form);
  const itemResponses = formResponse.getItemResponses();

  let email      = null;
  let isCancel   = false;
  const answerLines = [];

  itemResponses.forEach(ir => {
    const title  = ir.getItem().getTitle();
    const answer = ir.getResponse();

    if (title === EMAIL_QUESTION_TITLE) {
      email = answer;
      return; // don't include email in the receipt body
    }

    if (title === CANCEL_QUESTION_TITLE) {
      isCancel = Array.isArray(answer) && answer.length > 0;
      return; // don't include cancel flag in receipt body
    }

    answerLines.push("  • " + title + ": " + (Array.isArray(answer) ? answer.join(", ") : answer));
  });

  const editUrl     = formResponse.getEditResponseUrl();
  const summaryText = answerLines.join("\n");

  // Update the spreadsheet row's Status column
  updateRowStatus(form, formResponse, isCancel ? "CANCELLED" : "Active");

  // Send confirmation email
  if (email) {
    sendConfirmationEmail(email, activityKey, summaryText, editUrl, isCancel);
  }

  // Rebuild the live Dashboard tab
  rebuildDashboard();
}


// ── EMAIL ─────────────────────────────────────────────────────────────────────

function sendConfirmationEmail(toEmail, activityKey, summaryText, editUrl, isCancel) {
  const activityName = activityKey || "Field Day Activity";

  const subject = isCancel
    ? "Field Day — " + activityName + " signup CANCELLED"
    : "Field Day — " + activityName + " signup confirmed ✅";

  const body = isCancel
    ? "Your signup for " + activityName + " has been cancelled.\n\n" +
      "Changed your mind? Use this link to sign up again:\n" + editUrl + "\n\n" +
      "— Field Day Team"
    : "Thanks for signing up! Here is your " + activityName + " receipt:\n\n" +
      summaryText + "\n\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "Need to update or cancel? Use this personal link — it will always\n" +
      "show your current information pre-filled and ready to edit:\n\n" +
      editUrl + "\n\n" +
      "— Field Day Team";

  MailApp.sendEmail(toEmail, subject, body);
}


// ── SPREADSHEET HELPERS ───────────────────────────────────────────────────────

/**
 * Find the row for this response in its linked sheet and set the Status column.
 * Uses response Timestamp to match the row (accurate to the second).
 */
function updateRowStatus(form, formResponse, status) {
  const activityKey = getActivityKey(form);
  const tabName     = SHEET_TAB_NAMES[activityKey];
  if (!tabName) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const timestamp = formResponse.getTimestamp();
  const data      = sheet.getDataRange().getValues();
  const headers   = data[0];

  // Ensure a Status and Last Updated column exist
  const statusCol     = getOrCreateColumn(sheet, headers, "Status");
  const updatedCol    = getOrCreateColumn(sheet, headers, "Last Updated");

  // Match by Timestamp (column A = index 0)
  for (let row = 1; row < data.length; row++) {
    if (Math.abs(new Date(data[row][0]).getTime() - new Date(timestamp).getTime()) < 2000) {
      sheet.getRange(row + 1, statusCol).setValue(status);
      sheet.getRange(row + 1, updatedCol).setValue(new Date());

      // Color cancelled rows grey for quick visual scan
      if (status === "CANCELLED") {
        sheet.getRange(row + 1, 1, 1, sheet.getLastColumn())
          .setFontColor("#999999")
          .setBackground("#F0F0F0");
      }
      break;
    }
  }
}

/**
 * Return the column number for `header`, creating it if it doesn't exist.
 * Re-reads the header row each time to handle multiple additions in one call.
 */
function getOrCreateColumn(sheet, headers, header) {
  let col = headers.indexOf(header) + 1;
  if (col === 0) {
    col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue(header).setFontWeight("bold");
    headers.push(header); // keep local array in sync for the second call
  }
  return col;
}

/**
 * Return the FORM_IDS key that matches this form's ID.
 */
function getActivityKey(form) {
  const id = form.getId();
  return Object.keys(FORM_IDS).find(k => FORM_IDS[k] === id) || form.getTitle();
}


// ── LIVE DASHBOARD ────────────────────────────────────────────────────────────

/**
 * Rebuilds the Dashboard tab: rows = sites, columns = activities.
 * Only counts rows where Status = "Active" (or Status is blank, for legacy rows).
 */
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_SHEET_NAME);
  } else {
    dash.clear();
  }

  const activities = Object.keys(FORM_IDS);

  // Header row
  const headerRow = ["Site", ...activities];
  dash.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
    .setFontWeight("bold")
    .setBackground("#1F3864")
    .setFontColor("#FFFFFF");

  // Data rows
  SITE_LIST.forEach((site, rowIdx) => {
    const rowData = [site];
    activities.forEach(act => {
      const tabName = SHEET_TAB_NAMES[act];
      let count = 0;
      if (tabName) {
        const sheet = ss.getSheetByName(tabName);
        if (sheet && sheet.getLastRow() > 1) {
          const data      = sheet.getDataRange().getValues();
          const colHeaders = data[0];
          const siteCol   = colHeaders.indexOf("Which site are you from?");
          const statusCol = colHeaders.indexOf("Status");

          for (let r = 1; r < data.length; r++) {
            const rowSite   = data[r][siteCol];
            const rowStatus = statusCol >= 0 ? String(data[r][statusCol]).trim() : "";
            // Count Active rows (and blank-status legacy rows)
            if (rowSite === site && rowStatus !== "CANCELLED" && rowStatus !== "SUPERSEDED") {
              count++;
            }
          }
        }
      }
      rowData.push(count);
    });
    dash.getRange(rowIdx + 2, 1, 1, rowData.length).setValues([rowData]);
  });

  // Auto-resize columns for readability
  dash.autoResizeColumns(1, activities.length + 1);
  dash.setFrozenRows(1);
  dash.setFrozenColumns(1);

  Logger.log("Dashboard rebuilt: " + new Date());
}
