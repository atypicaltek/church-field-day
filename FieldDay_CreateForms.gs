/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  CHURCH FIELD DAY — Auto-Form Builder
 *  Google Apps Script  |  Run ONCE to create all 11 signup forms +
 *  one master Google Spreadsheet with a tab per form.
 *
 *  HOW TO USE:
 *  1. Go to script.google.com  →  New project
 *  2. Paste this entire file, replacing the default code
 *  3. Click  Run  ▶  →  select  createAllFieldDayForms
 *  4. Approve the permissions popup (Forms + Sheets access)
 *  5. When done, open the Execution log (View → Logs) to see every form URL
 *  6. Copy the URLs into your Google Sites pages (see Site guide)
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Config ──────────────────────────────────────────────────────────────────
const SPREADSHEET_NAME = "Field Day Signups — Master";
const FORM_FOLDER      = "Field Day Forms";   // Google Drive folder name

// Question title used to detect cancellation — must match EXACTLY in FieldDay_Confirmation_System.gs
const CANCEL_Q_TITLE = "Cancel my signup (check to withdraw)";

const SITES = [
  "The Cathedral",
  "New Beginnings",
  "Newport News",
  "Downtown Hampton",
  "Kecoughtan Road",
  "Ingleside Road",
  "Port Norfolk",
  "Downtown Franklin",
];

const AVAILABILITY = ["Full event", "Setup / serving only", "Cleanup only"];
const SHIFTS_DAY   = ["Morning", "Afternoon", "All day"];
const SHIFTS_PARK  = ["Arrival / early shift", "Mid-day", "Departure / late shift", "Full day"];


// ── Main entry point ─────────────────────────────────────────────────────────
function createAllFieldDayForms() {
  // 1. Create (or open) the master spreadsheet
  const ss  = getOrCreateSpreadsheet(SPREADSHEET_NAME);
  const ssId = ss.getId();

  // 2. Create (or open) the Drive folder
  const folder = getOrCreateFolder(FORM_FOLDER);

  // 3. Build all 11 forms
  const results = [];
  results.push(createFoodForm(ss, folder));
  results.push(createDrinksForm(ss, folder));
  results.push(createGamesForm(ss, folder));
  results.push(createDJForm(ss, folder));
  results.push(createKaraokeForm(ss, folder));
  results.push(createBasketballForm(ss, folder));
  results.push(createKidGamesForm(ss, folder));
  results.push(createParkingForm(ss, folder));
  results.push(createSetupForm(ss, folder));
  results.push(createPopcornForm(ss, folder));
  results.push(createPretzelsForm(ss, folder));

  // 4. Log a summary
  Logger.log("\n════════════ FIELD DAY FORM URLS ════════════");
  Logger.log("Master Spreadsheet: https://docs.google.com/spreadsheets/d/" + ssId);
  Logger.log("");
  results.forEach(r => {
    Logger.log("✅ " + r.name);
    Logger.log("   Respondent URL : " + r.url);
    Logger.log("   Embed URL      : " + r.embedUrl);
    Logger.log("");
  });
  Logger.log("Paste the Embed URLs into Google Sites using Insert → Embed.");
  Logger.log("════════════════════════════════════════════════");

  // 5. Write a summary tab in the spreadsheet
  writeSummaryTab(ss, results);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 1 — FOOD SIGNUP
// ══════════════════════════════════════════════════════════════════════════════
function createFoodForm(ss, folder) {
  const form = FormApp.create("Field Day — Food Signup");
  form.setDescription(
    "Each site is pre-assigned a food item by the Food Chair. " +
    "This form signs up servers who will prepare and serve at their station. " +
    "The Cathedral needs 12 servers; all other sites need 8 servers each."
  );
  form.setCollectEmail(false);
  form.setConfirmationMessage("Thank you for signing up to serve food! Your Site Lead will be in touch.");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Server Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  const foodItems = form.addListItem().setTitle("Assigned Food Item").setRequired(true);
  foodItems.setChoiceValues(SITES.map(s => s + " – [Item TBD — Food Chair will update]"));

  form.addMultipleChoiceItem().setTitle("Availability").setRequired(true)
    .setChoiceValues(AVAILABILITY);

  form.addParagraphTextItem().setTitle("Allergy / Dietary Notes (optional)").setRequired(false);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Food Signups");
  moveToFolder(form, folder);

  return formResult("Food Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 2 — DRINKS
// ══════════════════════════════════════════════════════════════════════════════
function createDrinksForm(ss, folder) {
  const form = FormApp.create("Field Day — Drinks Signup");
  form.setDescription("Sign up to restock coolers, run beverages to stations, or staff the drinks table.");
  form.setConfirmationMessage("Drinks signup received! Stay hydrated 💧");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addMultipleChoiceItem().setTitle("Role").setRequired(true)
    .setChoiceValues(["Restock / runner", "Table staff", "Setup / teardown"]);

  form.addTextItem().setTitle("Time Availability (optional)").setRequired(false);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Drinks Signups");
  moveToFolder(form, folder);
  return formResult("Drinks Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 3 — GAMES
// ══════════════════════════════════════════════════════════════════════════════
function createGamesForm(ss, folder) {
  const form = FormApp.create("Field Day — Games Signup");
  form.setDescription("Staff a game station and keep the competition fun all day!");
  form.setConfirmationMessage("Games signup received — see you on the field! 🎯");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addListItem().setTitle("Which Game Station?").setRequired(true)
    .setChoiceValues([
      "Station 1 – [TBD by Games Chair]",
      "Station 2 – [TBD by Games Chair]",
      "Station 3 – [TBD by Games Chair]",
      "Station 4 – [TBD by Games Chair]",
      "No preference",
    ]);

  form.addMultipleChoiceItem().setTitle("Available for Setup?").setRequired(true)
    .setChoiceValues(["Yes", "No"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Games Signups");
  moveToFolder(form, folder);
  return formResult("Games Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 4 — DJ / MUSIC
// ══════════════════════════════════════════════════════════════════════════════
function createDJForm(ss, folder) {
  const form = FormApp.create("Field Day — DJ / Music Requests");
  form.setDescription("Submit song requests, announcements, or volunteer as an MC helper for your site.");
  form.setConfirmationMessage("DJ request submitted! 🎵 Keep an ear out for your song.");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Your Name").setRequired(true);
  form.addTextItem().setTitle("Song Request (optional)").setRequired(false);
  form.addParagraphTextItem().setTitle("Shoutout / Announcement Request (optional)").setRequired(false);

  form.addMultipleChoiceItem().setTitle("Volunteer as MC Helper?").setRequired(true)
    .setChoiceValues([
      "Yes — I'll help with site shoutouts and announcements",
      "No — just submitting a request",
    ]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "DJ Signups");
  moveToFolder(form, folder);
  return formResult("DJ / Music", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 5 — KARAOKE
// ══════════════════════════════════════════════════════════════════════════════
function createKaraokeForm(ss, folder) {
  const form = FormApp.create("Field Day — Karaoke Signup");
  form.setDescription("Ready to take the stage? Solo acts, duets, and groups all welcome. Each site needs at least 2 singers!");
  form.setConfirmationMessage("You're on the list! See you on stage 🎤");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Singer Name(s)").setRequired(true)
    .setHelpText("For duets or groups, list all names");
  form.addTextItem().setTitle("Phone Number").setRequired(true);
  form.addTextItem().setTitle("Song Choice").setRequired(true);
  form.addTextItem().setTitle("Backup Song (optional)").setRequired(false)
    .setHelpText("In case your first choice is taken");

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Karaoke Signups");
  moveToFolder(form, folder);
  return formResult("Karaoke Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 6 — BASKETBALL TOURNAMENT
// ══════════════════════════════════════════════════════════════════════════════
function createBasketballForm(ss, folder) {
  const form = FormApp.create("Field Day — Basketball Tournament Registration");
  form.setDescription("Register your site's team (5–8 players + 1 captain). One team per site.");
  form.setConfirmationMessage("Team registered! Bracket info coming soon 🏀");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Team Name").setRequired(true);
  form.addTextItem().setTitle("Team Captain Name").setRequired(true);
  form.addTextItem().setTitle("Captain Phone Number").setRequired(true);
  form.addParagraphTextItem().setTitle("Player Names (list all 5–8 players)").setRequired(true)
    .setHelpText("One player per line");
  form.addTextItem().setTitle("Jersey Color Preference (optional)").setRequired(false)
    .setHelpText("Helps avoid same-color matchups");

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Basketball Signups");
  moveToFolder(form, folder);
  return formResult("Basketball Tournament", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 7 — KID GAMES
// ══════════════════════════════════════════════════════════════════════════════
function createKidGamesForm(ss, folder) {
  const form = FormApp.create("Field Day — Kid Games Signup");
  form.setDescription("Help keep kids safe and entertained in the children's area!");
  form.setConfirmationMessage("Thank you for signing up to help with the kids! 🧒");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Chaperone / Helper Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addMultipleChoiceItem().setTitle("Age Group Comfortable Supervising").setRequired(true)
    .setChoiceValues(["Ages 3–6", "Ages 7–10", "Ages 11+", "Any age group"]);

  form.addMultipleChoiceItem().setTitle("First Aid / CPR Certified?").setRequired(true)
    .setChoiceValues(["Yes", "No"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Kid Games Signups");
  moveToFolder(form, folder);
  return formResult("Kid Games Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 8 — PARKING
// ══════════════════════════════════════════════════════════════════════════════
function createParkingForm(ss, folder) {
  const form = FormApp.create("Field Day — Parking Volunteer Signup");
  form.setDescription(
    "Keep traffic flowing smoothly by directing parking on event day. " +
    "The Cathedral needs 4 attendants, Kecoughtan Road needs 2, and all other sites need 1 each."
  );
  form.setConfirmationMessage("Thanks for keeping us organized in the lot! 🅿");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addMultipleChoiceItem().setTitle("Shift Preference").setRequired(true)
    .setChoiceValues(SHIFTS_PARK);

  form.addMultipleChoiceItem().setTitle("Own Reflective Vest or Flashlight?").setRequired(true)
    .setChoiceValues(["Yes — I have a vest / flashlight", "No — will need one provided"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Parking Signups");
  moveToFolder(form, folder);
  return formResult("Parking Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 9 — SITE SETUP
// ══════════════════════════════════════════════════════════════════════════════
function createSetupForm(ss, folder) {
  const form = FormApp.create("Field Day — Site Setup Signup");
  form.setDescription("Help set up and break down the event — tables, chairs, tents, and general venue prep.");
  form.setConfirmationMessage("Setup crew confirmed! We'll be in touch with a call time ⛺");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  const equip = form.addCheckboxItem().setTitle("Equipment Your Site Can Bring").setRequired(true);
  equip.setChoiceValues(["Tables", "Chairs", "Tent(s)", "None — volunteer labor only"]);

  form.addTextItem().setTitle("Quantity of Equipment (optional)").setRequired(false)
    .setHelpText("e.g., 2 tables, 10 chairs");

  form.addMultipleChoiceItem().setTitle("Available For").setRequired(true)
    .setChoiceValues(["Setup only", "Teardown only", "Both setup and teardown"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Site Setup Signups");
  moveToFolder(form, folder);
  return formResult("Site Setup Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 10 — POPCORN MACHINE
// ══════════════════════════════════════════════════════════════════════════════
function createPopcornForm(ss, folder) {
  const form = FormApp.create("Field Day — Popcorn Machine Signup");
  form.setDescription("Staff the popcorn station and keep the snacks flowing all day!");
  form.setConfirmationMessage("You're popping in! Signup recorded 🍿");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addMultipleChoiceItem().setTitle("Shift Preference").setRequired(true)
    .setChoiceValues(SHIFTS_DAY);

  form.addMultipleChoiceItem().setTitle("Prior Experience Running a Popcorn Machine?").setRequired(true)
    .setChoiceValues(["Yes", "No"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Popcorn Signups");
  moveToFolder(form, folder);
  return formResult("Popcorn Machine Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  FORM 11 — PRETZELS
// ══════════════════════════════════════════════════════════════════════════════
function createPretzelsForm(ss, folder) {
  const form = FormApp.create("Field Day — Pretzels Signup");
  form.setDescription("Staff the pretzel station — warm, salty, and crowd-pleasing!");
  form.setConfirmationMessage("Twisted and registered! Signup recorded 🥨");

  addSiteDropdown(form);
  form.addTextItem().setTitle("Volunteer Name").setRequired(true);
  form.addTextItem().setTitle("Phone Number").setRequired(true);

  form.addMultipleChoiceItem().setTitle("Shift Preference").setRequired(true)
    .setChoiceValues(SHIFTS_DAY);

  form.addMultipleChoiceItem().setTitle("Prior Experience Running a Food Station?").setRequired(true)
    .setChoiceValues(["Yes", "No"]);

  addEmailAndCancel(form);
  linkToSheet(form, ss, "Pretzels Signups");
  moveToFolder(form, folder);
  return formResult("Pretzels Signup", form);
}


// ══════════════════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Add Email Address field + Cancel checkbox to every form.
 * Call this AFTER all activity-specific questions, BEFORE linkToSheet().
 * Also enables "Allow response editing" so the edit-link flow works.
 */
function addEmailAndCancel(form) {
  form.setAllowResponseEdits(true);

  form.addTextItem()
    .setTitle("Email Address")
    .setHelpText("We'll send your signup receipt here and include a link to update or cancel.")
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle(CANCEL_Q_TITLE)
    .setChoiceValues(["Yes — remove me from this activity"])
    .setRequired(false)
    .setHelpText("Leave unchecked when signing up. Check this only if you need to cancel a previous signup.");
}

/** Add the standard Site dropdown (required, same order on every form) */
function addSiteDropdown(form) {
  form.addListItem()
    .setTitle("Which site are you from?")
    .setRequired(true)
    .setChoiceValues(SITES);
}

/** Link a form's responses to a new tab in the master spreadsheet */
function linkToSheet(form, ss, tabName) {
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  // Google creates "Form Responses 1" — rename it
  Utilities.sleep(1200); // let the tab appear
  const sheets = ss.getSheets();
  const newSheet = sheets.find(s => s.getName().startsWith("Form Responses"));
  if (newSheet) newSheet.setName(tabName);
}

/** Move a form to the Drive folder so everything is organized */
function moveToFolder(form, folder) {
  const file = DriveApp.getFileById(form.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

/** Return a standard result object */
function formResult(name, form) {
  const id = form.getId();
  return {
    name:     name,
    url:      form.getPublishedUrl(),
    embedUrl: "https://docs.google.com/forms/d/" + id + "/viewform?embedded=true",
    editUrl:  form.getEditUrl(),
    id:       id,
  };
}

/** Create master spreadsheet or find existing one */
function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(name);
  // Rename the default sheet
  ss.getActiveSheet().setName("Summary");
  return ss;
}

/** Create Drive folder or find existing one */
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/** Write a summary tab with all form links */
function writeSummaryTab(ss, results) {
  let sheet = ss.getSheetByName("Form Links & Embed Codes");
  if (!sheet) sheet = ss.insertSheet("Form Links & Embed Codes", 0);
  sheet.clearContents();

  const headers = ["Activity", "Form ID (copy into Confirmation Script)", "Respondent URL (share with volunteers)", "Embed URL (paste into Google Sites)", "Edit URL (form owner only)"];
  sheet.appendRow(headers);
  sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1F3864").setFontColor("#FFFFFF");

  results.forEach((r, i) => {
    sheet.appendRow([r.name, r.id, r.url, r.embedUrl, r.editUrl]);
    // Alternate row color
    if (i % 2 === 0) sheet.getRange(i+2,1,1,5).setBackground("#D6E4F7");
  });

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 360);
  sheet.setColumnWidth(4, 380);
  sheet.setColumnWidth(5, 360);
  sheet.setFrozenRows(1);

  Logger.log("Summary tab written: open the spreadsheet and check 'Form Links & Embed Codes'");
}
