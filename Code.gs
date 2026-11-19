/**
 * ERP Implementation Dashboard — server side.
 *
 * Reads the BA, Scope, AFI, Reports and Issues tabs and returns them as JSON to the
 * dashboard page (index.html, hosted on GitHub Pages).
 * Nothing here is hard-coded to a week or a department: add a week block,
 * add a department row, and the dashboard picks it up on the next load.
 *
 * Expected layout of every tab (as it is today):
 *   row with week names   ("1st Week", "2nd Week", ... merged across the block)
 *   row with week dates   (10/31/2026 ...)
 *   header row            (No | Department | metric | metric | ...)
 *   one row per department
 *
 * Deploy with  Execute as: Me  and  Who has access: Anyone  (the page must be able to read it
 * without a Google sign-in). Paste the /exec link into DATA_URL at the top of index.html.
 */

// Used only if the script is NOT opened from inside the sheet (Extensions > Apps Script).
var SHEET_ID = '1QO0LqAgwiXZ53j3iqwgaOyze_VV9YUF3-EXgFPY-G4A';

// Tab names to read. Rename here if a tab is renamed in the sheet.
// "Scope" has no week blocks (just No | Department | Scope Doc | Scope Doc Approval); it is read as one fixed set of values.
var TABS = ['BA', 'Scope', 'AFI', 'Reports', 'Issues'];

/**
 * DATA FEED. The dashboard page itself is hosted on GitHub Pages; this web app only returns
 * the sheet figures as JSON. Add ?brand=1 to also get the company name and logo.
 */
function doGet(e) {
  var wantBrand = !!(e && e.parameter && e.parameter.brand === '1');
  var out;
  try {
    out = getDashboardData(wantBrand);
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Shown in the page heading.
var COMPANY_NAME = 'Imo Chicken & Agro (Pvt) Ltd.';

// Company logo: upload the image to Google Drive, open it, and copy the ID from its link
// (https://drive.google.com/file/d/THIS_PART/view). Paste the ID between the quotes. Leave empty for no logo.
var LOGO_FILE_ID = '1IZ8nwjoCYNVkLau2Xq7WrYPyIwrsLXrj';

/**
 * Run this ONCE from the editor after setting LOGO_FILE_ID: choose "authorizeLogo" in the function
 * list at the top and press Run. Google then asks for permission to read the logo file from Drive.
 * Without that approval the web app cannot open the file. Then deploy a new version.
 */
function authorizeLogo() {
  var file = DriveApp.getFileById(LOGO_FILE_ID);
  Logger.log('Logo found: ' + file.getName() + ' (' + file.getMimeType() + ', ' + file.getSize() + ' bytes)');
}

/** Company name and logo (as an embedded image) for the heading. Sent once, on first load. */
function getBrand_() {
  var brand = { name: COMPANY_NAME, logo: '' };
  if (LOGO_FILE_ID) {
    try {
      var blob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
      brand.logo = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      // Could not read the file (usually: Drive permission not approved yet). The page then tries the
      // file's public thumbnail, which works only if the file is shared as "Anyone with the link".
      brand.logoError = String(e && e.message || e);
      brand.logoUrl = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(LOGO_FILE_ID) + '&sz=w400';
    }
  }
  return brand;
}

/** Called from the page (google.script.run) on load and on every refresh. */
function getDashboardData(wantBrand) {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var grids = {};
  TABS.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var range = sheet.getDataRange();
    grids[name] = { values: range.getValues(), display: range.getDisplayValues() };
  });
  var payload = buildPayload_(grids, tz, ss.getName());
  if (wantBrand) payload.brand = getBrand_();
  return payload;
}

function buildPayload_(grids, tz, title) {
  var tabs = {};
  var missing = [];
  TABS.forEach(function (name) {
    var parsed = grids[name] ? parseTab_(grids[name], tz) : null;
    if (parsed) tabs[name] = parsed; else missing.push(name);
  });
  return { title: title || '', tabs: tabs, missing: missing, generatedAt: new Date().getTime() };
}

/** Turns one tab into { departments: [...], weeks: [{label, date, metrics: [{name, values}]}] }. */
function parseTab_(grid, tz) {
  var values = grid.values, display = grid.display;
  var hdr = -1, deptCol = -1, r, c;

  // Find the header row: the row that has a "Department" cell.
  for (r = 0; r < Math.min(display.length, 15) && hdr < 0; r++) {
    for (c = 0; c < display[r].length; c++) {
      if (clean_(display[r][c]).toLowerCase() === 'department') { hdr = r; deptCol = c; break; }
    }
  }
  if (hdr < 0) return null;

  var labelRow = hdr >= 2 ? hdr - 2 : -1;
  var dateRow = hdr >= 1 ? hdr - 1 : -1;
  var weeks = [], current = null, weekly = false;   // weekly = the tab has "1st Week", "2nd Week" ... titles

  for (c = deptCol + 1; c < display[hdr].length; c++) {
    var label = labelRow >= 0 ? clean_(display[labelRow][c]) : '';
    var metric = clean_(display[hdr][c]);
    if (label) weekly = true;
    // A merged week title only has a value in its first cell, so a new title starts a new week.
    if (label || (!current && metric)) {
      current = { label: label || ('Week ' + (weeks.length + 1)), date: '', cols: [] };
      weeks.push(current);
    }
    if (!current || !metric) continue;
    if (!current.date && dateRow >= 0) current.date = dateText_(values[dateRow][c], display[dateRow][c], tz);
    current.cols.push({ name: metric, col: c });
  }

  var departments = [], rows = [];
  for (r = hdr + 1; r < display.length; r++) {
    var dept = clean_(display[r][deptCol]);
    if (!dept || /^(grand\s+)?total$/i.test(dept)) continue;   // skips blank rows and the totals row
    departments.push(dept);
    rows.push(r);
  }

  return {
    departments: departments,
    weekly: weekly,
    weeks: weeks.map(function (w) {
      return {
        label: w.label,
        date: w.date,
        metrics: w.cols.map(function (m) {
          return { name: m.name, values: rows.map(function (row) { return cell_(display[row][m.col]); }) };
        })
      };
    })
  };
}

function clean_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim();
}

/** "60%" -> 60, "1,250" -> 1250, "Yes" -> "Yes", "" -> null */
function cell_(shown) {
  var s = clean_(shown);
  if (!s) return null;
  var n = s.replace(/,/g, '');
  if (/^-?\d+(\.\d+)?%$/.test(n)) return parseFloat(n);
  if (/^-?\d+(\.\d+)?$/.test(n)) return Number(n);
  return s;
}

/** Returns yyyy-MM-dd for a real date cell, otherwise whatever text is in the cell. */
function dateText_(value, shown, tz) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    if (typeof Utilities !== 'undefined') return Utilities.formatDate(value, tz || 'GMT', 'yyyy-MM-dd');
    var p = function (x) { return (x < 10 ? '0' : '') + x; };
    return value.getFullYear() + '-' + p(value.getMonth() + 1) + '-' + p(value.getDate());
  }
  return clean_(shown);
}
