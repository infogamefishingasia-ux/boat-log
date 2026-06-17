/**
 * Aqua Creed - Boat Log sync backend
 *
 * Deploy as a Web App (Extensions > Apps Script in your Google Sheet):
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Copy the /exec URL into CONFIG.SYNC_URL in app/index.html and app/admin.html.
 *
 * Sheet names expected in this spreadsheet (must match exactly):
 *   "Voyage Log", "Catches", "Config", "Captains", "Admins"
 * All must already have header rows matching Boat_Log_Schema.xlsx.
 */

// Optional shared secret. Leave blank to disable the check.
// If set, the PWA must send the same value as "token" in its JSON body.
const API_TOKEN = "";

// Admin passwords for the dashboard / settings screen (admin.html) now live
// in the "Admins" sheet (Name | Password | Active), so up to several people
// (Darran + managers) can each have their own login. This is "soft" security
// only (same level as the captain PINs) - don't expect it to stop a
// determined attacker.

const VOYAGE_SHEET = "Voyage Log";
const CATCHES_SHEET = "Catches";
const CONFIG_SHEET = "Config";
const CAPTAINS_SHEET = "Captains";
const ADMINS_SHEET = "Admins";

const BOATS = ["RAMPAGE", "REEFRAIDER", "TROPICAL STAR", "HALCYON III"];

// =====================================================================
// doGet - read-only endpoints
//   ?action=bootstrap  -> boats (tank capacity + current balance) + captains (name/pin/active). No auth.
//   ?action=dashboard&password=...  -> fleet aggregates for admin dashboard
//   ?action=csv&sheet=voyage|catches&password=...  -> CSV export
//   (no action) -> health check
// =====================================================================
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "";

    if (action === "bootstrap") {
      return jsonResponse({ status: "ok", boats: getConfigRows(), captains: getCaptainRows() });
    }

    if (action === "dashboard") {
      // IMPORTANT: password is in e.parameter (URL query string), not on e directly.
      const admin = checkAdminPassword(e.parameter);
      if (!admin) {
        return jsonResponse({ status: "error", message: "Invalid admin password" });
      }
      const dash = buildDashboard();
      dash.loggedInAs = admin.name;
      // isSuperAdmin gates the Admin Logins settings card in admin.html
      dash.isSuperAdmin = (admin.name.toLowerCase() === "darran");
      dash.admins = getAdminRows();
      return jsonResponse({ status: "ok", dashboard: dash });
    }

    if (action === "csv") {
      if (!checkAdminPassword(e.parameter)) {
        return jsonResponse({ status: "error", message: "Invalid admin password" });
      }
      const which = (e.parameter.sheet || "voyage").toLowerCase();
      const sheetName = which === "catches" ? CATCHES_SHEET : VOYAGE_SHEET;
      const csv = csvFromSheet(sheetName);
      return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
    }

    return jsonResponse({ status: "ok", message: "Aqua Creed Boat Log API is running" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// =====================================================================
// doPost - write endpoints
//   type="entry"        -> append a voyage + catches row, recalc fuel, update Config balance
//   type="adminUpdate"  -> admin password required, edit Config / Captains sheets
// =====================================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.type === "adminUpdate") {
      const admin = checkAdminPassword(body);
      if (!admin) {
        return jsonResponse({ status: "error", message: "Invalid admin password" });
      }
      return handleAdminUpdate(body, admin);
    }

    // Default: a voyage log entry from a captain
    const entry = body;

    if (API_TOKEN && entry.token !== API_TOKEN) {
      return jsonResponse({ status: "error", message: "Invalid token" });
    }

    if (!entry.entryId || !entry.voyage) {
      return jsonResponse({ status: "error", message: "Missing entryId or voyage data" });
    }

    const fuelUsed = appendVoyageRow(entry);
    appendCatchesRow(entry);
    const syncedBy = entry.voyage.captain ? (entry.voyage.captain + " (voyage sync)") : "Voyage sync";
    updateConfigBalance(entry.voyage.boat, entry.voyage.fuelStop, syncedBy);

    return jsonResponse({ status: "ok", entryId: entry.entryId, fuelUsed: fuelUsed });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// =====================================================================
// Voyage Log / Catches writers
// =====================================================================
function appendVoyageRow(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOYAGE_SHEET);
  const v = entry.voyage;
  const c = entry.catches || {};

  const engPortStart = num(v.engPortStart);
  const engPortStop = num(v.engPortStop);
  const engStbdStart = num(v.engStbdStart);
  const engStbdStop = num(v.engStbdStop);
  const fuelStart = num(v.fuelStart);
  const fuelFilled = num(v.fuelFilled);
  const fuelStop = num(v.fuelStop);

  const portEngRun = engPortStop - engPortStart;
  const stbdEngRun = engStbdStop - engStbdStart;

  // Fuel used is always recomputed server-side - never trust the client value.
  const fuelUsed = fuelStart + fuelFilled - fuelStop;

  const totalCatches = num(c.gtCount) + num(c.dogtoothCount) + num(c.otherFishCount);

  sheet.appendRow([
    entry.entryId,            // A Entry ID
    v.date || "",             // B Date
    v.boat || "",             // C Boat
    v.captain || "",          // D Captain
    v.crew || "",             // E Crew
    v.departurePort || "",    // F Departure Port
    v.arrivalPort || "",      // G Arrival Port
    v.areaFished || "",       // H Area Fished
    num(v.guests),            // I Guests
    v.windDirection || "",    // J Wind Direction
    v.seaCondition || "",     // K Sea Condition
    v.skyWeather || "",       // L Sky / Weather
    engPortStart,             // M Engine Hrs Port - Start
    engPortStop,              // N Engine Hrs Port - Stop
    engStbdStart,             // O Engine Hrs Stbd - Start
    engStbdStop,              // P Engine Hrs Stbd - Stop
    portEngRun,                // Q Port Engine Hrs Run
    stbdEngRun,                // R Stbd Engine Hrs Run
    fuelStart,                 // S Fuel Tank - Start (L)
    fuelFilled,                // T Fuel Filled (L)
    fuelUsed,                  // U Fuel Used (L) [auto-calc]
    fuelStop,                  // V Fuel Tank - Stop (L)
    num(v.distanceRun),       // W Distance Run (NM)
    v.dayNotes || "",          // X Day Notes
    totalCatches,              // Y # Catches (link)
    v.submittedBy || "",       // Z Submitted By
    "Synced " + new Date().toISOString() // AA Sync Status
  ]);

  return fuelUsed;
}

function appendCatchesRow(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATCHES_SHEET);
  const v = entry.voyage;
  const c = entry.catches || {};

  const gtCount = num(c.gtCount);
  const dogtoothCount = num(c.dogtoothCount);
  const otherFishCount = num(c.otherFishCount);
  const gtOver30 = c.gtOver30 || "";
  const dogtoothOver20 = c.dogtoothOver20 || "";
  const gtLength = c.gtLength || "";

  sheet.appendRow([
    entry.entryId,                     // A Entry ID
    v.date || "",                      // B Date
    v.boat || "",                      // C Boat
    gtCount,                            // D GT Count (Total)
    dogtoothCount,                      // E Dogtooth Count (Total)
    otherFishCount,                     // F Other Fish Count (Total)
    gtOver30,                           // G GT Over 30kg or 120cm (comma list)
    dogtoothOver20,                     // H Dogtooth Over 20kg (kg, comma list)
    gtCount + dogtoothCount + otherFishCount, // I Total Catches
    countCommaList(gtOver30),           // J GT >30kg Count (auto)
    countCommaList(dogtoothOver20),     // K Dogtooth >20kg Count (auto)
    gtLength,                            // L GT Length (cm, comma list)
    c.notes || ""                       // M Notes
  ]);
}

// =====================================================================
// Config / Captains - bootstrap reads
// =====================================================================
function getConfigRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({
      boat: String(r[0]),
      tankCapacity: num(r[1]),
      currentBalance: num(r[2]),
      lastUpdated: r[3] ? String(r[3]) : "",
      lastEditedBy: r[5] ? String(r[5]) : ""
    });
  }
  return rows;
}

function getCaptainRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAPTAINS_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({
      name: String(r[0]),
      pin: r[1] === "" || r[1] === null ? "" : String(r[1]),
      active: String(r[2] || "Y").toUpperCase() !== "N",
      lastEditedBy: r[3] ? String(r[3]) : ""
    });
  }
  return rows;
}

// =====================================================================
// Admins (admin dashboard logins)
// =====================================================================
function getAdminRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMINS_SHEET);
  if (!sheet) {
    throw new Error('Sheet "' + ADMINS_SHEET + '" not found. Add it from Boat_Log_Schema.xlsx (upload to Drive > open as Google Sheets).');
  }
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({
      name: String(r[0]),
      password: r[1] === "" || r[1] === null ? "" : String(r[1]),
      active: String(r[2] || "Y").toUpperCase() !== "N",
      lastEditedBy: r[3] ? String(r[3]) : ""
    });
  }
  return rows;
}

function setAdminRows(admins, editedBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMINS_SHEET);
  if (!sheet) {
    throw new Error('Sheet "' + ADMINS_SHEET + '" not found. Add it from Boat_Log_Schema.xlsx.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  if (admins.length === 0) return;
  const stampVal = editedBy ? stamp(editedBy) : "";
  const rows = admins.map(function (a) {
    return [a.name || "", a.password === undefined || a.password === null ? "" : String(a.password), (a.active === false ? "N" : "Y"), stampVal];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

// Update the running balance for a boat after a voyage is logged.
function updateConfigBalance(boat, newBalance, editedBy) {
  if (!boat) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(boat)) {
      sheet.getRange(i + 1, 3).setValue(num(newBalance)); // column C - Current Balance (L)
      sheet.getRange(i + 1, 4).setValue(new Date());      // column D - Last Updated
      if (editedBy) sheet.getRange(i + 1, 6).setValue(stamp(editedBy)); // column F - Last Edited By
      return;
    }
  }
}

// =====================================================================
// Admin updates (settings screen)
//   action="setConfig"   { boat, tankCapacity, currentBalance }
//   action="setCaptains" { captains: [{name, pin, active}, ...] }  (full replace)
// =====================================================================
function handleAdminUpdate(body, admin) {
  const editedBy = admin && admin.name;

  if (body.action === "setConfig") {
    setConfigRow(body.boat, body.tankCapacity, body.currentBalance, editedBy);
    return jsonResponse({ status: "ok", boats: getConfigRows() });
  }

  if (body.action === "setCaptains") {
    setCaptainRows(body.captains || [], editedBy);
    return jsonResponse({ status: "ok", captains: getCaptainRows() });
  }

  if (body.action === "setAdmins") {
    const admins = body.admins || [];
    const hasActiveLogin = admins.some(function (a) {
      return a.active !== false && a.name && String(a.password || "").trim() !== "";
    });
    if (!hasActiveLogin) {
      return jsonResponse({ status: "error", message: "At least one active admin with a password is required - changes not saved." });
    }
    setAdminRows(admins, editedBy);
    return jsonResponse({ status: "ok", admins: getAdminRows() });
  }

  return jsonResponse({ status: "error", message: "Unknown adminUpdate action" });
}

function setConfigRow(boat, tankCapacity, currentBalance, editedBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(boat)) {
      if (tankCapacity !== undefined && tankCapacity !== null && tankCapacity !== "") {
        sheet.getRange(i + 1, 2).setValue(num(tankCapacity));
      }
      if (currentBalance !== undefined && currentBalance !== null && currentBalance !== "") {
        sheet.getRange(i + 1, 3).setValue(num(currentBalance));
      }
      sheet.getRange(i + 1, 4).setValue(new Date());
      if (editedBy) sheet.getRange(i + 1, 6).setValue(stamp(editedBy));
      return;
    }
  }
  // Boat not found - append a new row
  sheet.appendRow([boat, num(tankCapacity), num(currentBalance), new Date(), "Added via Admin Dashboard", editedBy ? stamp(editedBy) : ""]);
}

function setCaptainRows(captains, editedBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAPTAINS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  if (captains.length === 0) return;
  const stampVal = editedBy ? stamp(editedBy) : "";
  const rows = captains.map(function (c) {
    return [c.name || "", c.pin === undefined || c.pin === null ? "" : String(c.pin), (c.active === false ? "N" : "Y"), stampVal];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

// =====================================================================
// Admin dashboard aggregates
// =====================================================================
function buildDashboard() {
  const config = getConfigRows();
  const voyageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOYAGE_SHEET);
  const catchesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATCHES_SHEET);

  const voyageData = voyageSheet.getDataRange().getValues();
  const catchesData = catchesSheet.getDataRange().getValues();

  // Per-boat accumulators
  const perBoat = {};
  BOATS.forEach(function (b) {
    perBoat[b] = {
      fuelUsedTotal: 0, voyageCount: 0,
      portEngHrs: 0, stbdEngHrs: 0,
      lastPortEngStop: 0, lastStbdEngStop: 0,
      gtCount: 0, gtOver30Count: 0, dogtoothCount: 0
    };
  });

  // Voyage Log columns (0-indexed):
  // A=0 entryId, B=1 Date, C=2 Boat, D=3 Captain, E=4 Crew, F=5 DeptPort, G=6 ArrPort,
  // H=7 AreaFished, I=8 Guests, J=9 WindDir, K=10 SeaCond, L=11 Sky,
  // M=12 PortEngStart, N=13 PortEngStop, O=14 StbdEngStart, P=15 StbdEngStop,
  // Q=16 PortEngRun, R=17 StbdEngRun,
  // S=18 FuelStart, T=19 FuelFilled, U=20 FuelUsed, V=21 FuelStop,
  // W=22 Distance, X=23 DayNotes, Y=24 TotalCatches, Z=25 SubmittedBy
  const voyageEntries = [];
  for (let i = 1; i < voyageData.length; i++) {
    const row = voyageData[i];
    if (!row[0]) continue; // skip blank rows
    const boat = String(row[2] || "");
    if (perBoat[boat]) {
      perBoat[boat].fuelUsedTotal += num(row[20]);
      perBoat[boat].voyageCount += 1;
      perBoat[boat].portEngHrs += num(row[16]);
      perBoat[boat].stbdEngHrs += num(row[17]);
      // Last stop reading = most recent odometer value (rows append in order, last wins)
      if (num(row[13]) > 0) perBoat[boat].lastPortEngStop = num(row[13]);
      if (num(row[15]) > 0) perBoat[boat].lastStbdEngStop = num(row[15]);
    }
    // Format date value from the sheet
    const dateVal = row[1];
    let dateStr = "";
    if (Object.prototype.toString.call(dateVal) === "[object Date]") {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      dateStr = String(dateVal || "");
    }
    voyageEntries.push({
      entryId:       String(row[0] || ""),
      date:          dateStr,
      boat:          boat,
      captain:       String(row[3] || ""),
      crew:          String(row[4] || ""),
      departurePort: String(row[5] || ""),
      arrivalPort:   String(row[6] || ""),
      areaFished:    String(row[7] || ""),
      guests:        num(row[8]),
      windDirection: String(row[9] || ""),
      seaCondition:  String(row[10] || ""),
      skyWeather:    String(row[11] || ""),
      engPortStart:  num(row[12]),
      engPortStop:   num(row[13]),
      engStbdStart:  num(row[14]),
      engStbdStop:   num(row[15]),
      portEngRun:    num(row[16]),
      stbdEngRun:    num(row[17]),
      fuelStart:     num(row[18]),
      fuelFilled:    num(row[19]),
      fuelUsed:      num(row[20]),
      fuelStop:      num(row[21]),
      distanceRun:   num(row[22]),
      dayNotes:      String(row[23] || "")
    });
  }
  // Reverse so most recent entry is first
  voyageEntries.reverse();

  // Catches columns (0-indexed): A=0 entryId, B=1 Date, C=2 Boat, D=3 GT, E=4 Dogtooth, J=9 GT>30
  let totalGT = 0;
  let totalGTOver30 = 0;
  let totalDogtooth = 0;
  for (let i = 1; i < catchesData.length; i++) {
    const row = catchesData[i];
    if (!row[0]) continue;
    const boat = String(row[2] || "");
    totalGT += num(row[3]);
    totalDogtooth += num(row[4]);
    totalGTOver30 += num(row[9]);
    if (perBoat[boat]) {
      perBoat[boat].gtCount      += num(row[3]);
      perBoat[boat].gtOver30Count += num(row[9]);
      perBoat[boat].dogtoothCount += num(row[4]);
    }
  }

  const boats = config.map(function (cfg) {
    const s = perBoat[cfg.boat] || {
      fuelUsedTotal: 0, voyageCount: 0, portEngHrs: 0, stbdEngHrs: 0,
      lastPortEngStop: 0, lastStbdEngStop: 0, gtCount: 0, gtOver30Count: 0, dogtoothCount: 0
    };
    return {
      boat:              cfg.boat,
      tankCapacity:      cfg.tankCapacity,
      fuelAvailable:     cfg.currentBalance,
      lastUpdated:       cfg.lastUpdated,
      avgFuelConsumption: s.voyageCount > 0
                            ? Math.round((s.fuelUsedTotal / s.voyageCount) * 10) / 10
                            : 0,
      voyageCount:       s.voyageCount,
      portEngineHours:   s.portEngHrs,
      stbdEngineHours:   s.stbdEngHrs,
      lastPortEngStop:   s.lastPortEngStop,
      lastStbdEngStop:   s.lastStbdEngStop,
      gtCount:           s.gtCount,
      gtOver30Count:     s.gtOver30Count,
      dogtoothCount:     s.dogtoothCount
    };
  });

  return {
    boats: boats,
    totals: { gtCount: totalGT, gtOver30Count: totalGTOver30, dogtoothCount: totalDogtooth },
    voyageEntries: voyageEntries
  };
}

// =====================================================================
// CSV export
// =====================================================================
function csvFromSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  return data.map(function (row) {
    return row.map(csvCell).join(",");
  }).join("\r\n");
}

function csvCell(value) {
  if (value === null || value === undefined) return "";
  let s;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    s = value.toISOString();
  } else {
    s = String(value);
  }
  if (/[",\r\n]/.test(s)) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

// =====================================================================
// Helpers
// =====================================================================
// Returns the matching admin record ({name, password, active}) if
// source.password matches an active row in the Admins sheet, or null.
function checkAdminPassword(source) {
  if (!source || !source.password) return null;
  const admins = getAdminRows();
  for (let i = 0; i < admins.length; i++) {
    const a = admins[i];
    if (a.active && a.password && a.password === source.password) {
      return a;
    }
  }
  return null;
}

function countCommaList(str) {
  if (!str || String(str).trim() === "") return 0;
  return String(str).split(",").filter(x => x.trim() !== "").length;
}

function num(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}

// "Darran — 2026-06-14 10:32 UTC" - stamped into "Last Edited By" columns
// on Config / Captains / Admins so changes are traceable across logins.
function stamp(name) {
  if (!name) return "";
  return name + " — " + new Date().toISOString().slice(0, 16).replace("T", " ") + " UTC";
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
