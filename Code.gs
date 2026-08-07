const PRODUCTION_SPREADSHEET_ID = "1pb046ZhvDjc-9VX4KqeqUlmJyp0yxT9nSe8Vnt_3Vt0";
const PRODUCTION_SHEET_NAME = "Blank";
const PHOTOS_FOLDER_ID = "1d4Qz8Lh4K4YwmgsQ9gjAwiQvTRp1mNSu";

// Вставь сюда ссылку на свою Google My Maps.
const MY_MAPS_URL = "PASTE_GOOGLE_MY_MAPS_URL_HERE";

const WORK_ROWS = {
  newStrand: "PLACE  NEW STRAND",
  installDownGuy: "INSTALL DOWNGUY",
  reworkDownGuy: "TRANSFER / REWORK EXISTING DOWN GUY",
  overheadGuy: "PLACE OVERHEAD GUY",
  groundBond: "INSTALL POLE GROUND AND BOND",
  raiseLower: "RAISE OR LOWER  POLE ATTACHMENT",
  riserGuard: "INSTALL NEW RISER GUARD TO SECURE CABLES TO POLES",
  treeTrimming: "TREE TRIMMING",
  fArms: "PLACE F-ARMS",
  guardArm: "PLACE GUARD ARM",
  doubleGuardArm: "PLACE DOUBLE GUARD ARM",
  removeArm: "REMOVE ARM",
  poleTransfer: "POLE TRANSFER"
};

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Make Ready Agent");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getConfig() {
  return {myMapsUrl: MY_MAPS_URL};
}

function getProductionSheet_() {
  const sh = SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID)
    .getSheetByName(PRODUCTION_SHEET_NAME);
  if (!sh) throw new Error("Production sheet not found: " + PRODUCTION_SHEET_NAME);
  return sh;
}

function normalize_(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
}

function findDescriptionRow_(sh, description) {
  const lastRow = sh.getLastRow();
  const values = sh.getRange(1, 1, lastRow, 1).getDisplayValues().flat();
  const target = normalize_(description);

  const index = values.findIndex(v => {
    const text = normalize_(v);
    return text === target || text.startsWith(target) || target.startsWith(text);
  });

  if (index < 0) throw new Error("Production row not found: " + description);
  return index + 1;
}

function getPoleHeaderRow_(sh) {
  const values = sh.getRange(1, 1, Math.min(sh.getLastRow(), 20), sh.getLastColumn())
    .getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    if (values[r].filter(v => normalize_(v) === "POLE #").length >= 2) return r + 1;
  }
  throw new Error("Pole header row not found");
}

function getPoleNumberRow_(sh) {
  return getPoleHeaderRow_(sh) + 1;
}

function findTotalColumn_(sh) {
  const lastCol = sh.getLastColumn();
  const headerRows = sh.getRange(1, 1, Math.min(sh.getLastRow(), 15), lastCol)
    .getDisplayValues();

  for (let c = 0; c < lastCol; c++) {
    if (headerRows.some(row => normalize_(row[c]) === "TOTAL")) return c + 1;
  }
  return lastCol;
}

function ensurePoleColumn_(sh, projectPole) {
  const poleRow = getPoleNumberRow_(sh);
  let lastCol = sh.getLastColumn();
  let values = sh.getRange(poleRow, 1, 1, lastCol).getDisplayValues()[0];

  let found = values.findIndex(v => Number(v) === Number(projectPole));
  if (found >= 0) return found + 1;

  let totalCol = findTotalColumn_(sh);

  // Используем пустую колонку перед TOTAL, если она есть.
  for (let c = 5; c < totalCol; c++) {
    if (!values[c - 1]) {
      sh.getRange(poleRow - 1, c).setValue("Pole #");
      sh.getRange(poleRow, c).setValue(projectPole);
      return c;
    }
  }

  // Если пустых колонок нет — вставляем новую перед TOTAL.
  sh.insertColumnBefore(totalCol);

  // Копируем формат предыдущей pole-колонки.
  const sourceCol = Math.max(5, totalCol - 1);
  sh.getRange(1, sourceCol, sh.getMaxRows(), 1)
    .copyTo(sh.getRange(1, totalCol, sh.getMaxRows(), 1), {formatOnly: true});

  sh.getRange(poleRow - 1, totalCol).setValue("Pole #");
  sh.getRange(poleRow, totalCol).setValue(projectPole);
  return totalCol;
}

function toNumberOrBlank_(value) {
  if (value === "" || value === null || value === undefined) return "";
  const n = Number(value);
  return Number.isFinite(n) ? n : value;
}

function updateProduction_(projectPole, work) {
  const sh = getProductionSheet_();
  const poleCol = ensurePoleColumn_(sh, projectPole);

  Object.entries(WORK_ROWS).forEach(([key, description]) => {
    const row = findDescriptionRow_(sh, description);
    sh.getRange(row, poleCol).setValue(toNumberOrBlank_(work[key]));
  });

  SpreadsheetApp.flush();
}

function getChangesSheet_() {
  const ss = SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID);
  let sh = ss.getSheetByName("Changes Log");

  if (!sh) {
    sh = ss.insertSheet("Changes Log");
    sh.getRange(1, 1, 1, 17).setValues([[
      "Timestamp", "Project Pole", "Pole ID",
      "Original HOA", "Actual HOA", "Height Change Description",
      "Anchor Status", "Anchor Details",
      "Bonding Status", "Bonding Details",
      "VGR Status", "VGR Details",
      "Down Guy Actual", "Reason / Field Condition",
      "Status", "Field Notes", "Crew"
    ]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveChangeLog_(data) {
  const sh = getChangesSheet_();
  sh.appendRow([
    new Date(),
    data.projectPole || "",
    data.poleId || "",
    data.originalHoa || "",
    data.actualHoa || "",
    data.heightChangeDescription || "",
    data.anchorStatus || "",
    data.anchorDetails || "",
    data.bondingStatus || "",
    data.bondingDetails || "",
    data.vgrStatus || "",
    data.vgrDetails || "",
    data.downGuyActual || "",
    data.changeReason || "",
    data.status || "",
    data.fieldNotes || "",
    data.crew || ""
  ]);
}

function getLatestChanges() {
  const sh = getChangesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 17).getDisplayValues();
  const latest = {};

  values.forEach(r => {
    const pole = Number(r[1]);
    if (!pole) return;
    latest[pole] = {
      projectPole: pole,
      poleId: r[2],
      originalHoa: r[3],
      actualHoa: r[4],
      heightChangeDescription: r[5],
      anchorStatus: r[6],
      anchorDetails: r[7],
      bondingStatus: r[8],
      bondingDetails: r[9],
      vgrStatus: r[10],
      vgrDetails: r[11],
      downGuyActual: r[12],
      changeReason: r[13],
      status: r[14] || "Not started",
      fieldNotes: r[15],
      crew: r[16]
    };
  });

  return Object.values(latest);
}

function savePole(data) {
  updateProduction_(data.projectPole, data.work || {});
  saveChangeLog_(data);
  return {ok: true};
}

function uploadPhoto(payload) {
  const root = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
  const poleName = String(payload.poleId || ("Pole_" + payload.projectPole));
  const poleFolder = getOrCreateFolder_(root, poleName);
  const typeFolder = getOrCreateFolder_(poleFolder, payload.photoType || "OTHER");

  const bytes = Utilities.base64Decode(payload.base64);
  const mime = payload.mimeType || "image/jpeg";
  const ext = mime.includes("png") ? "png" : "jpg";
  const safeIndex = String(payload.index || 1).padStart(3, "0");
  const name = `${poleName}_${payload.photoType}_${safeIndex}.${ext}`;
  const blob = Utilities.newBlob(bytes, mime, name);
  const file = typeFolder.createFile(blob);

  return {ok: true, url: file.getUrl(), name: name};
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
