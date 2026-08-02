
const SPREADSHEET_ID = "1XpgEdhMX-OWar77KEOPUOAFl8HX_BIi6sQUSGMnGedg";
const SHEET_NAME = "Pole Register";
const PHOTOS_FOLDER_ID = "1d4Qz8Lh4K4YwmgsQ9gjAwiQvTRp1mNSu";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, service:"Make Ready Agent"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || "{}");
    let result;
    if (req.action === "savePole") result = savePole(req);
    else if (req.action === "uploadPhoto") result = uploadPhoto(req);
    else if (req.action === "getAll") result = getAll();
    else result = {ok:false, error:"Unknown action"};
    return json(result);
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

function json(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(){
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function getHeaders(sheet){
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}

function findRowByProjectPole(sheet, projectPole){
  const vals = sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,1),1).getValues().flat();
  const idx = vals.findIndex(v => Number(v) === Number(projectPole));
  return idx < 0 ? -1 : idx + 2;
}

function savePole(req){
  const sh = getSheet();
  const headers = getHeaders(sh);
  const row = findRowByProjectPole(sh, req.projectPole);
  if(row < 0) throw new Error("Pole not found: " + req.projectPole);

  const map = {
    "Measured Distance (ft)": req.measuredDistance || "",
    "Height Changed?": req.heightChanged || "No",
    "Trimming Required / Notes": req.trimmingNotes || "",
    "Actual HOA": req.actualHoa || "",
    "Before Photo 1": (req.beforeLinks||[])[0] || "",
    "Before Photo 2": (req.beforeLinks||[])[1] || "",
    "Before Photo 3": (req.beforeLinks||[])[2] || "",
    "After Photo 1": (req.afterLinks||[])[0] || "",
    "After Photo 2": (req.afterLinks||[])[1] || "",
    "Status": req.status || "Not started",
    "Field Notes": req.fieldNotes || ""
  };

  Object.keys(map).forEach(name=>{
    const c = headers.indexOf(name);
    if(c >= 0) sh.getRange(row,c+1).setValue(map[name]);
  });
  return {ok:true,row};
}

function uploadPhoto(req){
  const root = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
  const poleFolder = getOrCreateFolder(root, String(req.poleId || ("Pole_"+req.projectPole)));
  const typeFolder = getOrCreateFolder(poleFolder, req.photoType || "OTHER");
  const bytes = Utilities.base64Decode(req.base64);
  const blob = Utilities.newBlob(bytes, req.mimeType || "image/jpeg");
  const ext = (req.mimeType||"image/jpeg").includes("png") ? "png" : "jpg";
  const name = `${req.poleId || req.projectPole}_${req.photoType}_${req.index}.${ext}`;
  const file = typeFolder.createFile(blob).setName(name);
  return {ok:true,url:file.getUrl(),name};
}

function getOrCreateFolder(parent,name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getAll(){
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = name => headers.indexOf(name);
  const rows = values.map(r=>({
    projectPole:r[idx("Project Pole")],
    poleId:r[idx("Pole ID")],
    measuredDistance:r[idx("Measured Distance (ft)")],
    heightChanged:r[idx("Height Changed?")],
    trimmingNotes:r[idx("Trimming Required / Notes")],
    actualHoa:r[idx("Actual HOA")],
    status:r[idx("Status")],
    fieldNotes:r[idx("Field Notes")]
  }));
  return {ok:true,rows};
}
