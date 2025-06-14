/**********************************************/
/* SECTION 1: GLOBAL CONSTANTS & CONFIG REF   */
/**********************************************/

const SPREADSHEETS = { PROJECT: null, DATABASE: null };
const SHEETS = { REPORT_INPUT: null, PROPOSED: null, EXISTING: null, SYSTEMS: null, CHARTS: null };
const DATA = { INPUT: null, PROPOSED: null, EXISTING: null, SYSTEMS: null };


// Helper: Get config for a project type (ROOT or LEAF)
function getConfig(type) {
  return config.PROJECT_SOURCES[type] || config.PROJECT_SOURCES.ROOT;
}

// Get a Google Sheet by config and name
function getSheetByConfig(type, sheetName) {
  const cfg = getConfig(type);
  return SpreadsheetApp.openById(config.PROJECT_DATABASE_ID).getSheetByName(sheetName);
}

// Get all projects for a given type (ROOT/LEAF)
function getAllProjects(type) {
  const sheet = getSheetByConfig(type, "Projects");
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0].map(h => h.trim());
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// Log a new project entry to the project DB
function logProjectToDatabase(project, type) {
  const sheet = getSheetByConfig(type, "Projects");
  sheet.appendRow([
    project.sheetUrl, project.docUrl, project.status, project.created,
    project.company, project.property, project.address, project.city,
    project.state, project.zip, project.reportId, project.note
  ]);
}

// Create a new project (ROOT/LEAF), copy templates, log DB row
function createNewProject(data, type) {
  const cfg = getConfig(type);
  const dbSheet = getSheetByConfig(type, "Projects");
  const sequenceNumber = dbSheet.getLastRow().toString().padStart(3, '0');
  const year = new Date().getFullYear().toString().slice(-2);
  const zip = data.zip;
  const addressNum = (data.addressLine1.match(/^\d+/)?.[0] || "NoNum");
  const reportId = `${type}-${year}.${sequenceNumber}.${zip}-${addressNum}`;

  const folder = DriveApp.getFolderById(cfg.PROJECTS_FOLDER_ID);
  const newSheetFile = DriveApp.getFileById(cfg.TEMPLATE_SHEET_ID).makeCopy(reportId, folder);
  const sheet = SpreadsheetApp.open(newSheetFile);
  const input = sheet.getSheetByName(config.SHEET_NAMES.REPORT_INPUT
);

  // Map incoming fields to input sheet
  const mapping = {
    "Property Name": data.propertyName,
    "Client Company Name": data.companyName,
    "Property Address Line1": data.addressLine1,
    "City": data.city,
    "State": data.state,
    "Property ZIP": data.zip,
    "ReportID": reportId
  };
  const inputData = input.getDataRange().getValues();
  inputData.forEach((row, r) => {
    const label = (row[1] || '').toString().trim();
    if (mapping[label] !== undefined) {
      input.getRange(r + 1, 3).setValue(mapping[label]);
    }
  });

  const newDocFile = DriveApp.getFileById(config.DOC_TEMPLATE_ID).makeCopy(reportId, folder);

  // Log to DB
  logProjectToDatabase({
    reportId,
    property: data.propertyName,
    address: data.addressLine1,
    city: data.city,
    zip: data.zip,
    state: data.state,
    company: data.companyName,
    sheetUrl: newSheetFile.getUrl(),
    docUrl: newDocFile.getUrl(),
    status: 'Draft',
    created: new Date().toLocaleString(),
    note: ''
  }, type);

  // If LEAF, generate DVI url and save
  if (type === 'LEAF') {
    const lastRow = dbSheet.getLastRow();
    const reportIdCol = 11; // Adjust if your DB cols ever change
    const dviurlCol = 13;
    const savedReportId = dbSheet.getRange(lastRow, reportIdCol).getValue();
    const dviUrl = getDVIUrlForReportId(savedReportId);
    dbSheet.getRange(lastRow, dviurlCol).setValue(dviUrl);
  }
  return { reportId };
}

// Generate DVI external request URL for a reportId
function getDVIUrlForReportId(reportId) {
  // Replace 'your-deployed-app-id' with your own Google Apps Script Web App ID!
  return `https://script.google.com/macros/s/AKfycbwMZKNXpVI-BDSGpE0BAK638RtI7DmpujacQyaRX2Vkm3lYrOaLLXFO0lgokSpLtuw/exec?reportid=${encodeURIComponent(reportId)}`;
}

// Update a field (status or note) in Projects DB for a reportId
function updateProjectField(reportId, col, value, type) {
  const sheet = getSheetByConfig(type, "Projects");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === reportId) {
      sheet.getRange(i + 1, col).setValue(value);
      break;
    }
  }
}
function updateProjectStatus(type, reportId, newStatus) {
  var sheetId = (type === 'leaf')
    ? '12yJhaOAe4rHSCFSadanh1K1YCu8wJnjbxfBdeX_GeB8'
    : '193m8yOy51aDwSvqvCQQv7uM-fpiJrrnNK2kqCGYLq7E';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var reportCol = 10; // ReportID
  var statusCol = 2;  // Status
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][reportCol]) === String(reportId)) {
      sheet.getRange(i+1, statusCol+1).setValue(newStatus);
      return true;
    }
  }
  throw new Error("Project not found");
}

function updateProjectNote(reportId, note, type) {
  updateProjectField(reportId, 12, note, type);
}

// Delete a project (move files to trash, delete row from DB)
function deleteProject(reportId, type) {
  const sheet = getSheetByConfig(type, "Projects");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === reportId) {
      [data[i][0], data[i][1]].forEach(url => {
        if (url) {
          try {
            const fileId = url.match(/[-\w]{25,}/)[0];
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {
            Logger.log(`⚠️ Error trashing file for ${reportId}: ${err}`);
          }
        }
      });
      sheet.deleteRow(i + 1);
      Logger.log(`🗑️ Deleted project row for ${reportId}`);
      break;
    }
  }
}

// LEAF → DVI Status Update (calls DVI workflow trigger if needed)
function updateLEAFProjectStatus(projectId, newStatus) {
  const ss = SpreadsheetApp.openById(config.PROJECT_SOURCES.LEAF.PROJECT_DATABASE_ID);
  const sheet = ss.getSheetByName('Projects');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf('projectId');
  const statusIdx = header.indexOf('status');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == projectId) {
      sheet.getRange(r + 1, statusIdx + 1).setValue(newStatus);
      if (newStatus === "DVI Requested") {
        triggerDVIWorkflow(projectId);
      }
      return { success: true };
    }
  }
  return { success: false, error: "Project not found" };
}

// Copy LEAF Project to DVI Sheet (triggered from above)
function triggerDVIWorkflow(projectId) {
  const leafSS = SpreadsheetApp.openById(config.PROJECT_SOURCES.LEAF.PROJECT_DATABASE_ID);
  const leafSheet = leafSS.getSheetByName('Projects');
  const dviSS = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const dviSheet = dviSS.getSheetByName(config.DVI_PROJECTS_SHEET);

  const leafData = leafSheet.getDataRange().getValues();
  const leafHeader = leafData[0];
  const idIdx = leafHeader.indexOf('projectId');
  const rowIdx = leafData.findIndex((row, idx) => idx && row[idIdx] == projectId);
  if (rowIdx === -1) return { success: false, error: "LEAF project not found" };

  const dviHeader = dviSheet.getDataRange().getValues()[0];
  const dviRow = Array(dviHeader.length).fill('');
  function safeIdx(h, name) { const idx = h.indexOf(name); return idx !== -1 ? idx : null; }
  dviRow[safeIdx(dviHeader, 'reportid')] = leafData[rowIdx][safeIdx(leafHeader, 'projectId')];
  dviRow[safeIdx(dviHeader, 'address')] = leafData[rowIdx][safeIdx(leafHeader, 'addressLine1')];
  dviRow[safeIdx(dviHeader, 'citystatezip')] = leafData[rowIdx][safeIdx(leafHeader, 'zip')];
  dviRow[safeIdx(dviHeader, 'ownername')] = leafData[rowIdx][safeIdx(leafHeader, 'propertyName')];
  dviRow[safeIdx(dviHeader, 'status')] = "DVI Requested";
  dviRow[safeIdx(dviHeader, 'created')] = new Date();

  // Only add if not present
  const dviData = dviSheet.getDataRange().getValues();
  const dviIdIdx = dviHeader.indexOf('reportid');
  if (!dviData.some((r, idx) => idx && r[dviIdIdx] == projectId)) {
    dviSheet.appendRow(dviRow);
  }
  return { success: true };
}

// INTERNAL ADMIN: Get all DVI jobs (plus contractor list for Kanban UI)
function getAllDVIProjects(token) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "internal") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];

  // Contractors
  const contractorsSheet = ss.getSheetByName(config.CONTRACTORS_SHEET_NAME);
  const cHeader = contractorsSheet.getDataRange().getValues()[0];
  const contractors = contractorsSheet.getDataRange().getValues().slice(1).map(r => ({
    id: r[cHeader.indexOf("contractorid")] || r[cHeader.indexOf("id")] || r[0],
    name: r[cHeader.indexOf("name")] || r[cHeader.indexOf("company")] || "",
    email: r[cHeader.indexOf("email")] || "",
  }));

  // Build jobs array (attach readable contractor name, and try to attach system list)
  let jobs = data.slice(1).map(row => {
    let jobObj = {};
    header.forEach((h, i) => jobObj[h] = row[i]);
    jobObj.contractorName = "";
    if (jobObj.contractorid) {
      const c = contractors.find(c => c.id == jobObj.contractorid);
      if (c) jobObj.contractorName = c.name;
    }
    // Try to load system list from LEAF file if present
    if (jobObj.leafprojectid || jobObj.reportid) {
      try {
        const leafDb = SpreadsheetApp.openById(config.PROJECT_SOURCES.LEAF.PROJECT_DATABASE_ID);
        const leafDbSheet = leafDb.getSheetByName('Projects');
        const leafDbRows = leafDbSheet.getDataRange().getValues();
        const projectIdCol = leafDbRows[0].indexOf("projectId") !== -1 ? leafDbRows[0].indexOf("projectId") : 10;
        const sheetUrlCol = leafDbRows[0].indexOf("sheetUrl") !== -1 ? leafDbRows[0].indexOf("sheetUrl") : 0;
        const projRow = leafDbRows.find(r => r[projectIdCol] == (jobObj.leafprojectid || jobObj.reportid));
        if (projRow) {
          const sheetId = projRow[sheetUrlCol].match(/[-\w]{25,}/)[0];
          const inputSheet = SpreadsheetApp.openById(sheetId).getSheetByName(config.SHEET_NAMES.REPORT_INPUT);
          const systems = inputSheet.getRange(12, 5, 39, 1).getValues()
            .map(r => r[0]).filter(s => s && s !== "Placeholder");
          jobObj.systemsList = Array.from(new Set(systems));
        } else {
          jobObj.systemsList = [];
        }
      } catch (e) {
        jobObj.systemsList = [];
      }
    } else {
      jobObj.systemsList = [];
    }
    return jobObj;
  });
  return { header, jobs, contractors };
}

// INTERNAL: Update DVI job status, notes, assignment (admin Kanban actions)
function updateDVIJobStatus(token, jobId, newStatus) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "internal") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const statusIdx = header.indexOf("status");
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == jobId) {
      sheet.getRange(r + 1, statusIdx + 1).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false, error: "Job not found" };
}
function updateAdminJobNotes(token, jobId, notes) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "internal") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const notesIdx = header.indexOf("reinotes");
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == jobId) {
      sheet.getRange(r + 1, notesIdx + 1).setValue(notes);
      return { success: true };
    }
  }
  return { success: false, error: "Job not found" };
}
function assignContractorToJob(token, jobId, contractorId) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "internal") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const conIdx = header.indexOf("contractorid");
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == jobId) {
      sheet.getRange(r + 1, conIdx + 1).setValue(contractorId);
      return { success: true };
    }
  }
  return { success: false, error: "Job not found" };
}

// CONTRACTOR: Get all jobs assigned to contractor (uses contractorId from user object)
function getContractorJobs(token) {
  var session = getUserFromToken(token);
  if (!session.success) return { header: [], jobs: [] };
  var user = session.user;
  if (user.role !== 'Contractor') return { header: [], jobs: [] };
  var contractorId = String(user.contractorId).trim();

  // Load your DVI projects sheet (replace with correct IDs/names)
  var sheet = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID).getSheetByName(config.DVI_PROJECTS_SHEET);
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var contractorIdCol = header.indexOf("ContractorID"); // Adjust to your column name
  if (contractorIdCol === -1) return { header: header, jobs: [] };
  var filteredJobs = [];
  for (var i = 1; i < data.length; i++) {
    var assigned = (data[i][contractorIdCol] || "").toString();
    // Split on comma, trim, compare case-insensitive
    var assignedList = assigned.split(",").map(function(x) { return x.trim().toLowerCase(); });
    if (assignedList.includes(contractorId.toLowerCase())) {
      filteredJobs.push(data[i]);
    }
  }
  return { header: header, jobs: filteredJobs };
}

function updateContractorJobStatus(token, jobId, newStatus) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "contractor") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const statusIdx = header.indexOf("status");
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == jobId) {
      sheet.getRange(r + 1, statusIdx + 1).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false, error: "Job not found" };
}
function updateContractorJobNotes(token, jobId, notes) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  if (!user || user.role.toLowerCase() !== "contractor") throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const notesIdx = header.indexOf("contractornotes");
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == jobId) {
      sheet.getRange(r + 1, notesIdx + 1).setValue(notes);
      return { success: true };
    }
  }
  return { success: false, error: "Job not found" };
}

// Export PDF for Internal/Contractor
function getProjectPDFUrl(token, reportid) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf("reportid");
  const docCol = header.indexOf("docfileid");
  const contractorCol = header.indexOf("contractorid");
  if (idCol === -1 || docCol === -1 || contractorCol === -1)
    throw new Error("Missing columns in dvi projects sheet.");

  let docFileId = null;
  for (let i = 1; i < data.length; i++) {
    if (
      ("" + data[i][idCol]) === ("" + reportid) &&
      ((user.role.toLowerCase() === "internal") ||
        (user.role.toLowerCase() === "contractor" && ("" + data[i][contractorCol]) === user.contractorId))
    ) {
      docFileId = data[i][docCol];
      break;
    }
  }
  if (!docFileId) throw new Error("No Google Doc found or job not assigned to you.");

  // Convert Doc to PDF Blob, store in temp, share
  var docFile = DriveApp.getFileById(docFileId);
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(docFile.getName() + ".pdf");
  var tempFolder = DriveApp.getRootFolder();
  var pdfFile = tempFolder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = pdfFile.getUrl();
  return { url: url };
}

// DVI Customer Request Form Prefill Info (for dvirequest.html)
function getDVIRequestFormInfo(reportid) {
  var configLeaf = config.PROJECT_SOURCES.LEAF;
  var db = SpreadsheetApp.openById(configLeaf.PROJECT_DATABASE_ID);
  var sheet = db.getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var rowIdx = data.findIndex(row => row[10] && row[10].toString() === reportid);
  if (rowIdx === -1) throw new Error('Project not found');
  var row = data[rowIdx];
  var name = row[6] || "";   // property (address) owner name
  var address = row[6] || "";   // property address
  var phone = row[9] || "";   // owner phone (adjust as needed)
  var sheetUrl = row[0];
  var sheetId = sheetUrl.match(/[-\w]{25,}/)[0];
  var projectSheet = SpreadsheetApp.openById(sheetId);
  var inputSheet = projectSheet.getSheetByName(config.SHEET_NAMES.REPORT_INPUT);
  var systemsList = [];
  if (inputSheet) {
    var sysVals = inputSheet.getRange(12, 5, 39, 1).getValues(); // E12:E50
    sysVals.forEach(r => {
      var val = (r[0] || "").trim();
      if (
        val &&
        val.toLowerCase() !== "placeholder" &&
        !systemsList.includes(val)
      ) {
        systemsList.push(val);
      }
    });
  }
  return {
    name: name,
    address: address,
    phone: phone,
    reportid: reportid,
    systems: systemsList
  };
}

// CONTRACTORS: CRUD for admin manage-contractors page

function getContractorSheet() {
  const ss = SpreadsheetApp.openById(config.CONTRACTORS_SPREADSHEET_ID);
  return ss.getSheetByName(config.CONTRACTORS_SHEET_NAME);
}

// GET ALL CONTRACTORS
function getAllContractors(token) {
  try {
    const sheet = getContractorSheet();
    const values = sheet.getDataRange().getValues();
    const headers = values.shift(); // First row = header

    const contractors = values.map(row => {
      const obj = {};
      headers.forEach((key, i) => {
        obj[key] = row[i];
      });
      return obj;
    });

    return { success: true, contractors: contractors };
  } catch (err) {
    return { success: false, error: err.message, contractors: [] };
  }
}


// SAVE (CREATE or UPDATE) CONTRACTOR
function saveContractor(contractor) {
  try {
    const sheet = getContractorSheet();
    const data = sheet.getDataRange().getValues();
    const header = data[0];

    const idIdx = header.indexOf("ContractorID");
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      if ((data[i][idIdx] || "") === (contractor.ContractorID || "")) {
        foundRow = i;
        break;
      }
    }

    const row = [
      contractor.ContractorID,
      contractor.CompanyName,
      contractor.ContactName,
      contractor.Email,
      contractor.Phone,
      contractor.Industries,
      contractor.Location,
      contractor.ContractTerms,
      contractor.InternalNotes,
      foundRow === -1 ? new Date() : data[foundRow][header.indexOf("DateAdded")],
      contractor.Status || "Active"
    ];

    if (foundRow === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(foundRow + 1, 1, 1, row.length).setValues([row]);
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// DELETE CONTRACTOR
function deleteContractor(contractorId) {
  try {
    const sheet = getContractorSheet();
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const idIdx = header.indexOf("ContractorID");

    for (let i = 1; i < data.length; i++) {
      if ((data[i][idIdx] || "") === (contractorId || "")) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }

    return { success: false, error: "Contractor not found" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


function doGet(e) {
  // If this is a public dvirequest link, serve that form only
  if (e && e.parameter && e.parameter.reportid) {
    return HtmlService.createHtmlOutputFromFile('dvirequest')
      .setTitle("Data Verification Inspection – Renewable Energy Incentives")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }

  // Otherwise, always serve the SPA shell (index.html)
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("Renewable Energy Incentives")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * submitDVIRequest(payload)
 * Receives {reportid, systems:[], notes}
 *  - Updates LEAF project status to "DVI Requested"
 *  - Triggers DVI workflow
 *  - Emails your support team
 */
function submitDVIRequest(payload) {
  var reportid = payload.reportid;
  var systems = payload.systems || [];
  var notes = payload.notes || "";

  // 1. Update project status in LEAF DB
  var leafCfg = config.PROJECT_SOURCES.LEAF;
  var db = SpreadsheetApp.openById(leafCfg.PROJECT_DATABASE_ID);
  var sheet = db.getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var statusCol = header.indexOf('status');
  var notesCol = header.indexOf('reinotes') !== -1 ? header.indexOf('reinotes') : header.length - 1;
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] && data[i][10].toString() === reportid) {
      // Update status
      if (statusCol !== -1) sheet.getRange(i + 1, statusCol + 1).setValue("DVI Requested");
      // Append note about systems requested, if notesCol exists
      var noteStr = "DVI Requested (" + new Date().toLocaleString() + "): " +
        "Systems: " + systems.join(', ');
      if (notes) noteStr += " | Notes: " + notes;
      if (notesCol !== -1) sheet.getRange(i + 1, notesCol + 1).setValue(noteStr);
      found = true;
      break;
    }
  }
  if (!found) throw new Error("Project not found in LEAF Projects");

  // 2. Move to DVI board (call your existing workflow)
  triggerDVIWorkflow(reportid);

  // 3. Notify support
  try {
    MailApp.sendEmail({
      to: "Support@RenewableEnergyIncentives.com",
      subject: "New DVI Request Submitted",
      htmlBody:
        "<b>DVI Request Submitted</b><br>" +
        "Report ID: " + reportid + "<br>" +
        "Systems: " + systems.join(", ") + "<br>" +
        (notes ? "Notes: " + notes + "<br>" : "") +
        "Submitted: " + new Date().toLocaleString()
    });
  } catch (err) {
    Logger.log("Failed to send support email: " + err);
    // Not fatal; still finish
  }
  return { success: true };
}

/**
 * updateDVIProject: Multi-contractor admin modal save
 * Supports multi-contractor CSV; sends email to assigned contractors.
 */
function updateDVIProject(token, reportid, status, contractorIdsCsv, notes) {
  const session = getUserFromToken(token);
  if (!session.success) return session;
  const user = session.user;
  // Only allow admin/internal to assign jobs
  if (!user || (user.role.toLowerCase() !== "internal" && user.role.toLowerCase() !== "admin"))
    throw new Error("No permission.");
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const statusIdx = header.indexOf("status");
  const contractorIdx = header.indexOf("contractorid");
  const notesIdx = header.indexOf("reinotes");

  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] == reportid) {
      if (statusIdx >= 0) sheet.getRange(r + 1, statusIdx + 1).setValue(status);
      if (contractorIdx >= 0) sheet.getRange(r + 1, contractorIdx + 1).setValue(contractorIdsCsv);
      if (notesIdx >= 0) sheet.getRange(r + 1, notesIdx + 1).setValue(notes);

      // Send assignment email to contractors
      if (contractorIdsCsv) {
        const contractorIds = contractorIdsCsv.split(",").map(x => x.trim()).filter(Boolean);
        contractorIds.forEach(cid => {
          var contractor = getContractorById(cid); // Defined below
          if (contractor && contractor.email) {
            MailApp.sendEmail({
              to: contractor.email,
              subject: 'You have been assigned a DVI Project',
              htmlBody: 'A new DVI project has been assigned to you. Please log in to your portal to review and accept.<br>Project ID: ' + reportid
            });
          }
        });
      }
      break;
    }
  }
  return { success: true };
}

/**
 * Helper: Get contractor info by ID.
 * Looks up contractor from DVI Contractors sheet (by id or contractorid).
 */
function getContractorById(cid) {
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.CONTRACTORS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("contractorid") !== -1 ? header.indexOf("contractorid") : header.indexOf("id");
  const emailIdx = header.indexOf("email");
  const nameIdx = header.indexOf("name");
  for (let i = 1; i < data.length; i++) {
    if ((data[i][idIdx] || "").toString().trim() === cid) {
      return { id: cid, email: data[i][emailIdx], name: data[i][nameIdx] };
    }
  }
  return null;
}

/**
 * For legacy compatibility—single assignment.
 * Called by some admin panels.
 */
function updateDVIProjectLegacy(reportid, newStatus, contractorId, notes) {
  const ss = SpreadsheetApp.openById(config.DVI_PROJECTS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.DVI_PROJECTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idIdx = header.indexOf("reportid");
  const statusIdx = header.indexOf("status");
  const contractorIdx = header.indexOf("contractorid");
  const notesIdx = header.indexOf("reinotes");
  for (let i = 1; i < data.length; i++) {
    if ((data[i][idIdx] || "") === (reportid || "")) {
      if (newStatus !== undefined) sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      if (contractorId !== undefined) sheet.getRange(i + 1, contractorIdx + 1).setValue(contractorId);
      if (notes !== undefined) sheet.getRange(i + 1, notesIdx + 1).setValue(notes);
      break;
    }
  }
  return { success: true };
}

/**
 * Transfers data to the document
 * @param {string} reportId The report ID to transfer data to
 * @param {string} type The type of project to transfer data to
 */
function transferDataToDoc(reportId, type) {
  // Get Sheet and Doc IDs
  const [projectSheetId, docId] = getDocAndSheetId(reportId, type);

  // Open Project Sheet 
  SPREADSHEETS.PROJECT = SpreadsheetApp.openById(projectSheetId);

  // Get Sheets
  SHEETS.REPORT_INPUT = SPREADSHEETS.PROJECT.getSheetByName(config.SHEET_NAMES.REPORT_INPUT);
  SHEETS.PROPOSED = SPREADSHEETS.PROJECT.getSheetByName(config.SHEET_NAMES.PROPOSED);
  SHEETS.EXISTING = SPREADSHEETS.PROJECT.getSheetByName(config.SHEET_NAMES.EXISTING);
  SHEETS.SYSTEMS = SPREADSHEETS.PROJECT.getSheetByName(config.SHEET_NAMES.SYSTEMS);
  SHEETS.CHARTS = SPREADSHEETS.PROJECT.getSheetByName(config.SHEET_NAMES.CHARTS);

  if (!SHEETS.REPORT_INPUT) throw new Error('❌ "Report.Input" sheet not found.');
  if (!SHEETS.PROPOSED) throw new Error('❌ "Proposed" sheet not found.');
  if (!SHEETS.EXISTING) throw new Error('❌ "Existing" sheet not found.');
  if (!SHEETS.SYSTEMS) throw new Error('❌ "Systems" sheet not found.');

  // Get Data
  DATA.PROPOSED = new Table(SHEETS.PROPOSED.getDataRange().getDisplayValues());
  DATA.INPUT = new Table(SHEETS.REPORT_INPUT.getDataRange().getDisplayValues());
  DATA.SYSTEMS = new Table(SHEETS.SYSTEMS.getDataRange().getDisplayValues());
  DATA.EXISTING = new Table(SHEETS.EXISTING.getDataRange().getDisplayValues());

  const placeholderValues = {};
  const imageValues = {};

  // Build placeholder values from input sheet
  DATA.INPUT.json.forEach((row, idx) => {
    const keys = row['placeholder'].toString().split('|').map(value => value.trim());
    if (keys && keys.length > 0) {
      keys.forEach((key, index) => {
        if (!key.startsWith('<') && !key.endsWith('>')) return;
        if (!key.includes('systemphoto')) {
          placeholderValues[key] = row['value'].toString()?.split('|')[index]?.trim() || "";
        } else {
          const value = SHEETS.REPORT_INPUT.getRange(idx + 2, DATA.INPUT.colindex('value') + 1).getValue();
          if (value.valueType === SpreadsheetApp.ValueType.IMAGE) {
            const url = value.getContentUrl();
            const blob = UrlFetchApp.fetch(url).getBlob();
            imageValues[key] = blob;
          } else {
            placeholderValues[key] = "";
          }
        }
      });
    }
  });

  // Get proposed images
  getProposedImages(placeholderValues, imageValues);

  // Handle Forecast systems
  addForecastSystems(placeholderValues);

  // Special fields
  placeholderValues['<report.id>'] = reportId;
  placeholderValues['<e.waste.t>'] = calculateAverageWaste(DATA.EXISTING);

  // Handle address line 1 fallback
  if (!placeholderValues['<address.line1>'] && placeholderValues['<p.address.line1>']) {
    placeholderValues['<address.line1>'] = placeholderValues['<p.address.line1>'];
  }

  // System placeholders
  const numProposedSystems = extractSystemParameters(DATA.PROPOSED, "p", placeholderValues);

  // Total values for proposed systems
  getTotalValues(placeholderValues);

  // Process existing system data
  extractSystemParameters(DATA.EXISTING, "e", placeholderValues);

  // Open Document
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // Handle system template
  deleteUnusedSystemPages(body, placeholderValues);

  // Add conclusion numbering
  placeholderValues['<conclusion.number>'] = numProposedSystems + 3;

  // Charts
  Object.entries(config.CHARTS).forEach(([chartType, chartId]) => {
    Logger.log(`Processing ${chartType} chart with ID: ${chartId}`);
    const chartImage = exportChartAsImage(SHEETS.CHARTS, chartId);
    if (!chartImage) {
      Logger.log(`Failed to export ${chartType} chart`);
      return;
    }
    imageValues[chartType] = chartImage.blob;
  });

  // Replace placeholders in doc, header, footer
  replacePlaceholders(
    [body, doc.getHeader(), doc.getFooter()],
    placeholderValues,
    imageValues
  );
}

/**
 * Utility to get Doc/Sheet ID for a reportId and type
 * Returns [sheetId, docId]
 */
function getDocAndSheetId(reportId, type) {
  const cfg = config.getConfig(type);
  const dbSheet = SpreadsheetApp.openById(cfg.PROJECT_DATABASE_ID).getSheetByName('Projects');
  const data = dbSheet.getDataRange().getValues();
  const idx = data.findIndex(row => row[10] === reportId);
  if (idx === -1) throw new Error(`❌ Project not found: ${reportId}`);
  const sheetUrl = data[idx][0];
  const docUrl = data[idx][1];
  if (!docUrl || !sheetUrl) throw new Error(`❌ Missing Sheet or Doc URL for: ${reportId}`);
  const sheetId = sheetUrl.match(/[-\w]{25,}/)?.[0];
  const docId = docUrl.match(/[-\w]{25,}/)?.[0];
  return [sheetId, docId];
}

/**
 * Add forecast systems to the placeholder values.
 */
function addForecastSystems(placeholderValues) {
  for (let i = 1; i <= 4; i++) {
    placeholderValues[`<forcst.system${i}>`] = "N/A";
    placeholderValues[`<${i}.t.install.$>`] = "N/A";
    placeholderValues[`<${i}.incentive.savings>`] = "N/A";
    placeholderValues[`<${i}.est.annual.save>`] = "N/A";
    placeholderValues[`<${i}.roi>`] = "N/A";
  }
  let numForecastSystems = 0;
  for (let i = 1; i <= 20 && numForecastSystems < 4; i++) {
    const check = `<include.forecast${i}>`;
    if (!placeholderValues[check] || String(placeholderValues[check]).toLowerCase() === "false") continue;
    const name = placeholderValues[`<systempage${i}>`];
    const system = DATA.SYSTEMS.json.find(row =>
      row['modelNumber'] && row['modelNumber'].toString().trim() === placeholderValues[`<p.model#${i}>`]
    );
    const totalCost = (parseFloat(system?.['installEach']?.replace(/[\$\,]/g, "") || 0) +
                      parseFloat(system?.['unitCost']?.replace(/[\$\,]/g, "") || 0));
    const proposedRow = DATA.PROPOSED.json.find(row =>
      row['modelNumber'] && row['modelNumber'].toString().trim() === placeholderValues[`<p.model#${i}>`]
    );
    const incentiveSavings = parseFloat(proposedRow?.['incentiveSavings']?.replace(/[\$\,]/g, "") || 0);
    const annualSavings   = parseFloat(proposedRow?.['annualSavings']?.replace(/[\$\,]/g, "") || 0);
    const roi             = parseFloat(proposedRow?.['roi']?.replace("%", "") || 0);

    placeholderValues[`<forcst.system${numForecastSystems + 1}>`]   = name || "";
    placeholderValues[`<${numForecastSystems + 1}.t.install.$>`]    = formatCurrency(totalCost);
    placeholderValues[`<${numForecastSystems + 1}.incentive.savings>`] = formatCurrency(incentiveSavings);
    placeholderValues[`<${numForecastSystems + 1}.est.annual.save>`]   = formatCurrency(annualSavings);
    placeholderValues[`<${numForecastSystems + 1}.roi>`]            = formatPercent(roi);

    numForecastSystems++;
  }
}

/**
 * Extracts system parameters from sheet data, populates placeholderValues, returns system count.
 */
function extractSystemParameters(data, letter, placeholderValues) {
  let numSystems = 0;
  placeholderValues["<u.waste.t>"] = calculateAverageWaste(DATA.PROPOSED);
  for (let sysNum = 1; sysNum <= data.json.length; sysNum++) {
    const modelPlaceholder = `<${letter}.model#${sysNum}>`;
    const modelNumber = placeholderValues[modelPlaceholder] || "";
    if (modelNumber) {
      const targetRow = data.json.find(row => row['modelNumber'] && row['modelNumber'].toString().trim() === modelNumber);
      if (targetRow) {
        const percentOfTotalBill = parseFloat(targetRow['totalBill']) || 0;
        const carbonFootprint    = parseFloat(targetRow['carbonFootprint']) || 0;
        const wasteDecimal       = parseFloat(targetRow['waste']) || 0;
        const efficiencyDecimal  = parseFloat(targetRow['efficiency']) || 0;
        const costToUseSum = [
          targetRow['electricity'],
          targetRow['naturalGas'],
          targetRow['oil'],
          targetRow['propane'],
          targetRow['water'],
          targetRow['sewer']
        ].reduce((sum, col) => sum + (parseFloat((col||"").replace(/[\$\,]/g, "")) || 0), 0);

        placeholderValues[`<${letter}.cost.to.use${sysNum}>`]   = formatCurrency(costToUseSum);
        placeholderValues[`<${letter}.%oftotalbill${sysNum}>`]  = formatPercent(percentOfTotalBill);
        placeholderValues[`<${letter}.carbonfootprint${sysNum}>`] = `${formatCurrency(carbonFootprint, 0, '')} lbs CO₂`;
        placeholderValues[`<${letter}.waste%${sysNum}>`]        = formatPercent(wasteDecimal);
        placeholderValues[`<${letter}.efficiency${sysNum}>`]    = formatPercent(efficiencyDecimal);

        if (letter === 'p') placeholderValues[`<delta.cost.to.use${sysNum}>`] = targetRow['annualSavings'];
        numSystems++;
      }
      if (letter === 'p') {
        const matchRow = DATA.SYSTEMS.json.find(row => row['modelNumber'] && row['modelNumber'].toString().trim() === modelNumber);
        placeholderValues[`<${letter}.unitcost${sysNum}>`] = matchRow?.['unitCost']?.toString() || "";
      }
    }
  }
  return numSystems;
}

/**
 * Get proposed system images from sheet, populate imageValues and placeholderValues.
 */
function getProposedImages(placeholderValues, imageValues) {
  for (let i = 1; i <= 20; i++) {
    const modelPlaceholder = `<p.model#${i}>`;
    const modelNumber = placeholderValues[modelPlaceholder] || '';
    if (modelNumber) {
      const rowindex = DATA.SYSTEMS.json.findIndex(row => row['modelNumber'] && row['modelNumber'].toString().trim() === modelNumber);
      if (rowindex >= 0) {
        const cell = SHEETS.SYSTEMS.getRange(rowindex + 2, DATA.SYSTEMS.colindex('systemPhoto') + 1);
        const value = cell.getValue();
        if (value && value.valueType === SpreadsheetApp.ValueType.IMAGE) {
          const blob = UrlFetchApp.fetch(value.getContentUrl()).getBlob();
          imageValues[`<p.systemphoto${i}>`] = blob;
        } else {
          placeholderValues[`<p.systemphoto${i}>`] = "";
        }
      }
    }
  }
}

/**
 * Calculate and populate utility total values (electric, ngas, water, etc).
 */
function getTotalValues(placeholderValues) {
  let totalRow = DATA.PROPOSED.json.find(row => row['systemNo'] && row['systemNo'].trim() == "Total Cost per utility");
  placeholderValues['<u.ele.$>'] = totalRow?.['electricity'];
  placeholderValues['<u.ngas.$>'] = totalRow?.['naturalGas'];
  placeholderValues['<u.water.$>'] = totalRow?.['water'];
  placeholderValues['<u.cost.t>'] = formatCurrency(
    ['electricity', 'naturalGas', 'oil', 'propane', 'water', 'sewer']
      .reduce((sum, utility) => sum + (totalRow?.[utility] ? parseFloat((totalRow[utility]||"").replace(/[\$\,]/g, "")) : 0), 0)
    , 0
  );
  totalRow = DATA.PROPOSED.json.find(row => row['systemNo'] && row['systemNo'].trim() == "Total Utility consumed");
  placeholderValues['<u.ele.eg>'] = totalRow?.['electricity'];
  placeholderValues['<u.ngas.eg>'] = totalRow?.['naturalGas'];
  placeholderValues['<u.water.eg>'] = totalRow?.['water'];
  totalRow = DATA.PROPOSED.json.find(row => row['systemNo'] && row['systemNo'].trim() == "Total Carbon footprint per utility");
  placeholderValues['<u.co2.t>'] = formatCurrency(
    ['electricity', 'naturalGas', 'oil', 'propane', 'water', 'sewer']
      .reduce((sum, utility) => sum + (totalRow?.[utility] ? parseFloat(totalRow[utility].split(" ")[0]) : 0), 0)
    , 0
  ) + " lbs CO₂";
}

/**
 * Escapes special regex characters in keys for placeholder replacement.
 */
function escapeRegExp(string) {
  return string.replace(/[.*+\-?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Format currency with commas and dollar sign.
 */
function formatCurrency(num, decimals=0, prefix='$') {
  if (num === undefined || num === null || isNaN(num)) return '';
  return prefix + parseFloat(num).toLocaleString('en-US', {minimumFractionDigits: decimals, maximumFractionDigits: decimals});
}

/**
 * Format percent with percent sign.
 */
function formatPercent(num, decimals=1) {
  if (num === undefined || num === null || isNaN(num)) return '';
  return parseFloat(num).toFixed(decimals) + '%';
}

/**
 * Logging helper (pretty-print objects)
 */
function superLog(title, obj) {
  Logger.log(`=== ${title} ===`);
  Logger.log(JSON.stringify(obj, null, 2));
  Logger.log('================');
}

/**
 * Chart export utility (from Google Sheet to blob for Doc)
 */
function exportChartAsImage(sheet, chartId) {
  try {
    const numericId = typeof chartId === 'string' ? parseFloat(chartId) : chartId;
    const charts = sheet.getCharts();
    Logger.log(`Found ${charts.length} charts in sheet`);
    const chart = charts.find(chart => chart.getChartId() === numericId);
    if (!chart) throw new Error(`Chart with ID ${numericId} not found. Available: ${config.charts.map(c => c.getChartId()).join(', ')}`);
    const blob = chart.getBlob();
    return {
      blob: blob,
      title: chart.getOptions().get('title') || 'Chart',
      width: chart.getOptions().get('width') || 600,
      height: chart.getOptions().get('height') || 400
    };
  } catch (e) {
    Logger.log(`Error exporting chart ${chartId}: ${e.message}`);
    return null;
  }
}

/**
 * Calculate average waste % for a Table
 */
function calculateAverageWaste(data) {
  const wasteValues = data.json.map(row => {
    if (row['modelNumber'] === '') return;
    return parseFloat(row['waste'].replace("%", ""));
  }).filter(val => !isNaN(val));
  if (!wasteValues.length) return '';
  const average = Math.round((wasteValues.reduce((a, b) => a + b, 0) / wasteValues.length));
  return `${average}%`;
}

/**
 * TEST FUNCTION (dev only)
 */
function test() {
  var reportId = 'LEAF-25.002.11843-123';
  var type = 'LEAF';
  undoTransferToDoc(reportId, type);
  transferDataToDoc(reportId, type);
}

/**********************************************/
/* SECTION 8: FORMATTING HELPERS              */
/**********************************************/

function formatCurrency(num, decimals = 0, prefix = '$') {
  if (num === undefined || num === null || isNaN(num)) return '';
  return prefix + parseFloat(num).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function formatPercent(num, decimals = 1) {
  if (num === undefined || num === null || isNaN(num)) return '';
  return parseFloat(num).toFixed(decimals) + '%';
}

function superLog(title, obj) {
  Logger.log(`=== ${title} ===`);
  Logger.log(JSON.stringify(obj, null, 2));
  Logger.log('================');
}

/**********************************************/
/* SECTION 9: CHART EXPORT                    */
/**********************************************/

function exportChartAsImage(sheet, chartId) {
  try {
    const numericId = typeof chartId === 'string' ? parseFloat(chartId) : chartId;
    const charts = sheet.getCharts();
    Logger.log(`Found ${config.charts.length} charts in sheet`);
    const chart = charts.find(chart => chart.getChartId() === numericId);
    if (!chart) throw new Error(`Chart with ID ${numericId} not found. Available: ${config.charts.map(c => c.getChartId()).join(', ')}`);
    const blob = chart.getBlob();
    return {
      blob: blob,
      title: chart.getOptions().get('title') || 'Chart',
      width: chart.getOptions().get('width') || 600,
      height: chart.getOptions().get('height') || 400
    };
  } catch (e) {
    Logger.log(`Error exporting chart ${chartId}: ${e.message}`);
    return null;
  }
}

/**********************************************/
/* SECTION 10: UNDO DOC TRANSFER              */
/**********************************************/

function undoTransferToDoc(reportId, type) {
  const cfg = config.getConfig(type);
  const dbSheet = SpreadsheetApp.openById(cfg.PROJECT_DATABASE_ID).getSheetByName('Projects');
  const data = dbSheet.getDataRange().getValues();
  const project = data.find(row => row[10] === reportId);
  if (!project) throw new Error(`❌ Project not found: ${reportId}`);
  const docUrl = project[1];
  if (!docUrl) throw new Error(`❌ Missing Doc URL for: ${reportId}`);
  const docId = docUrl.match(/[-\w]{25,}/)?.[0];
  const docFile = DriveApp.getFileById(docId);
  const parentFolder = docFile.getParents().next();
  const fileName = docFile.getName();
  docFile.setTrashed(true);
  const newDocFile = DriveApp.getFileById(cfg.DOC_TEMPLATE_ID).makeCopy(fileName, parentFolder);
  const newDocUrl = newDocFile.getUrl();
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === reportId) {
      dbSheet.getRange(i + 1, 2).setValue(newDocUrl); // Update Doc Link
      break;
    }
  }
}

/**********************************************/
/* SECTION 12: TEST FUNCTIONS                 */
/**********************************************/

function test() {
  var reportId = 'LEAF-25.002.11843-123';
  var type = 'LEAF';
  undoTransferToDoc(reportId, type);
  transferDataToDoc(reportId, type);
}

/**********************************************/
/* SECTION 13: LEGACY & BACKWARDS-COMPAT HELPERS */
/**********************************************/

/**
 * Escapes regex special characters in a string.
 * Useful for safe placeholder replacement in documents.
 */
function escapeRegExp(string) {
  return String(string).replace(/[.*+\-?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Finds the nearest parent block element (PARAGRAPH or TABLE_CELL) in a Google Doc.
 * Used for advanced placeholder replacement or page removal.
 */
function findParentBlock(element) {
  while (
    element &&
    element.getType &&
    element.getType() !== DocumentApp.ElementType.PARAGRAPH &&
    element.getType() !== DocumentApp.ElementType.TABLE_CELL
  ) {
    element = element.getParent();
  }
  return element;
}

/**
 * Retrieves a unique list of all system names from the Report.Input sheet.
 * Used for system list previews or DVI forms.
 */
function getSystemListFromInputSheet(inputSheet) {
  // Assumes systems are in E12:E50 (col 5)
  var sysVals = inputSheet.getRange(12, 5, 39, 1).getValues(); // E12:E50
  var systemsList = [];
  sysVals.forEach(function (r) {
    var val = (r[0] || "").trim();
    if (val && val.toLowerCase() !== "placeholder" && systemsList.indexOf(val) === -1) {
      systemsList.push(val);
    }
  });
  return systemsList;
}

/**
 * Helper: Checks if a value is a valid image cell (from IMAGE formula or insert).
 * Used by getProposedImages or other image insert logic.
 */
function isGoogleSheetImageCell(cellValue) {
  // Google Sheets stores images as objects with valueType property (only in some contexts)
  return cellValue && typeof cellValue === 'object' && cellValue.valueType === SpreadsheetApp.ValueType.IMAGE;
}

/**
 * Returns the first parent TableCell of an element (if inside a table).
 */
function getParentTableCell(element) {
  while (element && element.getType && element.getType() !== DocumentApp.ElementType.TABLE_CELL) {
    element = element.getParent();
  }
  return element && element.getType && element.getType() === DocumentApp.ElementType.TABLE_CELL ? element : null;
}

/**
 * Helper: Get system model numbers from DATA.SYSTEMS for a quick lookup.
 */
function getModelNumbersFromSystems(dataSystemsTable) {
  if (!dataSystemsTable || !dataSystemsTable.json) return [];
  return dataSystemsTable.json
    .map(function (row) { return row['modelNumber'] && String(row['modelNumber']).trim(); })
    .filter(function (val, idx, arr) { return val && arr.indexOf(val) === idx; });
}

function getAllLeafProjects() {
  var sheet = SpreadsheetApp.openById('12yJhaOAe4rHSCFSadanh1K1YCu8wJnjbxfBdeX_GeB8').getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) {
    var reportId = data[i][10]; // This is the correct ReportID for the row!
    projects.push({
      SheetsLink:       data[i][0],
      DocsLink:         data[i][1],
      Status:           data[i][2],
      Created:          data[i][3],
      Company:          data[i][4],
      Property:         data[i][5],
      Address:          data[i][6],
      City:             data[i][7],
      State:            data[i][8],
      ZIP:              data[i][9],
      ReportID:         reportId,
      Note:             data[i][11],
      QuestionnaireURL: data[i][12] || `https://script.google.com/macros/s/AKfycbxWY-K0nOkgN_C3OH6Wby8DA3NBhXLgxyQ9gMSe90FegnJZRJJqE3derYNnL1h-ma0/exec?reportId=${encodeURIComponent(reportId)}`,
      DVIURL:           data[i][13] || `https://script.google.com/macros/s/AKfycbyOBD3TmfJrQyAsqlKvlOBk3-9cwOnWcOySCkgIaipFkXzibz9Eq1VV1Q7IhoIUbBU/exec?reportid=${encodeURIComponent(reportId)}`
    });
  }
  return projects;
}

function getLeafProjectsHtml() {
  var type = 'leaf';  // <-- Add this line
  var projects = getAllLeafProjects();
  var STATUS_OPTIONS = ['Pay.Received', 'Cust.Input', 'Report.Sent', 'DVI.Requested'];
  var html = '<div class="table-responsive"><table class="project-table"><thead><tr>';
  html += '<th>SheetsLink</th><th>DocsLink</th><th>Status</th><th>Created</th><th>Company</th><th>Property</th><th>Address</th><th>City</th><th>State</th><th>ZIP</th><th>ReportID</th><th>Note</th><th>Actions</th>';
  html += '</tr></thead><tbody>';
  projects.forEach(function(proj, idx) {
    html += `<tr>
      <td><a href="${proj.SheetsLink}" target="_blank">Open Sheet</a></td>
      <td><a href="${proj.DocsLink}" target="_blank">Open Doc</a></td>
      <td>
        <select class="status-dropdown" data-type="leaf" data-reportid="${proj.ReportID}" onchange="handleStatusChange(this)">
          ${STATUS_OPTIONS.map(opt => `<option value="${opt}"${proj.Status === opt ? ' selected' : ''}>${opt}</option>`).join('')}
        </select>
      </td>
      <td>${proj.Created || ''}</td>
      <td>${proj.Company || ''}</td>
      <td>${proj.Property || ''}</td>
      <td>${proj.Address || ''}</td>
      <td>${proj.City || ''}</td>
      <td>${proj.State || ''}</td>
      <td>${proj.ZIP || ''}</td>
      <td>${proj.ReportID || ''}</td>
      <td>
        <input type="text" 
          value="${proj.Note ? proj.Note.replace(/"/g,'&quot;') : ''}" 
          data-type="leaf" 
          data-reportid="${proj.ReportID}" 
          class="notes-input" 
          style="width:120px;" />
      </td>
      <td>
        <select class="action-dropdown" onchange="handleProjectAction(this, '${proj.ReportID}', '${type}')">
            <option value="">Actions…</option>
            <option value="delete">🗑 Delete</option>
            <option value="transfer">⬇️ Transfer.Data</option>
            <option value="redo">🔄 Redo.Transfer</option>
            <option value="pdf">📄 View PDF</option>
            <option value="questionnaire">📝 Questionnaire</option>
            <option value="dvi">🔗 DVI.LINK</option>
          </select>
        </td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

function getAllRootProjects() {
  var sheet = SpreadsheetApp.openById('193m8yOy51aDwSvqvCQQv7uM-fpiJrrnNK2kqCGYLq7E').getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) {
    var reportId = data[i][10]; // This is the correct ReportID for the row!
    projects.push({
      SheetsLink:       data[i][0],
      DocsLink:         data[i][1],
      Status:           data[i][2],
      Created:          data[i][3],
      Company:          data[i][4],
      Property:         data[i][5],
      Address:          data[i][6],
      City:             data[i][7],
      State:            data[i][8],
      ZIP:              data[i][9],
      ReportID:         reportId,
      Note:             data[i][11],
      QuestionnaireURL: data[i][12] || `https://script.google.com/macros/s/AKfycbxWY-K0nOkgN_C3OH6Wby8DA3NBhXLgxyQ9gMSe90FegnJZRJJqE3derYNnL1h-ma0/exec?reportId=${encodeURIComponent(reportId)}`,
      DVIURL:           data[i][13] || `https://script.google.com/macros/s/AKfycbyOBD3TmfJrQyAsqlKvlOBk3-9cwOnWcOySCkgIaipFkXzibz9Eq1VV1Q7IhoIUbBU/exec?reportid=${encodeURIComponent(reportId)}`
    });
  }
  return projects;
}

function getRootProjectsHtml() {
  var type = 'root';  // <-- Add this line
  var projects = getAllRootProjects();
  var STATUS_OPTIONS = ['Pay.Received', 'Cust.Input', 'Report.Sent', 'DVI.Requested'];
  var html = '<div class="table-responsive"><table class="project-table"><thead><tr>';
  html += '<th>SheetsLink</th><th>DocsLink</th><th>Status</th><th>Created</th><th>Company</th><th>Property</th><th>Address</th><th>City</th><th>State</th><th>ZIP</th><th>ReportID</th><th>Note</th><th>Actions</th>';
  html += '</tr></thead><tbody>';
  projects.forEach(function(proj, idx) {
    html += `<tr>
      <td><a href="${proj.SheetsLink}" target="_blank">Open Sheet</a></td>
      <td><a href="${proj.DocsLink}" target="_blank">Open Doc</a></td>
      <td>
        <select class="status-dropdown" data-type="root" data-reportid="${proj.ReportID}" onchange="handleStatusChange(this)">
          ${STATUS_OPTIONS.map(opt => `<option value="${opt}"${proj.Status === opt ? ' selected' : ''}>${opt}</option>`).join('')}
        </select>
      </td>
      <td>${proj.Created || ''}</td>
      <td>${proj.Company || ''}</td>
      <td>${proj.Property || ''}</td>
      <td>${proj.Address || ''}</td>
      <td>${proj.City || ''}</td>
      <td>${proj.State || ''}</td>
      <td>${proj.ZIP || ''}</td>
      <td>${proj.ReportID || ''}</td>
      <td>
        <input type="text" 
          value="${proj.Note ? proj.Note.replace(/"/g,'&quot;') : ''}" 
          data-type="root" 
          data-reportid="${proj.ReportID}" 
          class="notes-input" 
          style="width:120px;" />
      </td>
      <td>
        <select class="action-dropdown" onchange="handleProjectAction(this, '${proj.ReportID}', '${type}')">
            <option value="">Actions…</option>
            <option value="delete">🗑 Delete</option>
            <option value="transfer">⬇️ Transfer.Data</option>
            <option value="redo">🔄 Redo.Transfer</option>
            <option value="pdf">📄 View PDF</option>
            <option value="questionnaire">📝 Questionnaire</option>
            <option value="dvi">🔗 DVI.LINK</option>
          </select>
        </td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}
function adminGetAllProjects(token) {
  if (!validateTokenIsAdmin(token)) return [];
  const sheet = SpreadsheetApp.openById(PROJECTS_SHEET_ID).getSheetByName('Projects');
  const data = sheet.getDataRange().getValues();
  // [ID, Name, Status, ContractorId, ...]
  return data.slice(1).map(row => ({
    id: row[0],
    name: row[1],
    status: row[2],
    contractorId: row[3]
  }));
}
// --- Add this to Code.js ---
function validateTokenIsAdmin(token) {
  if (!token) return false;
  // TODO: Replace this logic with your real token/admin validation.
  // For example, compare to a list or check with PropertiesService.
  var adminTokens = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKENS');
  if (!adminTokens) return false;
  adminTokens = JSON.parse(adminTokens); // Should be an array of valid tokens
  return adminTokens.indexOf(token) !== -1;
}
/**
 * Return all LEAF/ROOT projects assigned to a specific contractor by ContractorID.
 * @param {string} contractorId - Contractor's unique ID (should come from the session)
 * @param {string} type - "LEAF" or "ROOT"
 * @return {Array<Object>} Array of project objects assigned to that contractor.
 */
function getContractorProjects(contractorId, type) {
  var cfg = config.PROJECT_SOURCES[type || 'LEAF'];
  var sheet = SpreadsheetApp.openById(cfg.PROJECT_DATABASE_ID).getSheetByName('Projects');
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var contractorColIdx = header.indexOf('ContractorID'); // make sure column exists!

  if (contractorColIdx < 0) throw new Error('ContractorID column not found in Projects sheet!');

  var projects = [];
  for (var i = 1; i < data.length; i++) {
    if ((data[i][contractorColIdx] + "") === (contractorId + "")) {
      projects.push({
        SheetsLink:       data[i][0],
        DocsLink:         data[i][1],
        Status:           data[i][2],
        Created:          data[i][3],
        Company:          data[i][4],
        Property:         data[i][5],
        Address:          data[i][6],
        City:             data[i][7],
        State:            data[i][8],
        ZIP:              data[i][9],
        ReportID:         data[i][10],
        Note:             data[i][11],
        QuestionnaireURL: data[i][12],
        DVIURL:           data[i][13]
      });
    }
  }
  return projects;
}
function addNewContractor(sessionToken, contractorObject) {
  const session = getUserFromToken(sessionToken);
  if (!session.success) throw new Error("Invalid session.");
  const user = session.user;
  if (user.role !== "Internal") throw new Error("No permission.");

  const ss = SpreadsheetApp.openById(config.CONTRACTORS_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(config.CONTRACTORS_SHEET_NAME);
  const headers = sheet.getDataRange().getValues()[0];

  const now = new Date();
  const row = headers.map(header => {
    if (header === 'DateAdded') return now;
    if (header === 'Status') return 'Active';
    return contractorObject[header] || "";
  });

  sheet.appendRow(row);
}
