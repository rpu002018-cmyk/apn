/***********************
 * CONFIGURATION
 ***********************/
const SHEET_ID = "sheet id"; // <-- set your real Spreadsheet ID here
const DATE_FORMAT = "dd-MM-yyyy";
const TZ = Session.getScriptTimeZone();

const SHEETS = {
  STUDENTS: "Students",
  DOCUMENTS: "Documents",
  DOCUMENTS_ISSUED: "Documents Issued",
  OPERATORS: "Operators",
  EMAIL_LOG: "Email Log"
};

/***********************
 * COMMON
 ***********************/
function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function toBool(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (v == null) return false;
  const s = String(v).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1";
}

/***********************
 * WEB APP ENTRY
 ***********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Admission Document Acknowledgement");
}

/***********************
 * OPERATOR LOGIN
 ***********************/
function validateOperator(id, pwd) {
  const sh = getSS().getSheetByName(SHEETS.OPERATORS);
  const data = sh.getDataRange().getValues();

  const uid = String(id || "").trim();
  const upwd = String(pwd || "").trim();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || "").trim();
    const rowPwd = String(data[i][1] || "").trim();
    if (rowId === uid && rowPwd === upwd) {
      return { success: true, name: String(data[i][2] || "").trim() };
    }
  }
  return { success: false };
}

/***********************
 * STUDENT CRUD
 ***********************/
function addStudent(student) {
  // student: { registerNo, applicationNo, name, section, subSection, email }
  const sh = getSS().getSheetByName(SHEETS.STUDENTS);
  if (!sh) throw new Error("Students sheet not found");
  // Basic validation
  if (!student || !student.registerNo || !student.applicationNo || !student.name) {
    return { success: false, error: "Register No, Application No and Name are required" };
  }

  // Prevent duplicates by applicationNo or registerNo
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(student.registerNo).trim() ||
        String(data[i][1]).trim() === String(student.applicationNo).trim()) {
      return { success: false, error: "Student with same Register No or Application No already exists" };
    }
  }

  sh.appendRow([
    student.registerNo,
    student.applicationNo,
    student.name,
    student.section || "",
    student.subSection || "",
    student.email || ""
  ]);

  return { success: true };
}

/***********************
 * STUDENT SEARCH
 ***********************/
function searchStudent(value) {
  const sh = getSS().getSheetByName(SHEETS.STUDENTS);
  const data = sh.getDataRange().getValues();

  const v = String(value || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0] || "").trim() === v || // Register No
      String(data[i][1] || "").trim() === v    // Application No
    ) {
      return {
        registerNo: data[i][0],
        applicationNo: data[i][1],
        name: data[i][2],
        section: data[i][3],
        subSection: data[i][4],
        email: data[i][5]
      };
    }
  }
  return null;
}

/***********************
 * LOAD EXISTING DOCUMENTS (RECEIVED MODE)
 ***********************/
function loadExistingDocuments(applicationNo) {
  const sh = getSS().getSheetByName(SHEETS.DOCUMENTS);
  const data = sh.getDataRange().getValues();

  return data
    .filter((r, i) => i > 0 && String(r[0]) === String(applicationNo))
    .map(r => {
      const submittedDate = r[6] ? Utilities.formatDate(new Date(r[6]), TZ, DATE_FORMAT) : "";
      return {
        name: r[2],
        original: toBool(r[3]),
        xerox: toBool(r[4]),
        status: r[5] || ( (toBool(r[3]) || toBool(r[4])) ? "Submitted" : "Pending" ),
        submittedDate: submittedDate,
        submittedBy: r[7] || "",
        receivedBy: r[8] || ""
      };
    });
}

/***********************
 * SAVE DOCUMENTS - RECEIVED MODE (FINAL LOGIC)
 * - Do NOT overwrite receivedBy/submittedBy/submittedDate if document was already submitted.
 * - Only set receivedBy when there is a NEW submission for that document (i.e., a flag changes from false to true).
 ***********************/
function saveDocuments(payload) {
  try {
    const sh = getSS().getSheetByName(SHEETS.DOCUMENTS);
    const data = sh.getDataRange().getValues();
    const now = new Date();

    payload.documents.forEach(doc => {
      // find existing row for this application+doc name
      let rowIndex = -1;
      let existing = null;
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][0]).trim() === String(payload.applicationNo).trim() &&
          String(data[i][2]).trim() === String(doc.name).trim()
        ) {
          rowIndex = i + 1; // sheet row number
          existing = data[i];
          break;
        }
      }

      const existingOriginal = existing ? toBool(existing[3]) : false;
      const existingXerox = existing ? toBool(existing[4]) : false;

      // Determine if this payload is trying to set original/xerox
      const wantsOriginal = Boolean(doc.original);
      const wantsXerox = Boolean(doc.xerox);

      // Determine if there is a NEW submission being added now
      const newOriginalSubmitted = wantsOriginal && !existingOriginal;
      const newXeroxSubmitted = wantsXerox && !existingXerox;
      const isNewSubmission = newOriginalSubmitted || newXeroxSubmitted;

      // If document already fully submitted (both flags true) and nothing new => skip (preserve)
      if (existing && existingOriginal && existingXerox && !isNewSubmission) {
        return;
      }

      // Final flags should be union of existing and new
      const finalOriginal = existingOriginal || wantsOriginal;
      const finalXerox = existingXerox || wantsXerox;

      // Decide submitted date / submittedBy / receivedBy:
      // - If there is an existing submitted date and no new submission, preserve it.
      // - If this call includes a new submission, set submittedDate = now and receivedBy = payload.operatorName
      let submitDateValue = "";
      let submittedByValue = "";
      let receivedByValue = "";

      if (existing && existing[6]) {
        // existing date present
        submitDateValue = existing[6];
      }

      if (existing && existing[7]) {
        submittedByValue = existing[7];
      }

      if (existing && existing[8]) {
        receivedByValue = existing[8];
      }

      if (isNewSubmission) {
        submitDateValue = now;
        // prefer document-level submittedBy if passed, else payload default
        submittedByValue = doc.submittedBy || submittedByValue || "";
        // Set receivedBy only for the first time this doc is submitted.
        // If already had receivedBy, preserve it; only set if empty.
        if (!receivedByValue) {
          receivedByValue = payload.operatorName || "";
        } else {
          // If receivedBy exists, do not overwrite (user requirement)
          // keep existing receivedByValue
        }
      }

      const status = (finalOriginal || finalXerox) ? "Submitted" : "Pending";

      const row = [
        payload.applicationNo,             // A
        payload.studentName,               // B
        doc.name,                          // C
        finalOriginal,                     // D Original
        finalXerox,                        // E Xerox
        status,                            // F Status
        submitDateValue,                   // G Submitted Date (Date or "")
        submittedByValue,                  // H Submitted By
        receivedByValue,                   // I Received By
        now,                               // J Last Updated
        payload.operatorId                 // K Updated By
      ];

      if (rowIndex > 0) {
        sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }
    });

    // EMAIL ACKNOWLEDGEMENT
    if (payload.email) {
      sendAcknowledgementEmail(payload);
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/***********************
 * LOAD EXISTING DOCUMENTS - ISSUED MODE
 ***********************/
function loadIssuedDocuments(applicationNo) {
  const sh = getSS().getSheetByName(SHEETS.DOCUMENTS_ISSUED);
  if (!sh) return [];
  
  const data = sh.getDataRange().getValues();

  return data
    .filter((r, i) => i > 0 && String(r[0]) === String(applicationNo))
    .map(r => {
      const issuedDate = r[4] ? Utilities.formatDate(new Date(r[4]), TZ, DATE_FORMAT) : "";
      return {
        name: r[2],
        quantity: r[3] || 1,
        issuedDate: issuedDate,
        issuedBy: r[5] || "",
        receivedBy: r[6] || "",
        remarks: r[7] || "",
        status: r[8] || "Issued"
      };
    });
}

/***********************
 * SAVE DOCUMENTS - ISSUED MODE
 * - Tracks documents issued TO students (certificates, admit cards, etc.)
 * - Documents can be issued multiple times with quantity tracking
 * - Preserves issuedBy and receivedBy information
 ***********************/
function saveIssuedDocuments(payload) {
  try {
    let sh = getSS().getSheetByName(SHEETS.DOCUMENTS_ISSUED);
    
    // Create sheet if it doesn't exist
    if (!sh) {
      sh = getSS().insertSheet(SHEETS.DOCUMENTS_ISSUED);
      // Add headers
      sh.appendRow([
        "Application No",      // A
        "Student Name",        // B
        "Document Name",       // C
        "Quantity",            // D
        "Issued Date",         // E
        "Issued By",           // F
        "Received By",         // G
        "Remarks",             // H
        "Status",              // I
        "Last Updated",        // J
        "Updated By"           // K
      ]);
    }

    const data = sh.getDataRange().getValues();
    const now = new Date();

    payload.documents.forEach(doc => {
      // find existing row for this application+doc name
      let rowIndex = -1;
      let existing = null;
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][0]).trim() === String(payload.applicationNo).trim() &&
          String(data[i][2]).trim() === String(doc.name).trim()
        ) {
          rowIndex = i + 1; // sheet row number
          existing = data[i];
          break;
        }
      }

      // Get existing values if document exists
      let issuedDateValue = now;
      let issuedByValue = doc.issuedBy || payload.operatorName || "";
      let receivedByValue = doc.receivedBy || "";
      let remarksValue = doc.remarks || "";
      let quantityValue = doc.quantity || 1;
      let statusValue = doc.status || "Issued";

      if (existing) {
        // Preserve existing issued date, only use current date if first time
        issuedDateValue = existing[4] || now;
        // Preserve existing issuedBy
        issuedByValue = existing[5] || issuedByValue;
        // Update receivedBy if provided
        receivedByValue = doc.receivedBy || existing[6] || "";
        // Update remarks
        remarksValue = doc.remarks || existing[7] || "";
        // Update quantity (sum if multiple issues)
        quantityValue = doc.quantity ? (parseInt(existing[3] || 0) + parseInt(doc.quantity)) : (existing[3] || 1);
        // Update status
        statusValue = doc.status || existing[8] || "Issued";
      }

      const row = [
        payload.applicationNo,             // A
        payload.studentName,               // B
        doc.name,                          // C
        quantityValue,                     // D Quantity
        issuedDateValue,                   // E Issued Date
        issuedByValue,                     // F Issued By
        receivedByValue,                   // G Received By
        remarksValue,                      // H Remarks
        statusValue,                       // I Status
        now,                               // J Last Updated
        payload.operatorId                 // K Updated By
      ];

      if (rowIndex > 0) {
        sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }
    });

    // EMAIL ACKNOWLEDGEMENT FOR ISSUED DOCUMENTS
    if (payload.email) {
      sendIssuanceAcknowledgementEmail(payload);
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/***********************
 * EMAIL ACKNOWLEDGEMENT - RECEIVED MODE
 ***********************/
function sendAcknowledgementEmail(payload) {
  if (!payload.email) return;

  const subject = `Admission Document Acknowledgement - ${payload.applicationNo}`;

  let table = `
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; width:100%;">
      <tr>
        <th>SL.NO</th>
        <th>Document</th>
        <th>Original</th>
        <th>Xerox</th>
        <th>Submitted By</th>
        <th>Date</th>
        <th>Received By</th>
      </tr>`;

  let sl = 1;
  payload.documents.forEach(d => {
    if (d.original || d.xerox) {
      // Prefer any provided submittedDate, otherwise use current date
      const dateStr = d.submittedDate
        ? d.submittedDate
        : Utilities.formatDate(new Date(), TZ, DATE_FORMAT);

      table += `
          <tr>
            <td>${sl++}</td>
            <td>${d.name}</td>
            <td style="text-align:center">${d.original ? "✔" : ""}</td>
            <td style="text-align:center">${d.xerox ? "✔" : ""}</td>
            <td>${d.submittedBy || ""}</td>
            <td>${dateStr}</td>
            <td>${payload.operatorName}</td>
          </tr>`;
    }
  });

  table += `</table>`;

  const body = `
      <p>Dear Student / Parent,</p>

<p>
  This is to acknowledge the receipt of the documents submitted by the student.
  The submitted documents have been reviewed and verified as per college records.
</p>

<p><strong>Student Details:</strong></p>
<p>
  Name: ${payload.studentName}<br>
  Application No: ${payload.applicationNo}<br>
  Section: ${payload.section || "-"}<br>
  Sub-Section: ${payload.subSection || "-"}
</p>

${table}

<p>
  Please note that this acknowledgement is issued for record purposes only.
</p>

<p>
  Yours faithfully,<br>
  Office<br>
  <strong>REVA Independent PU College</strong><br>
  Rukmini Knowledge Park,<br>
  Kattigenahalli, Yelahanka,<br>
  Bangalore – 560064.<br><br>

  Phone: 080-46966966 (Ext: 184)<br>
  Email: <a href="mailto:ripu.k@reva.edu.in">ripu.k@reva.edu.in</a>
</p>

    `;

  MailApp.sendEmail({
    to: payload.email,
    subject: subject,
    htmlBody: body
  });

  logEmail(payload);
}

/***********************
 * EMAIL ACKNOWLEDGEMENT - ISSUED MODE
 ***********************/
function sendIssuanceAcknowledgementEmail(payload) {
  if (!payload.email) return;

  const subject = `Document Issuance Acknowledgement - ${payload.applicationNo}`;

  let table = `
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; width:100%;">
      <tr>
        <th>SL.NO</th>
        <th>Document</th>
        <th>Quantity</th>
        <th>Issued By</th>
        <th>Date</th>
        <th>Status</th>
        <th>Remarks</th>
      </tr>`;

  let sl = 1;
  payload.documents.forEach(d => {
    const dateStr = d.issuedDate
      ? d.issuedDate
      : Utilities.formatDate(new Date(), TZ, DATE_FORMAT);

    table += `
          <tr>
            <td>${sl++}</td>
            <td>${d.name}</td>
            <td style="text-align:center">${d.quantity || 1}</td>
            <td>${d.issuedBy || payload.operatorName}</td>
            <td>${dateStr}</td>
            <td>${d.status || "Issued"}</td>
            <td>${d.remarks || "-"}</td>
          </tr>`;
  });

  table += `</table>`;

  const body = `
      <p>Dear Student / Parent,</p>

<p>
  This is to acknowledge the issuance of the following documents to the student.
  Please keep this acknowledgement for your records.
</p>

<p><strong>Student Details:</strong></p>
<p>
  Name: ${payload.studentName}<br>
  Application No: ${payload.applicationNo}<br>
  Section: ${payload.section || "-"}<br>
  Sub-Section: ${payload.subSection || "-"}
</p>

${table}

<p>
  Please collect the documents from the office at your earliest convenience.
  In case of any discrepancy, please contact the office immediately.
</p>

<p>
  Yours faithfully,<br>
  Office<br>
  <strong>REVA Independent PU College</strong><br>
  Rukmini Knowledge Park,<br>
  Kattigenahalli, Yelahanka,<br>
  Bangalore – 560064.<br><br>

  Phone: 080-46966966 (Ext: 184)<br>
  Email: <a href="mailto:ripu.k@reva.edu.in">ripu.k@reva.edu.in</a>
</p>

    `;

  MailApp.sendEmail({
    to: payload.email,
    subject: subject,
    htmlBody: body
  });

  logEmail(payload);
}

/***********************
 * EMAIL LOG
 ***********************/
function logEmail(payload) {
  const sh = getSS().getSheetByName(SHEETS.EMAIL_LOG);
  sh.appendRow([
    new Date(),
    payload.applicationNo,
    payload.studentName,
    payload.email,
    payload.operatorName
  ]);
}
