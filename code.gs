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
  EMAIL_LOG: "Email Log",
  DOCUMENT_MASTER: "Document Master"
};

// RECEIVED DOCUMENTS LIST
const RECEIVED_DOCUMENTS = [
  "10Th marks sheet Original",
  "10Th marks sheet Xerox",
  "10Th Hall Ticket Xerox",
  "10Th Transfer Certificate Original",
  "10Th Transfer Certificate Xerox",
  "Study Certificate Xerox",
  "Migration Certificate Original",
  "Migration Certificate Xerox",
  "Student Photo",
  "Caste And Income Certificate Xerox",
  "Aadhar Card Student Xerox",
  "Aadhar Card Mother Xerox",
  "Aadhar Card Father Xerox"
];

// ISSUED DOCUMENTS LIST
const ISSUED_DOCUMENTS = [
  "I PUC Marks Card Original",
  "I PUC Transfer Certificate Original",
  "I PUC Study Certificate Only Original",
  "II PUC Marks Card Original",
  "II PUC Transfer Certificate Original",
  "II PUC Study Certificate Only Original",
  "I and II PUC Study Certificate Original"
];

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
 * GET DOCUMENT LISTS
 ***********************/
function getReceivedDocumentsList() {
  return RECEIVED_DOCUMENTS;
}

function getIssuedDocumentsList() {
  return ISSUED_DOCUMENTS;
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
  if (!sh) return [];
  
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
 * SAVE DOCUMENTS - RECEIVED MODE
 ***********************/
function saveDocuments(payload) {
  try {
    const sh = getSS().getSheetByName(SHEETS.DOCUMENTS);
    if (!sh) throw new Error("Documents sheet not found");
    
    const data = sh.getDataRange().getValues();
    const now = new Date();

    payload.documents.forEach(doc => {
      let rowIndex = -1;
      let existing = null;
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][0]).trim() === String(payload.applicationNo).trim() &&
          String(data[i][2]).trim() === String(doc.name).trim()
        ) {
          rowIndex = i + 1;
          existing = data[i];
          break;
        }
      }

      const existingOriginal = existing ? toBool(existing[3]) : false;
      const existingXerox = existing ? toBool(existing[4]) : false;

      const wantsOriginal = Boolean(doc.original);
      const wantsXerox = Boolean(doc.xerox);

      const newOriginalSubmitted = wantsOriginal && !existingOriginal;
      const newXeroxSubmitted = wantsXerox && !existingXerox;
      const isNewSubmission = newOriginalSubmitted || newXeroxSubmitted;

      if (existing && existingOriginal && existingXerox && !isNewSubmission) {
        return;
      }

      const finalOriginal = existingOriginal || wantsOriginal;
      const finalXerox = existingXerox || wantsXerox;

      let submitDateValue = "";
      let submittedByValue = "";
      let receivedByValue = "";

      if (existing && existing[6]) {
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
        submittedByValue = doc.submittedBy || submittedByValue || "";
        if (!receivedByValue) {
          receivedByValue = payload.operatorName || "";
        }
      }

      const status = (finalOriginal || finalXerox) ? "Submitted" : "Pending";

      const row = [
        payload.applicationNo,
        payload.studentName,
        doc.name,
        finalOriginal,
        finalXerox,
        status,
        submitDateValue,
        submittedByValue,
        receivedByValue,
        now,
        payload.operatorId
      ];

      if (rowIndex > 0) {
        sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }
    });

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
 ***********************/
function saveIssuedDocuments(payload) {
  try {
    let sh = getSS().getSheetByName(SHEETS.DOCUMENTS_ISSUED);
    
    if (!sh) {
      sh = getSS().insertSheet(SHEETS.DOCUMENTS_ISSUED);
      sh.appendRow([
        "Application No",
        "Student Name",
        "Document Name",
        "Quantity",
        "Issued Date",
        "Issued By",
        "Received By",
        "Remarks",
        "Status",
        "Last Updated",
        "Updated By"
      ]);
    }

    const data = sh.getDataRange().getValues();
    const now = new Date();

    payload.documents.forEach(doc => {
      let rowIndex = -1;
      let existing = null;
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][0]).trim() === String(payload.applicationNo).trim() &&
          String(data[i][2]).trim() === String(doc.name).trim()
        ) {
          rowIndex = i + 1;
          existing = data[i];
          break;
        }
      }

      let issuedDateValue = now;
      let issuedByValue = doc.issuedBy || payload.operatorName || "";
      let receivedByValue = doc.receivedBy || "";
      let remarksValue = doc.remarks || "";
      let quantityValue = doc.quantity || 1;
      let statusValue = doc.status || "Issued";

      if (existing) {
        issuedDateValue = existing[4] || now;
        issuedByValue = existing[5] || issuedByValue;
        receivedByValue = doc.receivedBy || existing[6] || "";
        remarksValue = doc.remarks || existing[7] || "";
        quantityValue = doc.quantity ? (parseInt(existing[3] || 0) + parseInt(doc.quantity)) : (existing[3] || 1);
        statusValue = doc.status || existing[8] || "Issued";
      }

      const row = [
        payload.applicationNo,
        payload.studentName,
        doc.name,
        quantityValue,
        issuedDateValue,
        issuedByValue,
        receivedByValue,
        remarksValue,
        statusValue,
        now,
        payload.operatorId
      ];

      if (rowIndex > 0) {
        sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }
    });

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
