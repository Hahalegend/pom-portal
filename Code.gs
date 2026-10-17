function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action;

  if (action === "login") {
    const pin = String(params.pin || "").trim();
    const user = findUser(ss, pin);
    return user
      ? sendJSON({ success: true, user: user })
      : sendJSON({ success: false, message: "Invalid PIN" });
  }

  if (action === "getData") {
    const user = findUser(ss, params.pin);
    if (!user) return sendJSON({ success: false, message: "A valid PIN is required" });
    const taskSheet = ss.getSheetByName("Station_Tasks");
    const kpiSheet = ss.getSheetByName("Mill_KPI");
    if (!taskSheet || !kpiSheet) {
      return sendJSON({ success: false, message: "Station_Tasks or Mill_KPI sheet not found" });
    }

    const taskData = taskSheet.getLastRow() > 1 ? taskSheet.getDataRange().getValues() : [[]];
    const kpiData = kpiSheet.getLastRow() > 1 ? kpiSheet.getDataRange().getValues() : [[]];
    const tasks = taskData.slice(1).map(row => ({
      id: row[0], station: row[1], title: row[2], priority: row[3],
      assigned: row[4], deadline: row[5], status: row[6], delayReason: row[7]
    }));

    return sendJSON({
      tasks: isContractor(user.role) ? [] : tasks,
      kpi: isContractor(user.role) ? [] : (kpiData[1] || []),
      forms: isContractor(user.role) ? [] : readForms(ss)
    });
  }

  if (action === "getAccounts") {
    const user = findUser(ss, params.pin);
    if (!user || !hasAccountManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Manager access is required to view accounts" });
    }
    const userSheet = ss.getSheetByName("Users");
    if (!userSheet) return sendJSON({ success: false, message: "Users sheet not found" });
    const rows = userSheet.getLastRow() > 1
      ? userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 4).getValues()
      : [];
    return sendJSON({
      accounts: rows.map((row, index) => ({
        accountId: index + 2,
        name: String(row[1]),
        role: String(row[2]),
        station: String(row[3])
      }))
    });
  }

  if (action === "getSubmissions") {
    const user = findUser(ss, params.pin);
    if (!user || !hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to view reports" });
    }
    return sendJSON({ submissions: readSubmissions(ss) });
  }

  if (action === "checkSubmission") {
    const user = findUser(ss, params.pin);
    if (!user) return sendJSON({ success: false, message: "A valid PIN is required" });
    const submissionId = String(params.submissionId || "");
    const submission = readSubmissions(ss).find(item => item.id === submissionId);
    return sendJSON({
      exists: Boolean(submission && (hasManagementAccess(user.role) || submission.submittedBy === user.name))
    });
  }

  if (action === "getWorkReports") {
    const user = findUser(ss, params.pin);
    if (!user || (!hasManagementAccess(user.role) && !isContractor(user.role))) {
      return sendJSON({ success: false, message: "Contractor or management access is required" });
    }
    return sendJSON({ reports: readWorkReports(ss, user, params.pin) });
  }

  if (action === "getWorkReport") {
    const user = findUser(ss, params.pin);
    if (!user || (!hasManagementAccess(user.role) && !isContractor(user.role))) {
      return sendJSON({ success: false, message: "Contractor or management access is required" });
    }
    const report = readWorkReport(ss, String(params.reportId || ""), user, params.pin, true);
    return report
      ? sendJSON({ report: report })
      : sendJSON({ success: false, message: "Work report not found or access denied" });
  }

  return sendJSON({ status: "Ready" });
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const body = JSON.parse(e && e.postData ? e.postData.contents || "{}" : "{}");
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return sendJSON({ success: false, message: "Request body must be a JSON object" });
  }
  const user = findUser(ss, body.pin);
  if (!user) return sendJSON({ success: false, message: "A valid PIN is required" });

  if (body.action === "createAccount") {
    if (!hasAccountManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Manager access is required to create accounts" });
    }
    const name = String(body.name || "").trim();
    const role = String(body.role || "").trim();
    const station = String(body.station || "").trim();
    const newPin = String(body.newPin || "").trim();
    if (!name || name.length > 80) {
      return sendJSON({ success: false, message: "A name of at most 80 characters is required" });
    }
    if (!station || station.length > 80) {
      return sendJSON({ success: false, message: "A station of at most 80 characters is required" });
    }
    if (!/^\d{4}$/.test(newPin)) {
      return sendJSON({ success: false, message: "PIN must contain exactly four digits" });
    }
    if (!isSupportedAccountRole(role)) {
      return sendJSON({ success: false, message: "Unsupported account role" });
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "Account changes are busy. Please try again." });
    }
    try {
      const userSheet = ss.getSheetByName("Users");
      if (!userSheet) return sendJSON({ success: false, message: "Users sheet not found" });
      const existingPins = userSheet.getLastRow() > 1
        ? userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getValues().flat()
        : [];
      if (existingPins.some(existingPin => String(existingPin).trim() === newPin)) {
        return sendJSON({ success: false, message: "That PIN is already assigned to an account" });
      }
      const newRow = userSheet.getLastRow() + 1;
      userSheet.getRange(newRow, 1).setNumberFormat("@");
      userSheet.getRange(newRow, 1, 1, 4).setValues([[newPin, name, role, station]]);
      return sendJSON({ success: true });
    } finally {
      lock.releaseLock();
    }
  }

  if (body.action === "updateAccountRole") {
    if (!hasAccountManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Manager access is required to update account roles" });
    }
    const accountId = Number(body.accountId);
    const expectedName = String(body.expectedName || "");
    const expectedStation = String(body.expectedStation || "");
    const expectedRole = String(body.expectedRole || "");
    const role = String(body.role || "").trim();
    if (!Number.isInteger(accountId) || accountId < 2 || !expectedName || !isSupportedAccountRole(role)) {
      return sendJSON({ success: false, message: "Invalid account or role" });
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "Account changes are busy. Please try again." });
    }
    try {
      const userSheet = ss.getSheetByName("Users");
      if (!userSheet || accountId > userSheet.getLastRow()) {
        return sendJSON({ success: false, message: "Account not found; refresh the account list and try again" });
      }
      const target = userSheet.getRange(accountId, 1, 1, 4).getValues()[0];
      if (String(target[1]) !== expectedName || String(target[2]) !== expectedRole || String(target[3]) !== expectedStation) {
        return sendJSON({ success: false, message: "Account changed; refresh the account list and try again" });
      }
      const callerPin = String(body.pin).trim();
      if (String(target[0]).trim() === callerPin) {
        return sendJSON({ success: false, message: "You cannot change your own account role" });
      }
      userSheet.getRange(accountId, 3).setValue(role);
      return sendJSON({ success: true });
    } finally {
      lock.releaseLock();
    }
  }

  if (body.action === "createTask") {
    if (!hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to create tasks" });
    }
    const taskSheet = ss.getSheetByName("Station_Tasks");
    if (!taskSheet) return sendJSON({ success: false, message: "Station_Tasks sheet not found" });
    if (!body.title || !body.station) {
      return sendJSON({ success: false, message: "A task title and station are required" });
    }
    const taskId = String(body.taskId || Utilities.getUuid());
    taskSheet.appendRow([
      taskId, body.station, body.title, body.priority || "High",
      body.assigned || "Shift Operator", body.deadline || "Today", "Pending", "", new Date()
    ]);
    return sendJSON({ success: true, taskId: taskId });
  }

  if (body.action === "updateTask") {
    if (!hasManagementAccess(user.role) && !isWorkerRole(user.role)) {
      return sendJSON({ success: false, message: "Not authorized to update tasks" });
    }
    const taskSheet = ss.getSheetByName("Station_Tasks");
    if (!taskSheet) return sendJSON({ success: false, message: "Station_Tasks sheet not found" });
    if (body.status !== "Completed" && body.status !== "Delayed") {
      return sendJSON({ success: false, message: "Unsupported task status" });
    }
    if (body.status === "Delayed" && !String(body.delayReason || "").trim()) {
      return sendJSON({ success: false, message: "A delay reason is required" });
    }
    const data = taskSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.taskId)) {
        const userCanUpdate = hasManagementAccess(user.role)
          || String(data[i][1]) === String(user.station)
          || String(data[i][4]).includes(String(user.name));
        if (!userCanUpdate) {
          return sendJSON({ success: false, message: "You are not assigned to this task" });
        }
        taskSheet.getRange(i + 1, 7).setValue(body.status);
        taskSheet.getRange(i + 1, 8).setValue(body.delayReason || "");
        taskSheet.getRange(i + 1, 9).setValue(new Date());
        return sendJSON({ success: true });
      }
    }
    return sendJSON({ success: false, message: "Task not found" });
  }

  if (body.action === "deleteTask") {
    if (!hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management or executive access is required to delete tasks" });
    }
    const taskSheet = ss.getSheetByName("Station_Tasks");
    if (!taskSheet) return sendJSON({ success: false, message: "Station_Tasks sheet not found" });
    const data = taskSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.taskId)) {
        taskSheet.deleteRow(i + 1);
        return sendJSON({ success: true });
      }
    }
    return sendJSON({ success: false, message: "Task not found" });
  }

  if (body.action === "createForm") {
    if (!hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to create screens" });
    }
    const formSheet = getFormSheet(ss);
    let form;
    try {
      form = validateForm(body.form);
    } catch (error) {
      return sendJSON({ success: false, message: error.message });
    }
    const existingIds = formSheet.getLastRow() > 1
      ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 1).getValues().flat()
      : [];
    if (existingIds.includes(form.id)) {
      return sendJSON({ success: false, message: "A screen with this ID already exists" });
    }
    formSheet.appendRow([form.id, form.name, form.type, JSON.stringify(form.fields), new Date()]);
    return sendJSON({ success: true, form: form });
  }

  if (body.action === "submitForm") {
    const form = readForms(ss).find(item => item.id === String(body.formId || ""));
    if (!form) return sendJSON({ success: false, message: "Screen not found" });
    let submittedValues;
    try {
      submittedValues = validateSubmission(form, body.values);
    } catch (error) {
      return sendJSON({ success: false, message: error.message });
    }
    if (!hasManagementAccess(user.role) && !isWorkerRole(user.role)) {
      return sendJSON({ success: false, message: "Not authorized to submit forms" });
    }
    const submissionSheet = getSubmissionSheet(ss);
    submissionSheet.appendRow([
      String(body.submissionId || Utilities.getUuid()), form.id, form.name, form.type,
      user.name.slice(0, 120),
      user.station.slice(0, 120),
      new Date(), JSON.stringify(submittedValues)
    ]);
    return sendJSON({ success: true });
  }

  if (body.action === "createWorkReport") {
    if (!isContractor(user.role)) {
      return sendJSON({ success: false, message: "Only contractor accounts can submit work reports" });
    }
    let report;
    const uploadedPhotos = [];
    try {
      report = validateWorkReport(body.report);
      const reportSheet = getWorkReportSheet(ss);
      const exists = reportSheet.getLastRow() > 1
        && reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, 1).getValues().flat().includes(report.id);
      if (exists) return sendJSON({ success: false, message: "A report with this ID already exists" });

      report.before = saveWorkReportPhotos(report.before, "before", report.id);
      uploadedPhotos.push.apply(uploadedPhotos, report.before);
      report.after = saveWorkReportPhotos(report.after, "after", report.id);
      uploadedPhotos.push.apply(uploadedPhotos, report.after);
      reportSheet.appendRow([
        report.id, report.workName, report.description, report.comment, user.name,
        user.station, hashPin(body.pin), new Date(), JSON.stringify(report.before), JSON.stringify(report.after)
      ]);
    } catch (error) {
      trashWorkReportPhotos(uploadedPhotos);
      return sendJSON({ success: false, message: error.message });
    }
    return sendJSON({ success: true, reportId: report.id });
  }

  return sendJSON({ success: false, message: "Unknown action" });
}

function hasManagementAccess(role) {
  const normalized = String(role || "").trim().toLowerCase();
  return normalized === "admin" || normalized === "management" || normalized === "manager" || normalized === "executive";
}

function hasAccountManagementAccess(role) {
  const normalized = String(role || "").trim().toLowerCase();
  return normalized === "admin" || normalized === "management" || normalized === "manager";
}

function isSupportedAccountRole(role) {
  return ["admin", "management", "manager", "executive", "worker", "contractor"]
    .includes(String(role || "").trim().toLowerCase());
}

function isContractor(role) {
  return String(role || "").trim().toLowerCase() === "contractor";
}

function isWorkerRole(role) {
  return String(role || "").trim().toLowerCase() === "worker";
}

function hashPin(pin) {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty("WORK_REPORT_PIN_HASH_SECRET");
  if (!secret) {
    secret = Utilities.getUuid();
    properties.setProperty("WORK_REPORT_PIN_HASH_SECRET", secret);
  }
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    secret + ":" + String(pin).trim(),
    Utilities.Charset.UTF_8
  ).map(byte => (byte + 256).toString(16).slice(-2)).join("");
}

function findUser(ss, pin) {
  const userSheet = ss.getSheetByName("Users");
  if (!userSheet || !pin) return null;
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(pin).trim()) {
      return { name: String(data[i][1]), role: String(data[i][2]), station: String(data[i][3]) };
    }
  }
  return null;
}

function getFormSheet(ss) {
  let sheet = ss.getSheetByName("Portal_Forms");
  if (!sheet) {
    sheet = ss.insertSheet("Portal_Forms");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Form ID", "Name", "Type", "Fields JSON", "Created At"]);
  }
  return sheet;
}

function getSubmissionSheet(ss) {
  let sheet = ss.getSheetByName("Portal_Submissions");
  if (!sheet) {
    sheet = ss.insertSheet("Portal_Submissions");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Submission ID", "Form ID", "Form Name", "Type", "Submitted By", "Station", "Submitted At", "Values JSON"]);
  }
  return sheet;
}

function readForms(ss) {
  const sheet = getFormSheet(ss);
  if (sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().map(row => ({
    id: String(row[0]),
    name: String(row[1]),
    type: String(row[2]),
    fields: JSON.parse(String(row[3] || "[]"))
  }));
}

function readSubmissions(ss) {
  const sheet = getSubmissionSheet(ss);
  if (sheet.getLastRow() <= 1) return [];
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  return rows.reverse().map(row => ({
    id: String(row[0]),
    formId: String(row[1]),
    formName: String(row[2]),
    type: String(row[3]),
    submittedBy: String(row[4]),
    station: String(row[5]),
    submittedAt: row[6] instanceof Date ? row[6].toISOString() : String(row[6]),
    values: JSON.parse(String(row[7] || "{}"))
  }));
}

function getWorkReportSheet(ss) {
  let sheet = ss.getSheetByName("Contractor_Work_Reports");
  if (!sheet) sheet = ss.insertSheet("Contractor_Work_Reports");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Report ID", "Work Name", "Description", "Comment", "Contractor",
      "Station", "Owner PIN Hash", "Created At", "Before Photos JSON", "After Photos JSON"
    ]);
  }
  return sheet;
}

function getWorkReportFolder() {
  const properties = PropertiesService.getScriptProperties();
  const savedId = properties.getProperty("CONTRACTOR_WORK_REPORT_FOLDER_ID");
  if (savedId) return DriveApp.getFolderById(savedId);
  const folder = DriveApp.createFolder("POM Portal Contractor Work Report Photos");
  properties.setProperty("CONTRACTOR_WORK_REPORT_FOLDER_ID", folder.getId());
  return folder;
}

function validateWorkReport(candidate) {
  if (!candidate || typeof candidate !== "object") throw new Error("A work report is required");
  const id = String(candidate.id || "").trim();
  const workName = String(candidate.workName || "").trim();
  const description = String(candidate.description || "").trim();
  const comment = String(candidate.comment || "").trim();
  if (!id || id.length > 100) throw new Error("Invalid work report ID");
  if (!workName || workName.length > 160) throw new Error("Work name is required and must be at most 160 characters");
  if (!description || description.length > 5000) throw new Error("Description is required and must be at most 5000 characters");
  if (comment.length > 3000) throw new Error("Comment must be at most 3000 characters");
  if (!Array.isArray(candidate.before) || !Array.isArray(candidate.after)) {
    throw new Error("Before and after photos must be lists");
  }
  return {
    id: id,
    workName: workName,
    description: description,
    comment: comment,
    before: candidate.before.map(validateWorkReportPhoto),
    after: candidate.after.map(validateWorkReportPhoto)
  };
}

function validateWorkReportPhoto(photo) {
  if (!photo || typeof photo !== "object" || photo.mimeType !== "image/jpeg") {
    throw new Error("Work report photos must be JPEG images");
  }
  const name = String(photo.name || "photo.jpg").slice(0, 180);
  const data = String(photo.data || "");
  if (!data || data.length > 7000000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(data)) {
    throw new Error("A photo is empty, too large, or invalid");
  }
  return { name: name, mimeType: "image/jpeg", data: data };
}

function saveWorkReportPhotos(photos, section, reportId) {
  const saved = [];
  try {
    const folder = getWorkReportFolder();
    photos.forEach((photo, index) => {
      const bytes = Utilities.base64Decode(photo.data);
      const blob = Utilities.newBlob(bytes, photo.mimeType, reportId + "-" + section + "-" + (index + 1) + ".jpg");
      const file = folder.createFile(blob);
      saved.push({ id: file.getId(), name: photo.name, mimeType: photo.mimeType });
    });
    return saved;
  } catch (error) {
    trashWorkReportPhotos(saved);
    throw new Error("Could not store report photos in Google Drive: " + error.message);
  }
}

function trashWorkReportPhotos(photos) {
  photos.forEach(photo => {
    try {
      DriveApp.getFileById(photo.id).setTrashed(true);
    } catch (error) {
      Logger.log("Could not remove incomplete work report photo " + photo.id + ": " + error.message);
    }
  });
}

function readWorkReports(ss, user, pin) {
  const sheet = getWorkReportSheet(ss);
  if (sheet.getLastRow() <= 1) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const ownerHash = hashPin(pin);
  return rows.filter(row => hasManagementAccess(user.role) || String(row[6]) === ownerHash)
    .reverse()
    .map(row => ({
      id: String(row[0]),
      workName: String(row[1]),
      contractorName: String(row[4]),
      station: String(row[5]),
      createdAt: row[7] instanceof Date ? row[7].toISOString() : String(row[7]),
      beforeCount: JSON.parse(String(row[8] || "[]")).length,
      afterCount: JSON.parse(String(row[9] || "[]")).length
    }));
}

function readWorkReport(ss, reportId, user, pin, includePhotos) {
  const sheet = getWorkReportSheet(ss);
  if (sheet.getLastRow() <= 1) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const row = rows.find(item => String(item[0]) === reportId);
  if (!row) return null;
  if (!hasManagementAccess(user.role) && String(row[6]) !== hashPin(pin)) return null;

  const report = {
    id: String(row[0]),
    workName: String(row[1]),
    description: String(row[2]),
    comment: String(row[3]),
    contractorName: String(row[4]),
    station: String(row[5]),
    createdAt: row[7] instanceof Date ? row[7].toISOString() : String(row[7])
  };
  if (includePhotos) {
    report.before = readWorkReportPhotos(JSON.parse(String(row[8] || "[]")));
    report.after = readWorkReportPhotos(JSON.parse(String(row[9] || "[]")));
  }
  return report;
}

function readWorkReportPhotos(photos) {
  return photos.map(photo => {
    const file = DriveApp.getFileById(photo.id);
    return {
      name: photo.name,
      mimeType: photo.mimeType,
      data: Utilities.base64Encode(file.getBlob().getBytes())
    };
  });
}

function validateForm(candidate) {
  if (!candidate || typeof candidate !== "object") throw new Error("A screen definition is required");
  const name = String(candidate.name || "").trim();
  const id = String(candidate.id || "").trim();
  const type = String(candidate.type || "");
  if (!id || id.length > 100 || !name || name.length > 80) throw new Error("Screen ID and a name of at most 80 characters are required");
  if (type !== "Operation" && type !== "Report") throw new Error("Screen type must be Operation or Report");
  if (!Array.isArray(candidate.fields) || candidate.fields.length < 1 || candidate.fields.length > 30) {
    throw new Error("A screen must have between 1 and 30 fields");
  }

  const allowedTypes = ["text", "number", "date", "textarea", "checkbox"];
  const seenLabels = new Set();
  const fields = candidate.fields.map(field => {
    const label = String(field.label || "").trim();
    const fieldType = String(field.type || "");
    const normalizedLabel = label.toLowerCase();
    if (!label || label.length > 80 || seenLabels.has(normalizedLabel)) {
      throw new Error("Field names must be unique and at most 80 characters");
    }
    if (!allowedTypes.includes(fieldType)) throw new Error("Unsupported field type");
    seenLabels.add(normalizedLabel);
    return { label: label, type: fieldType, required: field.required === true };
  });
  return { id: id, name: name, type: type, fields: fields };
}

function validateSubmission(form, values) {
  if (!values || typeof values !== "object" || Array.isArray(values)) {
    throw new Error("Form values must be an object");
  }
  const normalized = {};
  form.fields.forEach(field => {
    const value = values[field.label];
    if (field.required && (value === undefined || value === null || value === "" || (field.type === "checkbox" && value !== true))) {
      throw new Error(field.label + " is required");
    }
    if (field.type === "checkbox") {
      normalized[field.label] = value === true;
      return;
    }
    const text = String(value === undefined || value === null ? "" : value);
    if (text.length > 5000) throw new Error(field.label + " is too long");
    normalized[field.label] = text;
  });
  return normalized;
}

function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
