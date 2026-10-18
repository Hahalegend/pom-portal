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

  if (action === "getTaskPhoto") {
    const user = findUser(ss, params.pin);
    if (!user) return sendJSON({ success: false, message: "A valid PIN is required" });
    const taskSheet = ss.getSheetByName("Station_Tasks");
    if (!taskSheet) return sendJSON({ success: false, message: "Station_Tasks sheet not found" });
    const rows = taskSheet.getLastRow() > 1 ? taskSheet.getDataRange().getValues().slice(1) : [];
    const task = rows.find(row => String(row[0]) === String(params.taskId || ""));
    if (!task) return sendJSON({ success: false, message: "Task not found" });
    if (!canViewTask(user, task)) {
      return sendJSON({ success: false, message: "You are not assigned to this task" });
    }
    const photoType = params.photoType === "completion" ? "completion" : "task";
    const completion = parseTaskCompletion(task[10]);
    const photo = photoType === "completion"
      ? completion.photo
      : JSON.parse(String(task[9] || "null"));
    if (!photo) return sendJSON({ success: false, message: "Task picture not found" });
    const file = DriveApp.getFileById(photo.id);
    return sendJSON({
      photo: {
        name: photo.name,
        mimeType: photo.mimeType,
        data: Utilities.base64Encode(file.getBlob().getBytes())
      }
    });
  }

  if (action === "getSubmissionPhoto") {
    const user = findUser(ss, params.pin);
    if (!user) return sendJSON({ success: false, message: "A valid PIN is required" });
    const submission = readSubmissions(ss).find(item => item.id === String(params.submissionId || ""));
    if (!submission || (!hasManagementAccess(user.role) && submission.submittedBy !== user.name)) {
      return sendJSON({ success: false, message: "Submission not found or access denied" });
    }
    const photo = submission.values[String(params.field || "")];
    if (!photo || !photo.id) return sendJSON({ success: false, message: "Submission picture not found" });
    const file = DriveApp.getFileById(photo.id);
    return sendJSON({
      photo: {
        name: photo.name,
        mimeType: photo.mimeType,
        data: Utilities.base64Encode(file.getBlob().getBytes())
      }
    });
  }

  if (action === "getStations") {
    const user = findUser(ss, params.pin);
    if (!user || (!hasManagementAccess(user.role) && !isContractor(user.role))) {
      return sendJSON({ success: false, message: "Contractor or management access is required to view stations" });
    }
    return sendJSON({ stations: readStations(ss) });
  }

  if (action === "getMetricHistory" || action === "getOerHistory") {
    const user = findUser(ss, params.pin);
    if (!user || !hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to view KPI history" });
    }
    const metric = getKpiMetric(params.metric || "oer");
    if (!metric) return sendJSON({ success: false, message: "Unsupported KPI metric" });
    return sendJSON({ records: readKpiHistory(getKpiHistorySheet(ss, metric)) });
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
      assigned: row[4], deadline: normalizeTaskDate(row[5]), status: row[6],
      delayReason: String(row[12] || row[7] || ""),
      delayDate: normalizeTaskDate(row[11]),
      hasPhoto: Boolean(row[9]),
      hasCompletionPhoto: Boolean(parseTaskCompletion(row[10]).photo),
      completionComment: parseTaskCompletion(row[10]).comment || "",
      completedAt: parseTaskCompletion(row[10]).completedAt || ""
    }));

    return sendJSON({
      tasks: isContractor(user.role) ? [] : tasks,
      kpi: isContractor(user.role) ? [] : (kpiData[1] || []),
      ffa: isContractor(user.role) ? null : getLatestKpiValue(ss, "ffa"),
      forms: isContractor(user.role) ? [] : readForms(ss).filter(form =>
        hasManagementAccess(user.role)
        || form.station === "All stations"
        || form.station.toLowerCase() === String(user.station || "").trim().toLowerCase()
      )
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
    if (!user || !hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required; contractors must use company sign-in" });
    }
    return sendJSON({ reports: readWorkReports(ss, user, null) });
  }

  if (action === "getWorkReport") {
    const user = findUser(ss, params.pin);
    if (!user || !hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required; contractors must use company sign-in" });
    }
    const report = readWorkReport(ss, String(params.reportId || ""), user, true, null);
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

  if (body.action === "getContractorWorkReports" || body.action === "getContractorWorkReport") {
    if (!isContractor(user.role)) {
      return sendJSON({ success: false, message: "Only contractor accounts can access company work reports" });
    }
    const company = verifyContractorCompany(ss, body.companyName, body.companyId);
    if (!company) {
      return sendJSON({ success: false, message: "Company name or unique ID is incorrect. Sign in to your company first." });
    }
    if (body.action === "getContractorWorkReports") {
      return sendJSON({ reports: readWorkReports(ss, user, {
        name: company.name,
        id: String(body.companyId).trim()
      }) });
    }
    const report = readWorkReport(
      ss,
      String(body.reportId || ""),
      user,
      true,
      { name: company.name, id: String(body.companyId).trim() }
    );
    return report
      ? sendJSON({ report: report })
      : sendJSON({ success: false, message: "Work report not found or access denied" });
  }

  if (body.action === "saveMetric" || body.action === "saveOer") {
    if (!hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to save KPI values" });
    }
    const metric = getKpiMetric(body.metric || "oer");
    if (!metric) return sendJSON({ success: false, message: "Unsupported KPI metric" });
    const date = String(body.date || "").trim();
    const valueText = String(body.value == null ? body.oer == null ? "" : body.oer : body.value).trim();
    const value = Number(valueText);
    if (!isValidTaskDate(date) || !valueText || !Number.isFinite(value) || value < 0 || value > metric.max) {
      return sendJSON({ success: false, message: "Enter a valid date and KPI value" });
    }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "KPI history is busy. Please try again." });
    }
    try {
      const sheet = getKpiHistorySheet(ss, metric);
      const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
        : [];
      const rowIndex = rows.findIndex(value => normalizeTaskDate(value) === date);
      const rowNumber = rowIndex < 0 ? sheet.getLastRow() + 1 : rowIndex + 2;
      sheet.getRange(rowNumber, 1).setNumberFormat("@");
      sheet.getRange(rowNumber, 2).setNumberFormat("0.00");
      sheet.getRange(rowNumber, 1, 1, 4).setValues([[
        date, value, new Date(), user.name
      ]]);
      SpreadsheetApp.flush();
      return sendJSON({ success: true, date: date, metric: metric.key, value: value, replaced: rowIndex >= 0 });
    } finally {
      lock.releaseLock();
    }
  }

  if (body.action === "addStation" || body.action === "removeStation") {
    if (!hasManagementAccess(user.role)) {
      return sendJSON({ success: false, message: "Management access is required to change stations" });
    }
    const station = String(body.station || "").trim();
    if (!station || station.length > 80) {
      return sendJSON({ success: false, message: "A station name of at most 80 characters is required" });
    }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "Station list is busy. Please try again." });
    }
    try {
      const sheet = getStationSheet(ss);
      const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
        : [];
      const stationIndex = rows.findIndex(value => String(value).trim().toLowerCase() === station.toLowerCase());
      if (body.action === "addStation") {
        if (stationIndex >= 0) {
          return sendJSON({ success: true, station: String(rows[stationIndex]), alreadyExists: true });
        }
        sheet.appendRow([station]);
        return sendJSON({ success: true, station: station });
      }
      if (stationIndex < 0) {
        return sendJSON({ success: false, message: "Station not found in the saved list" });
      }
      if (rows.length <= 1) {
        return sendJSON({ success: false, message: "At least one station must remain in the list" });
      }
      sheet.deleteRow(stationIndex + 2);
      return sendJSON({ success: true, station: String(rows[stationIndex]) });
    } finally {
      lock.releaseLock();
    }
  }

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
    const title = String(body.title || "").trim();
    const station = String(body.station || "").trim();
    const deadline = String(body.deadline || "").trim();
    if (!title || !station) {
      return sendJSON({ success: false, message: "A task title and station are required" });
    }
    if (!isValidTaskDate(deadline)) {
      return sendJSON({ success: false, message: "A valid completion deadline is required" });
    }
    if (!readStations(ss).some(savedStation => savedStation.toLowerCase() === station.toLowerCase())) {
      return sendJSON({ success: false, message: "Select a station from the saved station list" });
    }
    const taskId = String(body.taskId || Utilities.getUuid());
    let photo = null;
    if (body.photo) {
      try {
        photo = saveTaskPhoto(body.photo, taskId);
      } catch (error) {
        return sendJSON({ success: false, message: error.message });
      }
    }
    try {
      taskSheet.appendRow([
        taskId, station, title, body.priority || "High",
        body.assigned || "Shift Operator", deadline, "Pending", "", new Date(),
        photo ? JSON.stringify(photo) : ""
      ]);
    } catch (error) {
      if (photo) trashTaskPhoto(photo);
      return sendJSON({ success: false, message: "Could not save the task: " + error.message });
    }
    return sendJSON({ success: true, taskId: taskId });
  }

  if (body.action === "updateTask") {
    if (isContractor(user.role)) {
      return sendJSON({ success: false, message: "Not authorized to update tasks" });
    }
    const taskSheet = ss.getSheetByName("Station_Tasks");
    if (!taskSheet) return sendJSON({ success: false, message: "Station_Tasks sheet not found" });
    if (body.status !== "Completed" && body.status !== "Delayed") {
      return sendJSON({ success: false, message: "Unsupported task status" });
    }
    const delayReason = String(body.delayReason || "").trim();
    const delayDate = String(body.delayDate || "").trim();
    const completionComment = String(body.comment || "").trim();
    if (body.status === "Delayed" && (
      !delayReason
      || delayReason.length > 3000
      || !isValidTaskDate(delayDate)
      || delayDate < Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
    )) {
      return sendJSON({ success: false, message: "A revised date and justification of at most 3000 characters are required" });
    }
    if (body.status === "Completed" && completionComment.length > 3000) {
      return sendJSON({ success: false, message: "Completion comment must be at most 3000 characters" });
    }
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "Task updates are busy. Please try again." });
    }
    try {
      const data = taskSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(body.taskId)) {
          const userCanUpdate = canViewTask(user, data[i]);
          if (!userCanUpdate) {
            return sendJSON({ success: false, message: "You are not assigned to this task" });
          }
          let newPhoto = null;
          let completion = parseTaskCompletion(data[i][10]);
          const previousCompletionPhoto = completion.photo;
          if (body.status === "Completed") {
            if (body.photo) {
              try {
                newPhoto = saveTaskPhoto(body.photo, String(body.taskId) + "-completion");
              } catch (error) {
                return sendJSON({ success: false, message: error.message });
              }
            }
            completion = {
              completedAt: new Date().toISOString(),
              comment: completionComment,
              photo: newPhoto || completion.photo || null
            };
          }
          const nextDelayDate = body.status === "Delayed" ? delayDate : data[i][11] || "";
          const nextDelayReason = body.status === "Delayed" ? delayReason : data[i][12] || "";
          try {
            taskSheet.getRange(i + 1, 7, 1, 7).setValues([[
              body.status,
              body.status === "Delayed" ? delayReason : "",
              new Date(),
              data[i][9] || "",
              JSON.stringify(completion),
              nextDelayDate,
              nextDelayReason
            ]]);
            SpreadsheetApp.flush();
            const savedRow = taskSheet.getRange(i + 1, 7, 1, 7).getValues()[0];
            const savedCompletion = body.status === "Completed" ? parseTaskCompletion(savedRow[4]) : null;
            const updateSaved = String(savedRow[0]) === body.status
              && (body.status !== "Delayed" || (
                normalizeTaskDate(savedRow[5]) === delayDate
                && String(savedRow[6]) === delayReason
              ))
              && (body.status !== "Completed" || (
                savedCompletion.comment === completionComment
                && (!body.photo || (savedCompletion.photo && savedCompletion.photo.id === newPhoto.id))
              ));
            if (!updateSaved) {
              return sendJSON({ success: false, message: "The task update could not be verified in the spreadsheet" });
            }
          } catch (error) {
            return sendJSON({ success: false, message: "Could not update the task: " + error.message });
          }
          if (newPhoto && previousCompletionPhoto && previousCompletionPhoto.id) {
            trashTaskPhoto(previousCompletionPhoto);
          }
          return sendJSON({ success: true });
        }
      }
      return sendJSON({ success: false, message: "Task not found" });
    } finally {
      lock.releaseLock();
    }
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
        const photo = JSON.parse(String(data[i][9] || "null"));
        const completion = parseTaskCompletion(data[i][10]);
        taskSheet.deleteRow(i + 1);
        if (photo) trashTaskPhoto(photo);
        if (completion.photo) trashTaskPhoto(completion.photo);
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
      form = validateForm(body.form, ss);
    } catch (error) {
      return sendJSON({ success: false, message: error.message });
    }
    const existingIds = formSheet.getLastRow() > 1
      ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 1).getValues().flat()
      : [];
    if (existingIds.includes(form.id)) {
      return sendJSON({ success: false, message: "A screen with this ID already exists" });
    }
    formSheet.appendRow([form.id, form.name, form.type, JSON.stringify(form.fields), new Date(), form.station]);
    return sendJSON({ success: true, form: form });
  }

  if (body.action === "submitForm") {
    const form = readForms(ss).find(item => item.id === String(body.formId || ""));
    if (!form) return sendJSON({ success: false, message: "Screen not found" });
    if (!hasManagementAccess(user.role) && !isWorkerRole(user.role)) {
      return sendJSON({ success: false, message: "Not authorized to submit forms" });
    }
    if (!hasManagementAccess(user.role) && form.station !== "All stations"
      && form.station.toLowerCase() !== String(user.station || "").trim().toLowerCase()) {
      return sendJSON({ success: false, message: "This form is not assigned to your station" });
    }
    const submissionId = String(body.submissionId || Utilities.getUuid()).trim();
    if (!submissionId || submissionId.length > 100) {
      return sendJSON({ success: false, message: "Invalid submission ID" });
    }
    let submittedValues;
    try {
      submittedValues = validateSubmission(form, body.values);
    } catch (error) {
      return sendJSON({ success: false, message: error.message });
    }
    const savedPhotos = [];
    try {
      form.fields.filter(field => field.type === "photo").forEach((field, index) => {
        const candidate = submittedValues[field.label];
        if (!candidate) return;
        const photo = saveTaskPhoto(candidate, submissionId + "-" + index);
        savedPhotos.push(photo);
        submittedValues[field.label] = photo;
      });
      getSubmissionSheet(ss).appendRow([
        submissionId, form.id, form.name, form.type,
        user.name.slice(0, 120),
        user.station.slice(0, 120),
        new Date(), JSON.stringify(submittedValues)
      ]);
    } catch (error) {
      savedPhotos.forEach(trashTaskPhoto);
      return sendJSON({ success: false, message: "Could not save the submission: " + error.message });
    }
    return sendJSON({ success: true });
  }

  if (body.action === "createWorkReport") {
    if (!isContractor(user.role)) {
      return sendJSON({ success: false, message: "Only contractor accounts can submit work reports" });
    }
    const company = verifyContractorCompany(ss, body.companyName, body.companyId);
    if (!company) {
      return sendJSON({ success: false, message: "Sign in with your registered company name and unique ID before submitting" });
    }
    let report;
    const uploadedPhotos = [];
    try {
      report = validateWorkReport(body.report, ss);
      const reportSheet = getWorkReportSheet(ss);
      const exists = reportSheet.getLastRow() > 1
        && reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, 1).getValues().flat().includes(report.id);
      if (exists) return sendJSON({ success: false, message: "A report with this ID already exists" });

      report.before = saveWorkReportPhotos(report.before, "before", report.id);
      uploadedPhotos.push.apply(uploadedPhotos, report.before);
      report.after = saveWorkReportPhotos(report.after, "after", report.id);
      uploadedPhotos.push.apply(uploadedPhotos, report.after);
      reportSheet.appendRow([
        report.id, report.workName, report.description, report.comment, report.contractorName,
        report.station, hashPin(body.pin), new Date(), JSON.stringify(report.before), JSON.stringify(report.after),
        hashContractorCompany(company.name, body.companyId)
      ]);
    } catch (error) {
      trashWorkReportPhotos(uploadedPhotos);
      return sendJSON({ success: false, message: error.message });
    }
    return sendJSON({ success: true, reportId: report.id });
  }

  if (body.action === "registerContractorCompany" || body.action === "verifyContractorCompany") {
    if (!isContractor(user.role)) {
      return sendJSON({ success: false, message: "Only contractor accounts can access company work reports" });
    }
    let credentials;
    try {
      credentials = validateContractorCompanyCredentials(body.companyName, body.companyId);
    } catch (error) {
      return sendJSON({ success: false, message: error.message });
    }

    if (body.action === "verifyContractorCompany") {
      const company = verifyContractorCompany(ss, credentials.name, credentials.id);
      return company
        ? sendJSON({ success: true, companyName: company.name })
        : sendJSON({ success: false, message: "Company name or unique ID is incorrect. Register your company if this is your first visit." });
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return sendJSON({ success: false, message: "Company registration is busy. Please try again." });
    }
    try {
      const sheet = getContractorCompanySheet(ss);
      const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
        : [];
      const normalizedName = normalizeCompanyName(credentials.name);
      const companyIdHash = hashPin("contractor-company-id:" + credentials.id);
      if (rows.some(row =>
        normalizeCompanyName(row[0]) === normalizedName
        || String(row[1]) === companyIdHash
      )) {
        return sendJSON({ success: false, message: "That company name or unique ID is already registered. Sign in or choose different details." });
      }
      sheet.appendRow([credentials.name, companyIdHash, new Date()]);
      return sendJSON({ success: true, companyName: credentials.name });
    } finally {
      lock.releaseLock();
    }
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
  return ["admin", "management", "manager", "executive", "worker", "operator", "contractor"]
    .includes(String(role || "").trim().toLowerCase());
}

function isContractor(role) {
  return String(role || "").trim().toLowerCase() === "contractor";
}

function isWorkerRole(role) {
  const normalized = String(role || "").trim().toLowerCase();
  return normalized === "worker" || normalized === "operator";
}

function canViewTask(user, task) {
  return hasManagementAccess(user.role)
    || (!isContractor(user.role)
      && (String(task[1]) === String(user.station) || String(task[4]).includes(String(user.name))));
}

function isValidTaskDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return false;
  const date = new Date(value + "T00:00:00Z");
  return !isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function normalizeTaskDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const date = String(value || "").trim();
  return isValidTaskDate(date) ? date : "";
}

function parseTaskCompletion(value) {
  if (!value) return {};
  const completion = typeof value === "object" ? value : JSON.parse(String(value));
  return completion && typeof completion === "object" ? completion : {};
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
    sheet.appendRow(["Form ID", "Name", "Type", "Fields JSON", "Created At", "Station"]);
  } else if (!String(sheet.getRange(1, 6).getValue() || "").trim()) {
    sheet.getRange(1, 6).setValue("Station");
  }
  return sheet;
}

function getStationSheet(ss) {
  let sheet = ss.getSheetByName("Portal_Stations");
  if (!sheet) sheet = ss.insertSheet("Portal_Stations");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Station"]);
    sheet.getRange(2, 1, 6, 1).setValues([
      ["Loading Ramp"],
      ["Sterilizer"],
      ["Pressing"],
      ["Clarification"],
      ["Nut & Kernel"],
      ["Boiler & Power"]
    ]);
  }
  return sheet;
}

function getKpiMetric(key) {
  const metrics = {
    oer: { key: "oer", label: "OER", sheetName: "OER_History", unit: "%", max: 100 },
    ker: { key: "ker", label: "KER", sheetName: "KER_History", unit: "%", max: 100 },
    ffb: { key: "ffb", label: "FFB Process", sheetName: "FFB_History", unit: "MT", max: 1000000000 },
    ffa: { key: "ffa", label: "FFA Process", sheetName: "FFA_History", unit: "%", max: 100 },
    boiler: { key: "boiler", label: "Boiler Pressure", sheetName: "Boiler_History", unit: "Bar", max: 1000000000 }
  };
  return metrics[String(key || "").trim().toLowerCase()] || null;
}

function getKpiHistorySheet(ss, metric) {
  let sheet = ss.getSheetByName(metric.sheetName);
  if (!sheet) sheet = ss.insertSheet(metric.sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", metric.label + (metric.unit === "%" ? " (%)" : " (" + metric.unit + ")"), "Updated At", "Updated By"]);
  }
  return sheet;
}

function readKpiHistory(sheet) {
  if (sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(row => ({
      date: normalizeTaskDate(row[0]),
      value: row[1] === "" || row[1] == null ? NaN : Number(row[1]),
      updatedAt: row[2] instanceof Date ? row[2].toISOString() : String(row[2] || ""),
      updatedBy: String(row[3] || "")
    }))
    .filter(record => record.date && Number.isFinite(record.value) && record.value >= 0)
    .sort((left, right) => left.date.localeCompare(right.date));
}

function getLatestKpiValue(ss, key) {
  const metric = getKpiMetric(key);
  if (!metric) return null;
  const sheet = ss.getSheetByName(metric.sheetName);
  if (!sheet) return null;
  const records = readKpiHistory(sheet);
  return records.length ? records[records.length - 1].value : null;
}

function readStations(ss) {
  const sheet = getStationSheet(ss);
  if (sheet.getLastRow() <= 1) return [];
  const stations = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
    .map(value => String(value).trim())
    .filter(Boolean);
  return Array.from(new Set(stations));
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
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(6, sheet.getLastColumn())).getValues().map(row => ({
    id: String(row[0]),
    name: String(row[1]),
    type: String(row[2]),
    fields: JSON.parse(String(row[3] || "[]")),
    station: String(row[5] || "All stations")
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
      "Station", "Owner PIN Hash", "Created At", "Before Photos JSON", "After Photos JSON", "Company ID Hash"
    ]);
  } else if (!String(sheet.getRange(1, 11).getValue() || "").trim()) {
    sheet.getRange(1, 11).setValue("Company ID Hash");
  }
  return sheet;
}

function getContractorCompanySheet(ss) {
  let sheet = ss.getSheetByName("Contractor_Companies");
  if (!sheet) sheet = ss.insertSheet("Contractor_Companies");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Company Name", "Company ID Hash", "Registered At"]);
  }
  return sheet;
}

function normalizeCompanyName(name) {
  return String(name || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function validateContractorCompanyCredentials(name, id) {
  const normalizedName = String(name || "").trim().replace(/\s+/g, " ");
  const normalizedId = String(id || "").trim();
  if (!normalizedName || normalizedName.length > 120) {
    throw new Error("Enter a company name of at most 120 characters");
  }
  if (normalizedId.length < 12 || normalizedId.length > 100) {
    throw new Error("Company unique ID must be between 12 and 100 characters");
  }
  return { name: normalizedName, id: normalizedId };
}

function verifyContractorCompany(ss, name, id) {
  const companyName = String(name || "").trim().replace(/\s+/g, " ");
  const companyId = String(id || "").trim();
  if (!companyName || companyName.length > 120 || companyId.length < 12 || companyId.length > 100) return null;
  const sheet = getContractorCompanySheet(ss);
  if (sheet.getLastRow() <= 1) return null;
  const normalizedName = normalizeCompanyName(companyName);
  const idHash = hashPin("contractor-company-id:" + companyId);
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().find(row =>
    normalizeCompanyName(row[0]) === normalizedName && String(row[1]) === idHash
  );
  return match ? { name: String(match[0]) } : null;
}

function hashContractorCompany(name, id) {
  const credentials = validateContractorCompanyCredentials(name, id);
  return hashPin("contractor-company-report:" + normalizeCompanyName(credentials.name) + ":" + credentials.id);
}

function getWorkReportFolder() {
  const properties = PropertiesService.getScriptProperties();
  const savedId = properties.getProperty("CONTRACTOR_WORK_REPORT_FOLDER_ID");
  if (savedId) return DriveApp.getFolderById(savedId);
  const folder = DriveApp.createFolder("POM Portal Contractor Work Report Photos");
  properties.setProperty("CONTRACTOR_WORK_REPORT_FOLDER_ID", folder.getId());
  return folder;
}

function getTaskPhotoFolder() {
  const properties = PropertiesService.getScriptProperties();
  const savedId = properties.getProperty("STATION_TASK_PHOTO_FOLDER_ID");
  if (savedId) return DriveApp.getFolderById(savedId);
  const folder = DriveApp.createFolder("POM Portal Station Task Photos");
  properties.setProperty("STATION_TASK_PHOTO_FOLDER_ID", folder.getId());
  return folder;
}

function validateTaskPhoto(photo) {
  if (!photo || typeof photo !== "object" || photo.mimeType !== "image/jpeg") {
    throw new Error("Task pictures must be JPEG images");
  }
  const name = String(photo.name || "task-picture.jpg").slice(0, 180);
  const data = String(photo.data || "");
  if (!data || data.length > 7000000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(data)) {
    throw new Error("A task picture is empty, too large, or invalid");
  }
  return { name: name, mimeType: "image/jpeg", data: data };
}

function saveTaskPhoto(candidate, taskId) {
  const photo = validateTaskPhoto(candidate);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(photo.data),
    photo.mimeType,
    taskId + ".jpg"
  );
  const file = getTaskPhotoFolder().createFile(blob);
  return { id: file.getId(), name: photo.name, mimeType: photo.mimeType };
}

function trashTaskPhoto(photo) {
  try {
    DriveApp.getFileById(photo.id).setTrashed(true);
  } catch (error) {
    Logger.log("Could not remove station task photo " + photo.id + ": " + error.message);
  }
}

function validateWorkReport(candidate, ss) {
  if (!candidate || typeof candidate !== "object") throw new Error("A work report is required");
  const id = String(candidate.id || "").trim();
  const contractorName = String(candidate.contractorName || "").trim();
  const station = String(candidate.station || "").trim();
  const workName = String(candidate.workName || "").trim();
  const description = String(candidate.description || "").trim();
  const comment = String(candidate.comment || "").trim();
  if (!id || id.length > 100) throw new Error("Invalid work report ID");
  if (!contractorName || contractorName.length > 120) throw new Error("Contractor name is required and must be at most 120 characters");
  if (!station || !readStations(ss).some(item => item.toLowerCase() === station.toLowerCase())) {
    throw new Error("Choose a valid station for this work report");
  }
  if (!workName || workName.length > 160) throw new Error("Work name is required and must be at most 160 characters");
  if (!description || description.length > 5000) throw new Error("Description is required and must be at most 5000 characters");
  if (comment.length > 3000) throw new Error("Comment must be at most 3000 characters");
  if (!Array.isArray(candidate.before) || !Array.isArray(candidate.after)) {
    throw new Error("Before and after photos must be lists");
  }
  return {
    id: id,
    contractorName: contractorName,
    station: station,
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

function readWorkReports(ss, user, company) {
  const sheet = getWorkReportSheet(ss);
  if (sheet.getLastRow() <= 1) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  const companyHash = company ? hashContractorCompany(company.name, company.id) : "";
  return rows.filter(row => hasManagementAccess(user.role) || (
    companyHash
    && String(row[10]) === companyHash
  ))
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

function readWorkReport(ss, reportId, user, includePhotos, company) {
  const sheet = getWorkReportSheet(ss);
  if (sheet.getLastRow() <= 1) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  const row = rows.find(item => String(item[0]) === reportId);
  if (!row) return null;
  if (!hasManagementAccess(user.role) && (
    !company
    || String(row[10]) !== hashContractorCompany(company.name, company.id)
  )) return null;

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

function validateForm(candidate, ss) {
  if (!candidate || typeof candidate !== "object") throw new Error("A form definition is required");
  const name = String(candidate.name || "").trim();
  const id = String(candidate.id || "").trim();
  const type = String(candidate.type || "");
  if (!id || id.length > 100 || !name || name.length > 80) throw new Error("Form ID and a name of at most 80 characters are required");
  if (type !== "Operation" && type !== "Report") throw new Error("Screen type must be Operation or Report");
  const station = String(candidate.station || "All stations").trim();
  if (station !== "All stations" && !readStations(ss).some(item => item.toLowerCase() === station.toLowerCase())) {
    throw new Error("Choose a valid station for this form");
  }
  if (!Array.isArray(candidate.fields) || candidate.fields.length < 1 || candidate.fields.length > 30) {
    throw new Error("A screen must have between 1 and 30 fields");
  }

  const allowedTypes = ["text", "number", "date", "textarea", "checkbox", "photo"];
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
  return { id: id, name: name, type: type, station: station, fields: fields };
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
    if (field.type === "photo") {
      normalized[field.label] = value ? validateTaskPhoto(value) : null;
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
