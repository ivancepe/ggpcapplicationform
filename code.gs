/**
 * Main Apps Script handler for login, exam status, scoring, and Kintone webhook.
 */
function doPost(e) {
  // Debug logging (optional, but helpful)
  var debugSheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Debug');
  if (debugSheet) debugSheet.appendRow([new Date(), e ? JSON.stringify(e) : 'No event']);

  var params;
  try {
    if (e && e.postData && e.postData.type === "application/json") {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
    } else {
      params = {};
    }
  } catch (err) {
    params = {};
  }

  if (debugSheet) debugSheet.appendRow([new Date(), 'doPost params', JSON.stringify(params)]);

  // LOGIN HANDLER
  if (params.login) {
    var username = params.username;
    var password = params.password;
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Exam Credentials');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var trimmedHeaders = headers.map(h => h.trim());
    var statusCol = trimmedHeaders.indexOf("Exam Status");
    var emailCol = trimmedHeaders.indexOf("Email");
    var pwdCol = trimmedHeaders.indexOf("Password");
    var recordCol = trimmedHeaders.indexOf("Record Number");

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var email = row[emailCol];
      var pwd = row[pwdCol];
      var recordNumber = row[recordCol];
      if (username === email && password === pwd) {
        // Set "In Progress"
        if (statusCol !== -1) {
          sheet.getRange(i + 1, statusCol + 1).setValue("In Progress");
          // --- Generate new password when status is "In Progress" ---
          if (pwdCol !== -1) {
            var newPassword = generatePassword(12);
            sheet.getRange(i + 1, pwdCol + 1).setValue(newPassword);
          }
        }
        updateKintone(recordNumber, "In Progress", ""); // <-- Sync to Kintone (status only)
        return createCorsResponse({ success: true });
      }
    }
    return createCorsResponse({ success: false });
  }

  // EXAM SUBMISSION HANDLER
  if (params.done || params.timeExpired) {
    var email = params.email;
    var score = params.score;
    var answers = params.answers || "";
    var examVersion = params.examVersion || "";
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Exam Credentials');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var trimmedHeaders = headers.map(h => h.trim());
    var statusCol = trimmedHeaders.indexOf("Exam Status");
    var emailCol = trimmedHeaders.indexOf("Email");
    var scoreCol = trimmedHeaders.indexOf("Score");
    var answersCol = trimmedHeaders.indexOf("Answers");
    var examVersionCol = trimmedHeaders.indexOf("Exam Version");
    var recordCol = trimmedHeaders.indexOf("Record Number");

    // Debug: Log column indexes
    var debugSheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Debug');
    debugSheet.appendRow([new Date(), "Col Indexes", "statusCol", statusCol, "emailCol", emailCol, "scoreCol", scoreCol, "answersCol", answersCol, "examVersionCol", examVersionCol, "recordCol", recordCol]);

    // Defensive check for columns
    if (statusCol === -1 || emailCol === -1 || scoreCol === -1 || recordCol === -1) {
      debugSheet.appendRow([new Date(), "ERROR: One or more columns not found", statusCol, emailCol, scoreCol, recordCol]);
      return createCorsResponse({ success: false, error: "Column not found" });
    }

    // --- Secure backend score calculation ---
    var answersObj = {};
    var examVersionArr = [];
    try { answersObj = JSON.parse(answers); } catch(e) {}
    try { examVersionArr = JSON.parse(examVersion); } catch(e) {}

    var realScore = 0;
    examVersionArr.forEach(function(q) {
      var qKey = q.key || q.qid || q.name || q.questionId;
      var userAnswer = answersObj[qKey];
      if (userAnswer && q.correctAnswer && userAnswer === q.correctAnswer) {
        realScore++;
      }
    });
    score = Number(realScore); // Ensure score is a number
    debugSheet.appendRow([new Date(), "Calculated Score", score]);
    // --- End secure score calculation ---

    // Try to find the matching row by email
    for (var i = 1; i < data.length; i++) {
      debugSheet.appendRow([new Date(), "Row", i, "Email in sheet", data[i][emailCol], "Submitted email", email]);
      if (String(data[i][emailCol]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        // Only update to "Done" if not already "Time Expired"
        debugSheet.appendRow([new Date(), "Match found", "Current status", data[i][statusCol]]);
        if (params.done && data[i][statusCol] !== "Time Expired") {
          sheet.getRange(i + 1, statusCol + 1).setValue("Done");
          updateKintone(data[i][recordCol], "Done", score);
          debugSheet.appendRow([new Date(), "Status set to Done", "Score", score]);
        } else {
          sheet.getRange(i + 1, statusCol + 1).setValue("Time Expired");
          updateKintone(data[i][recordCol], "Time Expired", score);
          debugSheet.appendRow([new Date(), "Status set to Time Expired", "Score", score]);
        }
        sheet.getRange(i + 1, scoreCol + 1).setValue(score);
        if (answersCol !== -1) sheet.getRange(i + 1, answersCol + 1).setValue(answers);
        if (examVersionCol !== -1) sheet.getRange(i + 1, examVersionCol + 1).setValue(examVersion);
        debugSheet.appendRow([new Date(), "Row updated", i + 1]);
        return createCorsResponse({ success: true });
      }
    }
    debugSheet.appendRow([new Date(), "No matching row found for email", email]);
    return createCorsResponse({ success: false, error: "No matching row found" });
  }

  // KINTONE WEBHOOK HANDLER
  if (params.record && params.record.Status) {
    var kintoneStatus = params.record.Status.value;
    var recordNumber = params.record.Record_number ? params.record.Record_number.value : '';
    var kintoneEmail = params.record.Email ? params.record.Email.value : '';
    var fullName = params.record.Full_Name ? params.record.Full_Name.value : '';
    var password = generatePassword(12);

    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Exam Credentials');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var trimmedHeaders = headers.map(h => h.trim());
    var recordCol = trimmedHeaders.indexOf("Record Number");
    var emailCol = trimmedHeaders.indexOf("Email");
    var nameCol = trimmedHeaders.indexOf("Name");
    var pwdCol = trimmedHeaders.indexOf("Password");
    var statusCol = trimmedHeaders.indexOf("Exam Status");

    // Check if record already exists
    var found = false;
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][recordCol]) === String(recordNumber)) {
        found = true;
        rowIdx = i;
        break;
      }
    }

    // Always: If not found, append new row
    if (!found) {
      var newRow = [];
      newRow[recordCol] = recordNumber;
      newRow[emailCol] = kintoneEmail;
      newRow[nameCol] = fullName;
      newRow[pwdCol] = password;
      sheet.appendRow(newRow);
      rowIdx = sheet.getLastRow() - 1;
    }

    // Now only update, never append again
    if (kintoneStatus === 'Exam') {
      if (statusCol !== -1 && rowIdx !== -1) {
        sheet.getRange(rowIdx + 1, statusCol + 1).setValue("Exam Link Sent");
        // Set Date Sent to today
        var dateSentCol = headers.indexOf("Date Sent");
        if (dateSentCol !== -1) {
          sheet.getRange(rowIdx + 1, dateSentCol + 1).setValue(new Date());
        }
        updateKintone(recordNumber, "Exam Link Sent", ""); // Update Kintone field as well
      }
      return createCorsResponse({ success: true });
    }

    if (kintoneStatus === 'Screening') {
      // Update existing row with latest info
      if (emailCol !== -1) sheet.getRange(rowIdx + 1, emailCol + 1).setValue(kintoneEmail);
      if (nameCol !== -1) sheet.getRange(rowIdx + 1, nameCol + 1).setValue(fullName);
      if (pwdCol !== -1) sheet.getRange(rowIdx + 1, pwdCol + 1).setValue(password);

      // Update password in Kintone
      var kintoneDomain = 'vez7o26y38rb.cybozu.com';
      var apiToken = '7DEiGz9DyRxKHoT1xwwvWBBq5k999YGVr1gRhKkh';
      var appId = '1586';
      var url = 'https://' + kintoneDomain + '/k/v1/record.json';
      var payload = {
        app: appId,
        id: recordNumber,
        record: { Password: { value: password } }
      };
      var options = {
        method: 'put',
        contentType: 'application/json',
        headers: { 'X-Cybozu-API-Token': apiToken },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      try {
        UrlFetchApp.fetch(url, options);
      } catch (err) {
        // Optionally log error
      }
      return createCorsResponse({ success: true });
    }

    return createCorsResponse({ success: false, error: 'No action taken' });
  }

  // TIME EXPIRED HANDLER
  if (params.timeExpired) {
    var email = params.email;
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Exam Credentials');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var trimmedHeaders = headers.map(h => h.trim());
    var statusCol = trimmedHeaders.indexOf("Exam Status");
    var emailCol = trimmedHeaders.indexOf("Email");
    var recordCol = trimmedHeaders.indexOf("Record Number");

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        sheet.getRange(i + 1, statusCol + 1).setValue("Time Expired");
        var recordNumber = data[i][recordCol];
        updateKintone(recordNumber, "Time Expired", "");
        return createCorsResponse({ success: true });
      }
    }
    return createCorsResponse({ success: false, error: "No matching row found" });
  }

  // Interview Request Handler
  if (params.type === "INTERVIEW_REQUEST" || (params.interview && params.interview === "true")) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Interview Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Interview Requests');
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'RecordId', 'Status', 'Meetlink', 'Error', 'InterviewDate', 'InterviewTime']);
    }
    sheet.appendRow([
      new Date(),
      params.name || (params.record && params.record.Full_Name && params.record.Full_Name.value) || '',
      params.email || (params.record && params.record.Email && params.record.Email.value) || '',
      params.recordId || (params.record && params.record.Record_number && params.record.Record_number.value) || '',
      'Pending',
      '',
      '',
      params.interviewDateTime || (params.record && params.record.Date_and_time && params.record.Date_and_time.value) || '',
      ''
    ]);
    var debugSheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Debug');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'Writing Interview Request', JSON.stringify(params)]);
    }
    return createCorsResponse({ success: true });
  }

  // Status Update Handler
  if (
    params.type === "UPDATE_STATUS" &&
    params.record &&
    params.record.Status &&
    params.record.Status.value === "Initial Interview"
  ) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Interview Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Interview Requests');
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'RecordId', 'Status', 'Meetlink', 'Error', 'InterviewDate', 'InterviewTime']);
    }
    sheet.appendRow([
      new Date(),
      params.record.Full_Name ? params.record.Full_Name.value : '',
      params.record.Email ? params.record.Email.value : '',
      params.record.Record_number ? params.record.Record_number.value : '', // <-- This is correct for your payload
      'Pending',
      '',
      '',
      params.record.Date_and_time ? params.record.Date_and_time.value : '',
      ''
    ]);
    return createCorsResponse({ success: true });
  }

  // Final Interview Email Handler
  if (params.type === 'SEND_FINAL_INTERVIEW_EMAIL' && params.name && params.email) {
    try {
      sendFinalInterviewEmail(params.name, params.email);
      return createCorsResponse({ success: true });
    } catch (err) {
      return createCorsResponse({ success: false, error: err.toString() });
    }
  }

  // Final Interview Request Handler
  if (params.type === 'FINAL_INTERVIEW_EMAIL_REQUEST' && params.name && params.email && params.finalDateTime) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Final Interview Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Final Interview Requests');
      sheet.appendRow(['Timestamp', 'Full_Name', 'Email', 'Position', 'Final_Date_and_time', 'Status']);
    }
    sheet.appendRow([new Date(), params.name, params.email, params.position || '', params.finalDateTime, 'Pending']);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- New Application Submission Handler ---
  if (params.type === "NEW_APPLICATION" && params.name && params.email) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Applications');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Applications');
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Position', 'Phone', 'Birthdate', 'Address', 'Education', 'Course', 'ExpectedSalary', 'Availability']);
    }
    sheet.appendRow([
      new Date(),
      params.name,
      params.email,
      params.position || '',
      params.phone || '',
      params.birthdate || '',
      params.address || '',
      params.education || '',
      params.course || '',
      params.expectedSalary || '',
      params.availability || ''
    ]);
    // Send confirmation email
    sendApplicationConfirmationEmail(params.name, params.email);
    return createCorsResponse({ success: true });
  }

  return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Unknown request" })).setMimeType(ContentService.MimeType.JSON);
}

// Helper for CORS JSON response
function createCorsResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper function to generate a random password
function generatePassword(length) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()';
  var password = '';
  for (var i = 0; i < length; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

// Dummy doGet and doOptions for CORS preflight
function doGet(e) {
  // Final Interview GET handler (existing)
  if (e && e.parameter && e.parameter.finalInterview) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Final Interview Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Final Interview Requests');
      sheet.appendRow(['Timestamp', 'Full_Name', 'Email', 'Position', 'Final_Date_and_time', 'Status']);
    }
    sheet.appendRow([
      new Date(),
      e.parameter.name || '',
      e.parameter.email || '',
      e.parameter.position || '',
      e.parameter.finalDateTime || '',
      'Pending'
    ]);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  // Interview Request GET handler (NEW)
  if (e && e.parameter && e.parameter.interview) {
    var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Interview Requests');
    if (!sheet) {
      sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').insertSheet('Interview Requests');
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'RecordId', 'Status', 'Meetlink', 'Error', 'InterviewDate', 'InterviewTime']);
    }
    sheet.appendRow([
      new Date(),
      e.parameter.name || '',
      e.parameter.email || '',
      e.parameter.recordId || '',
      'Pending',
      '',
      '',
      e.parameter.interviewDateTime || '',
      ''
    ]);
    // Optional: log to Debug sheet
    var debugSheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Debug');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'Writing Interview Request (GET)', JSON.stringify(e.parameter)]);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var sheetName = "Exam Credentials";
    var ss = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo');
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return createCorsResponse({ success: false, error: "Sheet not found" });
    }
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return createCorsResponse({ success: false, error: "No data found" });
    }
    var headers = data[0];
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      results.push(row);
    }
    return createCorsResponse({ success: true, results: results });
  } catch (err) {
    return createCorsResponse({ success: false, error: err.toString() });
  }
}

function doOptions(e) {
  return createCorsResponse({});
}


/**
 * Daily trigger: expire links and notify HR/applicant if needed.
 */
function expireExamLinks() {
  var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Exam Credentials');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var statusCol = headers.indexOf("Exam Status");
  var dateSentCol = headers.indexOf("Date Sent");
  var emailCol = headers.indexOf("Email");
  var fullNameCol = headers.indexOf("Name"); // Use "Name" as per your header

  if (statusCol === -1 || dateSentCol === -1 || emailCol === -1) return;

  var now = new Date();
  var hrEmail = "ivangolosinda2@gmail.com"; // <-- Replace with your HR email

  for (var i = 1; i < data.length; i++) {
    var status = data[i][statusCol];
    var dateSent = data[i][dateSentCol];
    var applicantEmail = data[i][emailCol];
    var fullName = fullNameCol !== -1 ? data[i][fullNameCol] : "Applicant";

    if (status === "Exam Link Sent" && dateSent) {
      var sentDate = new Date(dateSent);
      var diffHours = (now - sentDate) / (1000 * 60 * 60); // difference in hours

      // Notify at 48 hours after sent
      var notifyHours = 48;
      var expireHours = 72;
      var notifiedCol = headers.indexOf("Notified");
      if (notifiedCol === -1) {
        sheet.getRange(1, headers.length + 1).setValue("Notified");
        notifiedCol = headers.length;
      }
      var notified = data[i][notifiedCol];

      if (diffHours >= notifyHours && diffHours < expireHours && !notified) {
        // Send email to applicant
        MailApp.sendEmail({
          to: applicantEmail,
          subject: "Exam Link Expiry Notice",
          htmlBody: "Dear " + fullName + ",<br><br>Your exam link will expire in less than 24 hours. Please complete your exam as soon as possible.<br><br>Thank you."
        });
        // Send email to HR
        MailApp.sendEmail({
          to: hrEmail,
          subject: "Applicant Exam Link Expiry Notice",
          htmlBody: "The exam link for applicant <b>" + fullName + "</b> (" + applicantEmail + ") will expire in less than 24 hours and is still pending."
        });
        // Mark as notified
        sheet.getRange(i + 1, notifiedCol + 1).setValue("Yes");
      }

      // Expire after 72 hours
      if (diffHours >= expireHours) {
        sheet.getRange(i + 1, statusCol + 1).setValue("Link Expired");
        // Generate a new password and update it in the sheet only
        var pwdCol = headers.indexOf("Password");
        if (pwdCol !== -1) {
          var newPassword = generatePassword(12);
          sheet.getRange(i + 1, pwdCol + 1).setValue(newPassword);
        }
        // Also update Kintone Exam_Status to "Link Expired"
        var recordCol = headers.indexOf("Record Number");
        var recordNumber = data[i][recordCol];
        updateKintone(recordNumber, "Link Expired", "");
      }
    }
  }
}

// Separate function to update Kintone record
function updateKintone(recordNumber, examStatusValue, scoreValue, interviewLink) {
  if (!recordNumber) return;
  var kintoneDomain = 'vez7o26y38rb.cybozu.com';
  var apiToken = '7DEiGz9DyRxKHoT1xwwvWBBq5k999YGVr1gRhKkh';
  var appId = '1586';
  var url = 'https://' + kintoneDomain + '/k/v1/record.json';
  var payload = {
    app: appId,
    id: recordNumber,
    record: {}
  };
  if (examStatusValue !== undefined && examStatusValue !== "") payload.record.Exam_Status = { value: examStatusValue };
  if (scoreValue !== undefined && scoreValue !== "") payload.record.Score = { value: scoreValue };
  if (interviewLink !== undefined && interviewLink !== "") payload.record.Interview_Link = { value: interviewLink }; // <-- Add this line
  var options = {
    method: 'put',
    contentType: 'application/json',
    headers: { 'X-Cybozu-API-Token': apiToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var debugSheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Debug');
  try {
    var response = UrlFetchApp.fetch(url, options);
    var respText = response.getContentText();
    if (debugSheet) {
      debugSheet.appendRow([new Date(), "Kintone update", recordNumber, examStatusValue, scoreValue, interviewLink, respText]);
    }
  } catch (err) {
    if (debugSheet) {
      debugSheet.appendRow([new Date(), "Kintone update error", recordNumber, examStatusValue, scoreValue, interviewLink, err.toString()]);
    }
  }
}

/**
 * Process interview requests: schedule interviews and send email invites.
 */
function processInterviewRequests() {
  var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Interview Requests');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // No data

  var headers = data[0];
  var nameCol = headers.indexOf("Name");
  var emailCol = headers.indexOf("Email");
  var recordIdCol = headers.indexOf("RecordId");
  var statusCol = headers.indexOf("Status");
  var meetLinkCol = headers.indexOf("Meetlink");
  var errorCol = headers.indexOf("Error");
  var interviewDateCol = headers.indexOf("InterviewDate");
  var interviewTimeCol = headers.indexOf("InterviewTime");

  // Check for error and status columns
  if (errorCol === -1 || statusCol === -1) {
    throw new Error('Error or Status column not found in Interview Requests sheet!');
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][statusCol] === 'Pending') {
      var name = data[i][nameCol];
      var email = data[i][emailCol];
      var recordId = data[i][recordIdCol];
      var interviewDateTime = data[i][interviewDateCol]; // ISO string from Kintone

      var interviewDate = '';
      var interviewTime = '';

       // Add this check here
    if (!name || !email || !recordId || !interviewDateTime) {
      sheet.getRange(i + 1, errorCol + 1).setValue('Missing required field(s)');
      sheet.getRange(i + 1, statusCol + 1).setValue('Error');
      continue;
    }

      try {
        if (interviewDateTime) {
          var iso = String(interviewDateTime);
          if (!iso.endsWith('Z') && !iso.match(/[+-]\d{2}:\d{2}$/)) {
            iso += 'Z';
          }
          var dt = new Date(iso);
          var tz = Session.getScriptTimeZone();
          interviewDate = Utilities.formatDate(dt, tz, "MMMM d, yyyy");
          interviewTime = Utilities.formatDate(dt, tz, "h:mm a");
        } else {
          // Log missing date/time error
          sheet.getRange(i + 1, errorCol + 1).setValue('InterviewDateTime is empty');
          continue;
        }

        var calendarId = CalendarApp.getDefaultCalendar().getId();
        var startTime = new Date(Date.now() + 3600 * 1000); // 1 hour from now
        var endTime = new Date(startTime.getTime() + 30 * 60000); // 30 min

        var event = {
          summary: `Initial Interview with ${name}`,
          description: `Dear ${name},\n\nYour interview is scheduled.\n\nGoogle Meet: (see below)\n\nBest regards,\nHR Team`,
          start: { dateTime: startTime.toISOString() },
          end: { dateTime: endTime.toISOString() },
          attendees: [{ email: email }],
          conferenceData: {
            createRequest: {
              requestId: Utilities.getUuid(),
              conferenceSolutionKey: { type: "hangoutsMeet" }
            }
          }
        };

        var createdEvent = Calendar.Events.insert(event, calendarId, { conferenceDataVersion: 1 });
        var meetLink = createdEvent.conferenceData.entryPoints[0].uri;

        sheet.getRange(i + 1, statusCol + 1).setValue('Done');
        sheet.getRange(i + 1, meetLinkCol + 1).setValue(meetLink);
        //Replace with HR Email 
        MailApp.sendEmail({
          to: email,
          subject: `Initial Interview Schedule and Google Meet Link`,
          htmlBody: `Dear ${name},<br><br>
            Congratulations! You have passed the pre-employment examination and
            you are invited to an initial interview.<br><br>
            <b>Date:</b> ${interviewDate}<br>
            <b>Time:</b> ${interviewTime}<br>
            <b>Google Meet Link:</b> <a href="${meetLink}">${meetLink}</a><br><br>
            Please ensure you have a stable internet connection and a quiet environment for the interview.<br>
            If you have any questions or need to reschedule, please contact us at <a href="mailto:ivangolosinda2@gmail.com">ivangolosinda2@gmail.com</a>.<br>
            Meeting link is auto generated. Please join at the specified time in this email.<br><br>
            Best regards,<br>
            GGPC Recruitment Team`
        });
        updateKintone(recordId, '', '', meetLink);
      } catch (err) {
        sheet.getRange(i + 1, statusCol + 1).setValue('Error');
        sheet.getRange(i + 1, errorCol + 1).setValue(err.toString());
      }
    }
  }
}

function processFinalInterviewRequests() {
  var sheet = SpreadsheetApp.openById('168CuUMDWgHveEqcI8GxKgXHz3Xy9pjQUcnWu-CH2bKo').getSheetByName('Final Interview Requests');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] !== 'Pending') continue; // Status column
    var name = data[i][1];
    var email = data[i][2];
    var position = data[i][3] || 'System Developer';
    var finalDateTimeRaw = data[i][4] || '';
var finalDateTimeFormatted = '';
if (finalDateTimeRaw) {
  var dt = new Date(finalDateTimeRaw);
  var tz = Session.getScriptTimeZone();
  finalDateTimeFormatted =
    Utilities.formatDate(dt, tz, "MMMM d, yyyy") + " at " +
    Utilities.formatDate(dt, tz, "h:mm a");
}
sendFinalInterviewEmail(name, email, position, finalDateTimeFormatted);
    sheet.getRange(i + 1, 6).setValue('Sent');
  }
}

function sendFinalInterviewEmail(name, email, position, finalDateTime) {
  MailApp.sendEmail({
    to: email,
    subject: "Congratulations! You are invited for the Final Interview",
    htmlBody: `Dear ${name},<br><br>
      Congratulations! You have passed the initial interview and are now invited for the <b>Final Interview</b> here in our office.<br><br>
      Here are the details:<br><br>
      Position: ${position}<br>
      Date & Time: ${finalDateTime}<br>
      Location: Lot 1 Blk. 7 Circuit St. Light Industry & Science Park of the Philippines 1 (LISP 1) <br>
      Diezmo, Cabuyao City, Laguna, 4025, Philippines<br>
      <a href="https://maps.app.goo.gl/L3CdZPCKpUF2uPfj6" target="_blank">View Map</a><br><br>

      If you are unfamiliar with the location, please use the link above to find directions.<br>
      Please ensure you arrive on time.<br>

      To facilitate access to our premises, please bring the following documents:<br>

      <ul>
        <li>Valid ID (e.g., government-issued ID)</li>
        <li>Copy of your resume</li>
        <li>A ballpoint pen</li>
      </ul>

      Note: Wearing sleeveless shirts, ripped jeans, and sandals is not allowed.<br>
      Please dress appropriately for the interview.<br><br>

      
      If you have any questions or need to reschedule, please contact us immediately.<br><br>
      We look forward to meeting you in person!<br><br>
      Best regards,<br>
      GGPC Recruitment Team`
  });
}

// Add this function to send confirmation email to applicant
function sendApplicationConfirmationEmail(name, email) {
  MailApp.sendEmail({
    to: email,
    subject: "Application Received",
    htmlBody: `Dear ${name},<br><br>
      Thank you for submitting your application through our online application form.<br><br>
      Our recruitment team will carefully review your submission and contact you if you are shortlisted for the next stage of the selection process.<br><br>
      If you have any questions or need to update your application, feel free to reach out to us at recruitment.ggpc@gmail.com<br><br>
      We appreciate your interest in joining our team and wish you the best of luck!<br><br>
      
      Sincerely,<br>
      GGPC Recruitment Team`
  });
}
