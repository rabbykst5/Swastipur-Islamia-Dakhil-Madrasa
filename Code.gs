var PHOTO_FOLDER_ID = "1-kAEVjWPb5t-3JWVikOqms6bde39NWyF";   
var PAYMENT_FOLDER_ID = "1hf89fvix6aOMIl8-qQ7sq1VuTdg4DQki"; 

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("All_Registrations") || ss.getActiveSheet();

    // 1. Status Update from Admin Panel
    if (data.action === "updateStatus") {
      var rows = sheet.getDataRange().getValues();
      var targetRow = -1;
      var searchPhone = (data.phone || "").toString().trim().replace(/^'+/, '');
      var searchName = (data.name || "").toString().trim();

      for (var i = 1; i < rows.length; i++) {
        var rowPhone = (rows[i][9] || "").toString().trim().replace(/^'+/, '');
        var rowName = (rows[i][2] || "").toString().trim();

        if (rowPhone === searchPhone || (searchName && rowName === searchName)) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow !== -1) {
        sheet.getRange(targetRow, 14).setValue(data.newStatus);
        SpreadsheetApp.flush();
        CacheService.getScriptCache().remove("all_students_cache");
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status updated" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "User not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Duplicate TrxID Strict Check
    var inputTrx = (data.trxId || "").toString().trim().toUpperCase();
    if (inputTrx && inputTrx !== "N/A") {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var trxValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
        for (var k = 0; k < trxValues.length; k++) {
          var existingTrx = (trxValues[k][0] || "").toString().trim().toUpperCase();
          if (existingTrx === inputTrx) {
            return ContentService.createTextOutput(JSON.stringify({
              status: "error",
              message: "এই TrxID (" + inputTrx + ") দিয়ে ইতিপূর্বে রেজিস্ট্রেশন সম্পন্ন হয়েছে! নতুন ট্রানজেকশন আইডি দিন।"
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }

    // Header row check
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "রেজিস্ট্রেশন তারিখ ও সময়", "ছবি", "নাম", "পিতার নাম", "মাতার নাম",
        "ব্যাচ", "দাখিল/এসএসসি পাসের সন", "বর্তমান পেশা", "বর্তমান ঠিকানা",
        "মোবাইল নাম্বার", "পেমেন্ট মাধ্যম", "TrxID", "পেমেন্ট স্লিপ লিঙ্ক", "স্ট্যাটাস", "ছবির লিঙ্ক"
      ]);
    }

    // 3. Save Student Photo
    var photoFolder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var photoFile = Utilities.base64Decode(data.photoData.split(",")[1]);
    var photoName = "Photo_" + data.name.replace(/\s+/g, '_') + "_" + data.phone + ".jpg";
    var photoBlob = Utilities.newBlob(photoFile, data.photoType || "image/jpeg", photoName);
    var photoDriveFile = photoFolder.createFile(photoBlob);
    var photoId = photoDriveFile.getId();
    var photoViewLink = photoDriveFile.getUrl();
    var photoFormula = '=IMAGE("https://lh3.googleusercontent.com/d/' + photoId + '", 1)';

    // 4. Save Payment Slip
    var payProofLink = "N/A";
    if (data.payProofData) {
      var payFolder = DriveApp.getFolderById(PAYMENT_FOLDER_ID);
      var proofFile = Utilities.base64Decode(data.payProofData.split(",")[1]);
      var proofName = "PaySlip_" + data.phone + "_" + (data.trxId || "manual") + ".jpg";
      var proofBlob = Utilities.newBlob(proofFile, data.payProofType || "image/jpeg", proofName);
      var proofDriveFile = payFolder.createFile(proofBlob);
      payProofLink = proofDriveFile.getUrl();
    }

    var formattedPhone = (data.phone || "").toString().trim();
    if (formattedPhone.length === 10 && formattedPhone.startsWith('1')) {
      formattedPhone = '0' + formattedPhone;
    }
    var phoneText = "'" + formattedPhone;

    sheet.appendRow([
      new Date().toLocaleString('bn-BD'),
      photoFormula,
      data.name,
      data.fatherName,
      data.motherName,
      data.batch,
      data.passingYear,
      data.profession,
      data.address,
      phoneText,
      data.payMethod,
      inputTrx || "N/A",
      payProofLink,
      "Pending",
      photoViewLink
    ]);

    CacheService.getScriptCache().remove("all_students_cache");

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "রেজিস্ট্রেশন ও পেমেন্ট তথ্য সফলভাবে জমা হয়েছে!"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("All_Registrations") || ss.getActiveSheet();

    // Live TrxID Verification Endpoint
    if (e && e.parameter && e.parameter.checkTrx) {
      var checkTrx = e.parameter.checkTrx.toString().trim().toUpperCase();
      var lastRow = sheet.getLastRow();
      var isDuplicate = false;

      if (lastRow > 1) {
        var trxValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
        for (var k = 0; k < trxValues.length; k++) {
          var existingTrx = (trxValues[k][0] || "").toString().trim().toUpperCase();
          if (existingTrx === checkTrx) {
            isDuplicate = true;
            break;
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        isDuplicate: isDuplicate
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Normal Admin Load with Cache Optimization
    var cache = CacheService.getScriptCache();
    var cachedData = cache.get("all_students_cache");

    if (cachedData != null && !(e && e.parameter && e.parameter.nocache)) {
      return ContentService.createTextOutput(cachedData)
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getDataRange().getValues();
    if (!rows || rows.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var dataRows = rows.slice(1);
    var studentList = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!row[2] && !row[9]) continue;

      var p = (row[9] || "").toString().trim().replace(/^'+/, '');
      if (p.length === 10 && p.startsWith('1')) p = '0' + p;

      studentList.push({
        timestamp: (row[0] || "").toString(),
        name: (row[2] || "").toString(),
        fatherName: (row[3] || "").toString(),
        motherName: (row[4] || "").toString(),
        batch: (row[5] || "").toString(),
        passingYear: (row[6] || "").toString(),
        profession: (row[7] || "").toString(),
        address: (row[8] || "").toString(),
        phone: p,
        payMethod: (row[10] || "").toString(),
        trxId: (row[11] || "").toString(),
        payProofLink: (row[12] || "").toString(),
        status: (row[13] || "Pending").toString(),
        photoLink: (row[14] || "").toString()
      });
    }

    var jsonString = JSON.stringify(studentList);
    cache.put("all_students_cache", jsonString, 120);

    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}