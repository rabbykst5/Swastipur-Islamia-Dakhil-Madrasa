var PHOTO_FOLDER_ID = "1-kAEVjWPb5t-3JWVikOqms6bde39NWyF";   // Profile Photo Folder
var PAYMENT_FOLDER_ID = "1hf89fvix6aOMIl8-qQ7sq1VuTdg4DQki"; // Payment Slip Folder

function testDriveAccess() {
  var photoFolder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  var payFolder = DriveApp.getFolderById(PAYMENT_FOLDER_ID);
  Logger.log("Photo Folder: " + photoFolder.getName());
  Logger.log("Payment Folder: " + payFolder.getName());
}

function doPost(e) {
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
        sheet.getRange(targetRow, 14).setValue(data.newStatus); // Column N (14) = Status
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status updated" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "User not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Header Row check
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "রেজিস্ট্রেশন তারিখ ও সময়",
        "ছবি",
        "নাম",
        "পিতার নাম",
        "মাতার নাম",
        "ব্যাচ",
        "দাখিল/এসএসসি পাসের সন",
        "বর্তমান পেশা",
        "বর্তমান ঠিকানা",
        "মোবাইল নাম্বার",
        "পেমেন্ট মাধ্যম",
        "TrxID",
        "পেমেন্ট স্লিপ লিঙ্ক",
        "স্ট্যাটাস",
        "ছবির লিঙ্ক"
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

    // 4. Save Payment Slip to Separate Folder
    var payProofLink = "N/A";
    if (data.payProofData) {
      var payFolder = DriveApp.getFolderById(PAYMENT_FOLDER_ID);
      var proofFile = Utilities.base64Decode(data.payProofData.split(",")[1]);
      var proofName = "PaySlip_" + data.phone + "_" + (data.trxId || "manual") + ".jpg";
      var proofBlob = Utilities.newBlob(proofFile, data.payProofType || "image/jpeg", proofName);
      var proofDriveFile = payFolder.createFile(proofBlob);
      payProofLink = proofDriveFile.getUrl();
    }

    // Phone number text hishebe rakhle 0 kokhono katbe na
    var formattedPhone = (data.phone || "").toString().trim();
    if (formattedPhone.length === 10 && formattedPhone.startsWith('1')) {
      formattedPhone = '0' + formattedPhone;
    }
    var phoneText = "'" + formattedPhone;

    // 5. Append Row to Sheet
    sheet.appendRow([
      new Date(),
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
      data.trxId || "N/A",
      payProofLink,
      "Pending",
      photoViewLink
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "রেজিস্ট্রেশন ও পেমেন্ট তথ্য সফলভাবে জমা হয়েছে!"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("All_Registrations") || ss.getActiveSheet();
    var rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }

    var dataRows = rows.slice(1);
    var studentList = dataRows.map(function(row) {
      var p = (row[9] || "").toString().trim().replace(/^'+/, '');
      if (p.length === 10 && p.startsWith('1')) {
        p = '0' + p;
      }

      return {
        timestamp: row[0],
        name: row[2],
        fatherName: row[3],
        motherName: row[4],
        batch: row[5],
        passingYear: row[6],
        profession: row[7],
        address: row[8],
        phone: p,
        payMethod: row[10],
        trxId: row[11],
        payProofLink: row[12],
        status: row[13] || "Pending",
        photoLink: row[14]
      };
    });

    return ContentService.createTextOutput(JSON.stringify(studentList))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}