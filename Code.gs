var FOLDER_ID = "1-kAEVjWPb5t-3JWVikOqms6bde39NWyF";

function testDriveAccess() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("Folder Found: " + folder.getName());
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("All_Registrations") || ss.getActiveSheet();

    // অ্যাডমিন প্যানেল থেকে স্ট্যাটাস পরিবর্তনের রিকোয়েস্ট
    if (data.action === "updateStatus") {
      var rows = sheet.getDataRange().getValues();
      var targetRow = -1;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][9].toString() === data.phone.toString() && rows[i][0].toString() === data.timestamp.toString()) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow !== -1) {
        sheet.getRange(targetRow, 14).setValue(data.newStatus); // Column N (14) হলো স্ট্যাটাস
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status updated" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "User not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // নতুন রেজিস্ট্রেশন ও হেডার যাচাই
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
        "ছবির লিঙ্ক" // সবার শেষে
      ]);
    }

    var folder = DriveApp.getFolderById(FOLDER_ID);

    // শিক্ষার্থীর প্রোফাইল ছবি ড্রাইভে সেভ করা
    var photoFile = Utilities.base64Decode(data.photoData.split(",")[1]);
    var photoName = "Photo_" + data.name.replace(/\s+/g, '_') + "_" + data.phone + ".jpg";
    var photoBlob = Utilities.newBlob(photoFile, data.photoType || "image/jpeg", photoName);
    var photoDriveFile = folder.createFile(photoBlob);
    var photoId = photoDriveFile.getId();
    var photoViewLink = photoDriveFile.getUrl();
    var photoFormula = '=IMAGE("https://lh3.googleusercontent.com/d/' + photoId + '", 1)';

    // পেমেন্ট স্লিপ/স্ক্রিনশট ড্রাইভে সেভ করা
    var payProofLink = "N/A";
    if (data.payProofData) {
      var proofFile = Utilities.base64Decode(data.payProofData.split(",")[1]);
      var proofName = "PayProof_" + data.phone + "_" + (data.trxId || "manual") + ".jpg";
      var proofBlob = Utilities.newBlob(proofFile, data.payProofType || "image/jpeg", proofName);
      var proofDriveFile = folder.createFile(proofBlob);
      payProofLink = proofDriveFile.getUrl();
    }

    // শীটে সারি যোগ করা (ছবির লিংক সবার শেষে Column O-তে)
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
      data.phone,
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
      return {
        timestamp: row[0],
        name: row[2],
        fatherName: row[3],
        motherName: row[4],
        batch: row[5],
        passingYear: row[6],
        profession: row[7],
        address: row[8],
        phone: row[9],
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