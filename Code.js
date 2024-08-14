// Code.gs
function doGet() {
  var output = HtmlService.createHtmlOutputFromFile('index').setTitle('BDT- Cuttingplan Upload');
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

function getListProjectCode() {
  let sheet = SpreadsheetApp.openById('16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo').getSheetByName('project_register');
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();

  let listProject = [];

  // เพิ่มข้อมูลจาก Google Sheets เข้าไปใน dropdown
  for (let i = 1; i < values.length; i++) {
    listProject.push(`${values[i][1]} ${values[i][2]}`);
  }
  Logger.log(listProject);
  return listProject;
}


function getListZoneName(projectCode) {
  let sheet = SpreadsheetApp.openById('16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo').getSheetByName('zone_register');
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();

  let listZone = values.filter(zone => zone[1] === projectCode).map(zone => zone[2]);

  Logger.log(listZone);
  return listZone;
}

function getLastVersion(projectCode , zoneName) {
  let sheet = SpreadsheetApp.openById('16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo').getSheetByName('transection_upload');
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();

  let versions = values.filter(transection => transection[1] === projectCode && transection[2] === zoneName).map(transection => transection[4]);

  let lastVersion = versions.length > 0 ? Math.max(...versions) : 0;
  
  lastVersion += 1;
  Logger.log(lastVersion);
  return lastVersion;
}


function getListCategoryName() {
  let sheet = SpreadsheetApp.openById('16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo').getSheetByName('category');
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();

  let listCategory = [];


  for (let i = 1; i < values.length; i++) {
    listCategory.push(values[i][1]);
  }
  Logger.log(listCategory);
  return listCategory;
}


function getListFolders() {
  var folderName = "cutting_plan_backlog(clean)"
  var childFolders = [];
  var folders = DriveApp.getFoldersByName(folderName).next().getFolders();

  while (folders.hasNext()) {
    var folder = folders.next();
    childFolders.push(folder.getName());
  }
  return childFolders;
}



function saveDataToSheet(date, projectCode, zone, category, version, revision , attach_id, status ,zip_id) {
  var sheetId = '16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo';
  var sheetName = 'transection_upload';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  var data = [[date, projectCode, zone, category, version, revision ,attach_id, status,zip_id]];


  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, data[0].length).setValues(data);
}

function getTransactionData() {
  var sheetId = '16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo';
  var sheetName = 'transection_upload';
  var range = 'A2:I';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var data = sheet.getRange(range).getValues();
  dataTask = data.filter(function (row) { return !row.every(function (cell) { return cell === ""; }); });
  Logger.log(dataTask)
  return dataTask;
}

function getSheetCleanData() {
  var sheetId = '16ILScFNVt2zcOVoAPhX_N3r8Vz4-AkS8UnVWr2J0PAk';
  var sheetName = 'cleansing_data';
  var range = 'A2:L';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var data = sheet.getRange(range).getValues();
  dataClean = data.filter(function (row) { return !row.every(function (cell) { return cell === ""; }); });
  Logger.log(dataClean)
  return dataClean;
}





function getFilesInFolder(folderName) {
  var folderIterator = DriveApp.getFoldersByName(folderName);

  // ตรวจสอบว่ามีโฟลเดอร์ที่ตรงกับชื่อหรือไม่
  if (folderIterator.hasNext()) {
    var folder = folderIterator.next();
    var fileList = [];

    // ดึงรายการไฟล์จากโฟลเดอร์
    var files = folder.getFiles();

    // สร้าง Array ของ File ทั้งหมดใน Folder
    while (files.hasNext()) {
      var file = files.next();
      var fileInfo = {
        name: file.getName(),
        url: file.getUrl(),
        content: readFileContent(file.getId())
      };
      fileList.push(fileInfo);
    }

    return fileList;
  } else {
    Logger.log('Folder not found:', folderName);
  }
}


function listFilesInFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var fileList = [];
  while (files.hasNext()) {
    var file = files.next();
    fileList.push({ name: file.getName(), id: file.getId(), type: file.getMimeType() });
  }
  return fileList;
}


function readFileContent(fileId) {
  // ดึงข้อมูลจากไฟล์ .txt โดยใช้ File ID
  var file = DriveApp.getFileById(fileId);
  var fileContent = file.getBlob().getDataAsString();
  return fileContent;
}




function reqCleanData(zip_id, project_name, zone_name, category, version , revision) {
  // var apiUrl = 'https://us-central1-cuttingplan.cloudfunctions.net/cuttingplan-process';
  var apiUrl = 'https://asia-southeast1-ssi-bdt.cloudfunctions.net/bdt_cuttingplan_etl_v3';
  var payload = {
    file_id: zip_id,
    project_name: project_name,
    zone_name: zone_name,
    category: category,
    version: version,
    revision: revision
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  Logger.log(response.getContentText());
  return response.getContentText();
}



// ฟังก์ชันเรียกดึงไฟล์จาก Google Drive
function getFilesFromDrive() {
  var folderName = "test_partlist";
  var folders = DriveApp.getFoldersByName(folderName);
  var fileList = [];

  // ตรวจสอบว่ามี Folder หรือไม่
  if (folders.hasNext()) {
    var folder = folders.next();
    var files = folder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      var fileContent = readFileContent(file.getId());
      fileList.push({ id: file.getId(), name: file.getName(), type: 'file', content: fileContent });
    }

    Logger.log('File List in Folder:', fileContent);
  } else {
    Logger.log('Folder not found:', folderName);
  }

  var spreadsheet = SpreadsheetApp.create('Converted Data');

  // ดึง Sheet ที่สร้างมา
  var sheet = spreadsheet.getActiveSheet();

  // แยกข้อมูลจากข้อความ
  var rows = fileContent.split('\n');

  // เขียนข้อมูลลงใน Google Sheet
  for (var i = 0; i < rows.length; i++) {
    var rowData = rows[i].split('\t');
    sheet.appendRow(rowData);
  }

  return fileContent;
}

// ฟังก์ชันอ่านข้อมูลจากไฟล์
function readFileContent(fileId) {
  // ดึงข้อมูลจากไฟล์ .txt โดยใช้ File ID
  var file = DriveApp.getFileById(fileId);
  var fileContent = file.getBlob().getDataAsString();
  return fileContent;
}

// ฟังก์ชันที่เรียกใช้เพื่อดึงไฟล์และอ่านข้อมูล
function processFiles() {
  var fileList = getFilesFromDrive();

  // วนลูปทุกไฟล์และอ่านข้อมูล
  for (var i = 0; i < fileList.length; i++) {
    Logger.log('File Data:', fileList[i].content);
  }
}




function upload(file, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const decodedData = Utilities.base64Decode(file.data);
    const blob = Utilities.newBlob(decodedData, file.mimeType, file.fileName);
    const dataFile = folder.createFile(blob);
    if (dataFile) {
      return "Upload successful";
    }
  } catch (e) {
    Logger.log('Error uploading file:', file.fileName, e.toString());
    return "Error uploading file";
  }
}



function createZipFileToDrive(attach_id, zipData, task_id) {
  try {
    const folder = DriveApp.getFolderById(attach_id);


    var blobs = zipData.map(function (file) {
      const decodedData = Utilities.base64Decode(file.data);
      return Utilities.newBlob(decodedData, file.mimeType, file.fileName);
    });
    var zipFile = Utilities.zip(blobs, `${task_id}.zip`);
    var createdFile = folder.createFile(zipFile);
    return createdFile.getId();
  } catch (e) {
    Logger.log('Error uploading file:', file.fileName, e.toString());
    return "Error uploading file"
  }
}

// function createZipFileToStorage(attach_id, zipData, transactionID) {
//   try {
//     // const folder = DriveApp.getFolderById(attach_id);
//     var blobs = zipData.map(function (file) {
//       const decodedData = Utilities.base64Decode(file.data);
//       return Utilities.newBlob(decodedData, file.mimeType, file.fileName);
//     });
//     var zipFile = Utilities.zip(blobs, `${transactionID}.zip`);


//     const bucketName = 'cutting-plant';  // เปลี่ยนชื่อบัคเก็ตตามที่ต้องการ
//     const fileName = `${transactionID}.zip`;
//     const url = `https://storage.googleapis.com/upload/storage/v1/b/${bucketName}/o?uploadType=media&name=${fileName}`;

//     const options = {
//       method: 'POST',
//       contentType: 'application/zip',
//       payload: zipFile.getBytes(),
//       headers: {
//         "Authorization": "Bearer " + ScriptApp.getOAuthToken()
//       },
//       muteHttpExceptions: true
//     };

//     const response = UrlFetchApp.fetch(url, options);

//     if (response.getResponseCode() === 200) {
//       const responseObject = JSON.parse(response.getContentText());
//       const fileId = responseObject.id;
//       Logger.log('File uploaded successfully. File ID: ' + fileId);
//       return fileId
//     } else {
//       Logger.log('Error uploading file: ' + response.getContentText());
//       return "Error uploading file"
//     }
//   } catch (e) {
//     Logger.log('Error uploading file: ' + e.toString());
//     return "Error uploading file"
//   }
// }


function deleteFolder(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    folder.setTrashed(true); // ย้ายโฟลเดอร์ไปยังถังขยะ
    Logger.log('Folder moved to trash: ' + folder.getName());
  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}


function getOrCreateFolder(folderName) {
  const rootFolderName = 'cutting-plant(file-storage)';
  const mainFolderName = 'cutting_plan_attachment';
  const rootFolders = DriveApp.getFoldersByName(rootFolderName);
  if (!rootFolders.hasNext()) throw new Error('Root folder not found.');
  const rootFolder = rootFolders.next();

  const mainFolders = rootFolder.getFoldersByName(mainFolderName);
  var mainFolder;
  if (!mainFolders.hasNext()) {
    mainFolder = rootFolder.createFolder(mainFolderName);
  } else {
    mainFolder = mainFolders.next();
  }

  const subFolders = mainFolder.getFoldersByName(folderName);
  if (!subFolders.hasNext()) {
    const new_folder = mainFolder.createFolder(folderName);
    return new_folder.getId();
  } else {
    return subFolders.getId();
  }
}


function sendEmailWithFolderAttachment() {
  var recipient = "turtleapisit@gmail.com";
  var subject = "Test Send File";
  var body = "Test Test Body";
  var folderName = "Your Folder Name"; // Replace with the actual folder name

  // Get all folders with the specified name
  var folders = DriveApp.getFoldersByName(folderName);

  // Check if there is a folder with the specified name
  if (folders.hasNext()) {
    var folder = folders.next();
    var files = folder.getFiles();

    // Attach each file in the folder to the email
    while (files.hasNext()) {
      var file = files.next();
      var attachment = file.getBlob();

      // Send the email with the attachment
      GmailApp.sendEmail(recipient, subject, body, {
        attachments: [attachment]
      });
    }
  } else {
    // Folder with the specified name not found
    Logger.log("Folder not found");
  }
}


// function getFile() {
//   try {
//     const bucketName = 'bdt-test-bucket';
//     const fileName = '150-5-1.txt';
//     const url = `https://storage.cloud.google.com/${bucketName}/${fileName}?alt=media`;
//     const response = UrlFetchApp.fetch(url, {
//       method: 'GET',
//       headers: {
//         "Authorization": "Bearer " + ScriptApp.getOAuthToken()
//       }
//     });

//     Logger.log(response.getBlob());

//     return response
//   } catch (error) {
//     console.error('Error getting file:', error);
//   }
// }


// function uploadTest() {
//   try {
//     const bucketName = 'bdt-test-bucket';
//     const fileName = 'test-file-4.txt';
//     const contentType = 'text/plain';
//     const fileContent = 'This is the content of the file.';
//     const url = `https://storage.googleapis.com/upload/storage/v1/b/${bucketName}/o?uploadType=media&name=${fileName}`;

//     const options = {
//       method: 'POST',
//       contentType: contentType,
//       payload: fileContent,
//       headers: {
//         "Authorization": "Bearer " + ScriptApp.getOAuthToken()
//       },
//       muteHttpExceptions: true
//     };

//     const response = UrlFetchApp.fetch(url, options);

//     if (response.getResponseCode() === 200) {
//       const responseObject = JSON.parse(response.getContentText());
//       const fileId = responseObject.id;
//       Logger.log('File uploaded successfully. File ID: ' + fileId);
//     } else {
//       Logger.log('Error uploading file: ' + response.getContentText());
//     }
//   } catch (error) {
//     Logger.log('Error uploading file: ' + error.toString());
//   }
// }



// function saveDataToSheet(data, projectCode, zoneName, version, category , status) {
//   try {
//     var spreadsheet = SpreadsheetApp.openById('1FhAEAvEpKMCLU94TA6KV3DuDGb-RxTNRbvRg4ZrotLw');
//     var sheet = spreadsheet.getSheets()[0];
//     // var spreadsheet = SpreadsheetApp.openById('1s7dvUZOPLoPBgRFAYKVYgStkF-rr8AhM11-nOAzYybI');
//     // var sheet = spreadsheet.getSheetByName(fileType);




//     var result;
//     result = dataFormat(data, projectCode, zoneName, category, version , status);
//     if (result.length > 0) {

//       // ดึงค่าทั้งหมดจากคอลัมน์ B
//       var columnB = sheet.getRange('B:B').getValues();

//       // หาแถวสุดท้ายที่มีค่าจากคอลัมน์ B
//       var lastRow = 0;
//       for (var i = columnB.length - 1; i >= 0; i--) {
//         if (columnB[i][0] !== '') {
//           lastRow = i + 1; // เนื่องจาก index เริ่มจาก 0 จึงต้องบวก 1
//           break;
//         }
//       }

//       var startRow = lastRow + 1;
//       var numberOfRows = result.length;


//       var runningNumbers = [];
//       for (var i = 0; i < numberOfRows; i++) {
//         var numIndex = i - 1
//         runningNumbers.push([startRow + numIndex]); 
//       }


//       sheet.getRange(startRow, 1, numberOfRows, 1).setValues(runningNumbers);


//       sheet.getRange(startRow, 2, result.length, result[0].length).setValues(result);

//       return "Upload Success";
//     } else {
//       return "No data to save.";
//     }

//   } catch (error) {
//     return "Error: " + error.message;
//   }
// }




function dataFormat(data, projectCode, zoneName, version, category) {
  var selectedData = [];
  for (var i = 0; i < data.length; i++) {
    if (i >= 1 && data[i][0] !== "" && data[i][0] !== undefined && data[i][0] !== null) {
      var row = [];
      for (var j = 0; j < data[i].length; j++) {
        row.push(data[i][j]);
      }
      if (row.length > 0) {
        row = [projectCode, zoneName, version, category, ...row];
        selectedData.push(row);
      }
    }
  }
  return selectedData;
}


function deleteByZipID(zipID, folderID) {
  var transactionSheet = '16syZB5m5Er9IDnNf5ph0AKjHQu47xDGjRmIeHvAD_Xo';
  var mopSheet = '1SefomdS9Np6exmdC9Q2DlVWM8m68-hgMREIjsM5MLNg';

  var folder = DriveApp.getFolderById(folderID);
  folder.setTrashed(true);



  copyFilteredData(transactionSheet, 'transection_upload', zipID, 8);
  copyFilteredData(mopSheet, 'nesting', zipID, 1);
  copyFilteredData(mopSheet, 'order_part', zipID, 1);
  copyFilteredData(mopSheet, 'plate_usage', zipID, 1);
  copyFilteredData(mopSheet, 'remnants', zipID, 1);
}



function copyFilteredData(sheetId, sheetName, transactionID, columnIndex) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var filteredData = data.filter(row => row[columnIndex] != transactionID);

  
  sheet.clearContents();

  
  if (filteredData.length > 0) {
    sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
}

