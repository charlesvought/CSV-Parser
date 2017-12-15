//Functional Folders
var folderCSVin = DriveApp.getFolderById('0B5AX8tprgBH8S2lxdm1kY2xNV2s');//drop folder for user-submitted CSV files
var folderLoaderIn = DriveApp.getFolderById('0B5AX8tprgBH8WTc0S255aGdBNzg');//recevies CSV files for processing
var folderRejectedFiles = DriveApp.getFolderById('0B5AX8tprgBH8WGd0NVp1Y0NDZzA');//contains rejected files upon submission
var folderTemplates = DriveApp.getFolderById('0B5AX8tprgBH8Z2ZyVXJHTjVrcFE');//contains templates
var folderTemp = DriveApp.getFolderById('0B5AX8tprgBH8d2VIMWhoRUZ0RjA');//Temp folder
var folderPDFout = DriveApp.getFolderById('0B5AX8tprgBH8TnoyeENLZWNmOEU');//PDF output folder
var folderTestIn = DriveApp.getFolderById('0B5AX8tprgBH8TWdHY2Y3SWpib1k');//Intake folder for test data
var folderTestOut = DriveApp.getFolderById('0B5AX8tprgBH8ZGUzSEFka0dpVzA');//Intake folder for test data
var logSheet = SpreadsheetApp.openById('18KuXT8NbG_MZJkUW2TYNX1EPlF_9HeZF_RbD7uHrlEA').getSheetByName('Log');//output Console Log

//Control intake of production files
var testMode = true;


function dropFolderExtraction() {
 var dropFolder = folderCSVin.getFiles();
 LockService.getScriptLock().tryLock(10000);
  if (LockService.getScriptLock().hasLock() == true && testMode == false) {      
    while (dropFolder.hasNext()) {
     var file = dropFolder.next();
     var fileFormat = file.getMimeType();
     var fileID = file.getId();
     var fileName = file.getName();
     var fileOwner = file.getOwner().getName() + ' (' + file.getOwner().getEmail() + ')';
       if (fileFormat == 'text/csv' && fileName.indexOf('-test') == -1) {
       folderLoaderIn.addFile(DriveApp.getFileById(fileID));
       folderCSVin.removeFile(DriveApp.getFileById(fileID));
       writeLog('File Received [ACCEPTED]: "'+ fileName + '" ' + '(' + fileID + ') ' + 'Format: ' + fileFormat + ' Transmitted by: ' + fileOwner);
       csvParser(fileID, fileName);
       } else {
       folderRejectedFiles.addFile(DriveApp.getFileById(fileID));
       folderCSVin.removeFile(DriveApp.getFileById(fileID));
       writeLog('File Received [REJECTED]: "'+ fileName + '" ' + '(' + fileID + ') ' + 'Format: ' + fileFormat + ' Transmitted by: ' + fileOwner);
     }  
   }
   LockService.getScriptLock().releaseLock();
   folderCSVin.setName('CSV.IN (Online)');
  } else {
   folderCSVin.setName('CSV.IN (Offline)');
 }
}

function csvParser(fileID, fileName) {
   var data = Utilities.parseCsv(DriveApp.getFileById(fileID).getBlob().getDataAsString());
    for (r = 1; r < data.length; r++) {//start parsing on 2nd line of CSV
      var record = [];
        for (p = 0; p < data[r].length; p++) {//start parsing in 1st position
         record.push(data[r][p]);
        }
      preprocessor(record, fileName, fileID);
    }
    writeLog('File Parse [Successful]: "'+ fileName + '" ' + '(' + fileID + ')');
}

function recordInjector(recordArray, fileName, fileID) {
 try {
  var sourceTemplate = folderTemplates.getFilesByName(recordArray[0]).next();
  var newDocument = sourceTemplate.makeCopy(recordArray[0] + '_' + recordArray[2] + '_' + Utilities.formatDate(new Date(), "EST", "MM-dd-yyyy"), folderTemp).getId();
  var template = DocumentApp.openById(newDocument);
   for (i = 0, s = 1; i < recordArray.length; i++, s++) {
     try {
    template.getBody().replaceText('{0' + s.toString() + '}', recordArray[i]);
    template.getHeader().replaceText('{0' + s.toString() + '}', recordArray[i]);
    template.getFooter().replaceText('{0' + s.toString() + '}', recordArray[i]);
     } catch (e) {
    }
   }
  template.saveAndClose();
   if(testMode == true || fileName.indexOf('-test') !== -1) {
    folderTestOut.createFile(template).getAs('application/pdf');
   } else {
    folderPDFout.createFile(template).getAs('application/pdf');
   }  
  writeLog('File Injection [Successful]: "'+ fileName + '" ' + '(' + fileID + ')');
 } catch (e) {
   writeLog('An error has occurred' + e);
 }
}

function manualCSVinTest() {//Grabs Most Recent to least recent...
  var dropFolder = folderCSVin.getFiles();
  if (testMode == true) {      
     var file = dropFolder.next();
     var fileFormat = file.getMimeType();
     var fileID = file.getId();
     var fileName = file.getName();
     var fileOwner = file.getOwner().getName() + ' (' + file.getOwner().getEmail() + ')';
       if (fileFormat == 'text/csv' && fileName.indexOf('-test') !== -1) {
       folderLoaderIn.addFile(DriveApp.getFileById(fileID));
       folderCSVin.removeFile(DriveApp.getFileById(fileID));
       writeLog('TEST File Received [ACCEPTED]: "'+ fileName + '" ' + '(' + fileID + ') ' + 'Format: ' + fileFormat + ' Transmitted by: ' + fileOwner);
       csvParser(fileID, fileName);
       } else {
       writeLog('NO TEST FILE DETECTED');
      }
   }
}

function manualTestIn() {//Grabs Most Recent to least recent...
  var dropFolder = folderTestIn.getFiles();
     var file = dropFolder.next();
     var fileFormat = file.getMimeType();
     var fileID = file.getId();
     var fileName = file.getName();
     var fileOwner = file.getOwner().getName() + ' (' + file.getOwner().getEmail() + ')';
       if (fileFormat == 'text/csv' && fileName.indexOf('-test') !== -1) {
       folderLoaderIn.addFile(DriveApp.getFileById(fileID));
       folderTestIn.removeFile(DriveApp.getFileById(fileID));
       writeLog('TEST File Received [ACCEPTED]: "'+ fileName + '" ' + '(' + fileID + ') ' + 'Format: ' + fileFormat + ' Transmitted by: ' + fileOwner);
       csvParser(fileID, fileName);
       } else {
       writeLog('NO TEST FILE DETECTED');
      }  
}


// https://stackoverflow.com/questions/15414077/merge-multiple-pdfs-into-one-pdf
function mergePDF() {
var files = folderPDFout.getFiles();
var tempBlob = Utilities.newBlob('').setName('blob')
for( var i in files ) {
    var blobs = tempBlob.setBytes(files.next().getBlob().setContentTypeFromExtension().getBytes());
    Logger.log(blobs);
  }
var myPDF = Utilities.newBlob(blobs).setName('bloboutput').getAs("application/pdf");
//folderPDFout.createFile(myPDF);
}


/*
function moveFiles(sourceFolder, destinationFolder) {
  var sourceFiles = sourceFolder.getFiles();
  while (sourceFiles.hasNext()) {
   var file = sourceFiles.next();
   destinationFolder.addFile(file);
   sourceFolder.removeFile(file);
  }
}

function stringCheck(sourceString, searchString) {
  if (sourceString.match(searchString) !== null) {
    var stringEvaluation = true;
    } else {
    var stringEvaluation = false;
  }
  return stringEvaluation
}
*/
function writeLog(string) {
logSheet.insertRowsAfter(1,1);
logSheet.getRange('A2').setValue(Utilities.formatDate(new Date(), "EST", "MM-dd-yyyy'@'HH:mm:ss") + '  ' + string);
}

function initialize() {
return;
}