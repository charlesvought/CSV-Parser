//Functional Folders
var folderCSVin = DriveApp.getFolderById('0B5AX8tprgBH8S2lxdm1kY2xNV2s');//drop folder for user-submitted CSV files
var folderLoaderIn = DriveApp.getFolderById('0B5AX8tprgBH8WTc0S255aGdBNzg');//recevies CSV files for processing
var folderRejectedFiles = DriveApp.getFolderById('0B5AX8tprgBH8WGd0NVp1Y0NDZzA');//contains rejected files upon submission
var folderTemplates = DriveApp.getFolderById('0B5AX8tprgBH8Z2ZyVXJHTjVrcFE');//contains templates
var folderTemp = DriveApp.getFolderById('0B5AX8tprgBH8d2VIMWhoRUZ0RjA');//Temp folder
var folderPDFout = DriveApp.getFolderById('0B5AX8tprgBH8TnoyeENLZWNmOEU');//PDF output folder

function dropFolderExtraction() {
 var dropFolder = folderCSVin.getFiles();
  LockService.getScriptLock().tryLock(10000);
  if (LockService.getScriptLock().hasLock() == true) {      
    while (dropFolder.hasNext()) {
     var file = dropFolder.next();
     var fileFormat = file.getMimeType();
     var fileID = file.getId();
     var fileName = file.getName();
     var fileOwner = file.getOwner().getName() + ' (' + file.getOwner().getEmail() + ')';
    if (fileFormat == 'text/csv') {
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
 }
}

function csvParser(fileID, fileName) {
   var data = Utilities.parseCsv(DriveApp.getFileById(fileID).getBlob().getDataAsString());
    for (r = 1; r < data.length; r++) {//start parsing on 2nd line of CSV
      var record = [];
        for (p = 0; p < data[r].length; p++) {//start parsing in 1st position
         record.push(data[r][p]);
        }
      recordInjector(record, fileName, fileID);
    }
    writeLog('File Parse [Successful]: "'+ fileName + '" ' + '(' + fileID + ')');
}

function recordInjector(recordArray, fileName, fileID) {
  var sourceTemplate = folderTemplates.getFilesByName(recordArray[0]).next();
  var newDocument = sourceTemplate.makeCopy(folderTemp).getId();
  var template = DocumentApp.openById(newDocument);
   for (i = 0, s = 1; i < recordArray.length; i++, s++) {
    template.getBody().replaceText('{0' + s.toString() + '}', recordArray[i]);
   }
  template.saveAndClose();
  folderPDFout.createFile(template).getAs('application/pdf');
  writeLog('File Injection [Successful]: "'+ fileName + '" ' + '(' + fileID + ')');
}

function moveFiles(sourceFolder, destinationFolder) {
  var sourceFiles = sourceFolder.getFiles();
  while (sourceFiles.hasNext()) {
   var file = sourceFiles.next();
   destinationFolder.addFile(file);
   sourceFolder.removeFile(file);
  }
}

function writeLog(string) {
var logSheet = SpreadsheetApp.openById('18KuXT8NbG_MZJkUW2TYNX1EPlF_9HeZF_RbD7uHrlEA').getSheetByName('Log');
logSheet.insertRowsAfter(1,1);
logSheet.getRange('A2').setValue(Utilities.formatDate(new Date(), "EST", "MM-dd-yyyy'@'HH:mm:ss") + '  ' + string);
}

function authenticator() {
Logger.clear();
}