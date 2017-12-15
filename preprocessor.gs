function preprocessor(recordArray, fileName, fileID) {
  switch(recordArray[0]) {
    case 'Affiliate1':
        if (recordArray[8].match(/[$]/) == null) {
        recordArray[8] = '$' + recordArray[8];
      }
        break;
    default:
        break;
}
  recordInjector(recordArray, fileName, fileID);
}
