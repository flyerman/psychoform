function onFormSubmit(event) {
    
    // Locate the drive folder
    var folder = DriveApp.getFolderById('1l1PNYtABFmZyZTI-4gqLGNyg1OwNMqm7');

    // Create a new report
    var report = DocumentApp.create('Report assessment');

    // Move the report into the drive folder
    var file = DriveApp.getFileById(report.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
}
