//import 'google-apps-script'

function onFormSubmit(event) {
    
    // Locate the drive folder
    var folder = DriveApp.getFolderById('1l1PNYtABFmZyZTI-4gqLGNyg1OwNMqm7');

    // Create a new report
    var report = DocumentApp.create('Report assessment');
    var repbody = report.getBody();

    // Move the report into the drive folder
    var file = DriveApp.getFileById(report.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // Reteive the questions from the form
    var form = event.source;
    var formItems = form.getItems();

    // Itererate over element in the form
    for (var i = 0, j = 0; i < formItems.length; i++) {
        var formItem = formItems[i];
        var formType = formItem.getType();

        switch (formItem.getType()) {

            // For supported question types
            case FormApp.ItemType.CHECKBOX:
            case FormApp.ItemType.CHECKBOX_GRID:
            case FormApp.ItemType.GRID:
            case FormApp.ItemType.LIST:
            case FormApp.ItemType.MULTIPLE_CHOICE:
            case FormApp.ItemType.PARAGRAPH_TEXT:
            case FormApp.ItemType.TEXT:
                addQuestion(repbody, formItem, event.response);

            // For other elements 
            default:
                break;
        }
    }

    // Save report
    report.saveAndClose();
}


// Add a question title to the report
function addQuestion(repbody, formItem, response) {

    // Add the question title
    repbody.appendParagraph('');
    repbody.appendParagraph('Question: ' + formItem.getTitle());

    switch (formItem.getType()) {
        
        // case FormApp.ItemType.CHECKBOX:
        //     break;
        // case FormApp.ItemType.CHECKBOX_GRID:
        //     break;
        // case FormApp.ItemType.GRID:
        //     break;
        // case FormApp.ItemType.LIST:
        //     break;
        // case FormApp.ItemType.MULTIPLE_CHOICE:
        //     break;
        
        // Short answer type
        case FormApp.ItemType.PARAGRAPH_TEXT:
        case FormApp.ItemType.TEXT: {
            repbody.appendParagraph("Answer: " + response.getResponseForItem(formItem).getResponse());            
            break;
        }

        default: {
            repbody.appendParagraph("ERROR: Unsupported question type.");
            break;
        }
    }

}