function onFormSubmit(event) {
    // Locate the drive folder
    var folder = DriveApp.getFolderById(folderId);

    var report = renderForm(event.source, event.response, folder);

    shareReport(report, getResponse(event.response, psychTitle));
}


function shareReport(report, psycho) {
    // Retreive the email
    var sheet = SpreadsheetApp.openById(spreadSheetId);
    var data = sheet.getDataRange().getValues();
    var email = null;
    for (var i = 0; i < data.length; i++) {
        if (data[i][0] == psycho) {
            email = data[i][1];
        }
    }

    if (!email || email == '') {
        Logger.log('Could not find email for ' + psycho);
        return;
    }

    report.addEditor(email);

    //var pdf = report.getAs(MimeType.PDF);
    //folder.createFile(pdf).setName(fileName);
    //GmailApp.sendEmail('foo@example.com', 'answers', 'See answers attached ', {attachments: pdf});

    GmailApp.sendEmail(email, report.getName(), 'See answers at ' + report.getUrl(), {});
}


function renderForm(form, response, folder) {

    var fileName = 'Form responses';

    // Create a new report
    var report = DocumentApp.create(fileName);
    var repbody = report.getBody();

    // Move the report into the drive folder
    var file = DriveApp.getFileById(report.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    repbody.appendParagraph('Answers: ')
           .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Itererate over element in the form
    var formItems = form.getItems();
    for (var i = 0, j = 0; i < formItems.length; i++) {
        var formItem = formItems[i];
        var responseItem = response.getResponseForItem(formItem);

        switch (formItem.getType()) {
            // Ignore unsupported elements
            default:
                break;
            // Supported question types
            case FormApp.ItemType.CHECKBOX:
                addQuestionCheckbox(repbody, formItem.asCheckboxItem(), responseItem);
                break;
            case FormApp.ItemType.CHECKBOX_GRID:
                addQuestionCheckboxGrid(repbody, formItem.asCheckboxGridItem(), responseItem);
                break;
            case FormApp.ItemType.GRID:
                addQuestionGrid(repbody, formItem.asGridItem(), responseItem);
                break;
            case FormApp.ItemType.LIST:
                addQuestionList(repbody, formItem.asListItem(), responseItem);
                break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
                addQuestionMultipleChoice(repbody, formItem.asMultipleChoiceItem(), responseItem);
                break;
            case FormApp.ItemType.SCALE:
                addQuestionScale(repbody, formItem.asScaleItem(), responseItem);
                break;
            case FormApp.ItemType.TEXT:
            case FormApp.ItemType.PARAGRAPH_TEXT:
            case FormApp.ItemType.DATE:
            case FormApp.ItemType.TIME:
                addQuestionText(repbody, formItem, responseItem);
                break;
        }
    }

    // Remove empty line at the beginning of the document
    repbody.removeChild(repbody.getChild(0));

    repbody.appendPageBreak();
    repbody.appendParagraph('Questions: ')
           .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Itererate over element in the form
    var formItems = form.getItems();
    var questionCount = 0;
    for (var i = 0, j = 0; i < formItems.length; i++) {
        var formItem = formItems[i];
        switch (formItem.getType()) {
            // Ignore unsupported elements
            default:
                break;
            // Supported question types
            case FormApp.ItemType.CHECKBOX:
            case FormApp.ItemType.CHECKBOX_GRID:
            case FormApp.ItemType.GRID:
            case FormApp.ItemType.LIST:
            case FormApp.ItemType.MULTIPLE_CHOICE:
            case FormApp.ItemType.SCALE:
            case FormApp.ItemType.TEXT:
            case FormApp.ItemType.PARAGRAPH_TEXT:
            case FormApp.ItemType.DATE:
            case FormApp.ItemType.TIME:
                questionCount++;
                repbody.appendParagraph(questionCount.toString() + '. ' + formItem.getTitle())
                       .setFontSize(8);
                break;
        }
    }

    report.saveAndClose();
    return report;
}


function getResponse(response, questionText) {
    for (const r of response.getItemResponses()) {
        if (!r) {
            continue;
        }
        var formItem = r.getItem();
        if (formItem.getTitle() == questionText) {
            return r.getResponse().toString();
        }
    }
    return '?';
}


function addQuestionCheckboxGrid(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem('Checkbox Grid:').setBold(false);
    var questionRows = question.getRows();
    var answers = responseItem.getResponse();
    for (var i = 0; i < questionRows.length; i++) {
        var row = [questionRows[i]];
        var text = row.length < 35 ? row : row.substr(0, 32) + '...';
        if (answers[i]) {
            text += ": ✅ " + answers[i].join(', ');
        } else {
            text += ": N/A";
        }
        repbody.appendListItem(text)
               .setNestingLevel(1)
               .setIndentStart(72)
               .setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    }

}


function addQuestionGrid(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem('Multiple Choice Grid:').setBold(false);
    var questionRows = question.getRows();
    var answers = responseItem.getResponse();
    for (var i = 0; i < questionRows.length; i++) {
        var row = [questionRows[i]];
        var text = row.length < 35 ? row : row.substr(0, 32) + '...';
        if (answers[i]) {
            text += ": ⦿ " + answers[i];
        } else {
            text += ": N/A";
        }
        repbody.appendListItem(text)
               .setNestingLevel(1)
               .setIndentStart(72)
               .setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    }
}


function addQuestionList(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem("➡ " + responseItem.getResponse()).setBold(false);
}


function addQuestionCheckbox(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem("✅ " + responseItem.getResponse().join(', ')).setBold(false);
}


function addQuestionMultipleChoice(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem("⦿ " + responseItem.getResponse()).setBold(false);
}


function addQuestionText(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem(responseItem.getResponse()).setBold(false);
}


function addQuestionScale(repbody, question, responseItem) {
    if (!responseItem) {
        repbody.appendListItem("N/A");
        return;
    }    
    repbody.appendListItem("⦿ " + responseItem.getResponse()).setBold(false);
}