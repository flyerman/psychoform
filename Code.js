//import 'google-apps-script'

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

    GmailApp.sendEmail(email, report.getName(), 'See answers at ' + report.getUrl(), {});
}


function renderForm(form, response, folder) {

    //var firstName = getResponse(response, 'First Name');
    //var lastName = getResponse(response, 'Last Name');
    //var fileName = 'Form responses for ' + firstName + ' ' + lastName;
    var fileName = 'Form responses';

    // Create a new report
    var report = DocumentApp.create(fileName);
    var repbody = report.getBody();

    // Move the report into the drive folder
    var file = DriveApp.getFileById(report.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

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

    // Save report
    report.saveAndClose();

    //var pdf = report.getAs(MimeType.PDF);
    //folder.createFile(pdf).setName(fileName);

    //GmailApp.sendEmail('foo@example.com', 'answers', 'See answers attached ', {attachments: pdf});

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


// Add the question title
function addQuestionHeader(repbody, question) {
    repbody.appendParagraph('Question: ' + question.getTitle())
           .setHeading(DocumentApp.ParagraphHeading.HEADING3)
           .setBold(true);
}


function addQuestionCheckboxGrid(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var questionRows = question.getRows();
    var questionCols = question.getColumns();
    var answers = responseItem ? responseItem.getResponse() : [[]];

    var tableCells = [];

    // Top header row:  ' ' , col1, col2, col3...
    var header = [''].concat(questionCols);
    tableCells.push(header);

    // Rows:
    for (var i = 0; i < questionRows.length; i++) {
        var row = [questionRows[i]];
        for (var j = 0; j < questionCols.length; j++) {
            var found = false;
            if (answers[i]) {
                for (var a = 0; a < answers[i].length; a++) {
                    if (answers[i][a] == questionCols[j]) {
                        found = true;
                        break;
                    }
                }
            }
            if (found) {
                row.push('✅');
            } else {
                row.push('🔲');
            }
        }
        tableCells.push(row);
    }

    var table = repbody.appendTable(tableCells);
    alignInnerCells(table, questionRows.length, questionCols.length);
}


function alignInnerCells(table, rowCount, colCount) {
    var style = {};
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    for (var i = 0; i < rowCount; i++) {
        for (var j = 0; j < colCount; j++) {
            var cell = table.getCell(i + 1, j + 1);
            var firstChild = cell.getChild(0);
            if (firstChild.getType() == DocumentApp.ElementType.PARAGRAPH) {
                firstChild.asParagraph().setAttributes(style);
            }
        }
    }
}


function addQuestionGrid(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var questionRows = question.getRows();
    var questionCols = question.getColumns();
    var answers = responseItem ? responseItem.getResponse() : [];

    var tableCells = [];

    // Top header row:  ' ' , col1, col2, col3...
    var header = [''].concat(questionCols);
    tableCells.push(header);

    // Rows:
    for (var i = 0; i < questionRows.length; i++) {
        var row = [questionRows[i]];
        for (var j = 0; j < questionCols.length; j++) {
            if (answers[i] == questionCols[j]) {
                row.push('⦿');
            } else {
                row.push('◦');
            }
        }
        tableCells.push(row);
    }

    var table = repbody.appendTable(tableCells);
    alignInnerCells(table, questionRows.length, questionCols.length);
}


function addQuestionDate(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    repbody.appendParagraph('ERROR: date not yet supported').setBold(false);
}


function addQuestionTime(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    repbody.appendParagraph('ERROR: time not yet supported').setBold(false);
}


function addQuestionList(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var choices = question.getChoices();
    var responseText = responseItem ? responseItem.getResponse() : '';
    for (const choice of choices) {
        var choiceText = choice.getValue();
        if (responseText == choiceText) {
            repbody.appendParagraph(choiceText + ' (*)').setBold(true);
        }
        else {
            repbody.appendParagraph(choiceText).setBold(false);
        }
    }
}


function addQuestionCheckbox(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);

    var choices = question.getChoices();
    var responseList = responseItem ? responseItem.getResponse() : [];
    // Add each box tat was checked
    for (const choice of choices) {
        var choiceText = choice.getValue();
        var bullet = '🔲';
        var found = false;
        for (const responseText of responseList) {
            if (responseText == choiceText) {
                bullet = '✅';
                found = true;
                break;
            }
        }
        repbody.appendParagraph(bullet + ' ' + choiceText).setBold(found);
    }
    if (!responseItem) {
        return;
    }
    // Detect and add the 'Other' box
    for (const responseText of responseList) {
        var found = false;
        for (const choice of choices) {
            var choiceText = choice.getValue();
            if (responseText == choiceText) {
                found = true;
                break;
            }
        }
        if (!found) {
            repbody.appendParagraph("✅ Other: " + responseText).setBold(true);
        }
    }
}


function addQuestionMultipleChoice(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);

    var choices = question.getChoices();
    var responseText = responseItem ? responseItem.getResponse() : '';
    var found = false;
    for (const choice of choices) {
        var choiceText = choice.getValue();
        var bullet = '◦';
        var bold = false;
        if (responseText == choiceText) {
            bullet = '⦿';
            found = true;
            bold = true;
        }
        repbody.appendParagraph(bullet + ' ' + choiceText).setBold(bold);
    }
    if (responseItem && !found) {
        repbody.appendParagraph("⦿ Other: " + responseText).setBold(true);
    }
}


function addQuestionText(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var responseText = responseItem ? responseItem.getResponse() : '';
    repbody.appendParagraph("➡ " + responseText).setBold(false);
}


function addQuestionScale(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var responseText = responseItem ? responseItem.getResponse() : '';
    var steps = [];
    for (var i = question.getLowerBound(); i <= question.getUpperBound(); i++) {
        if (i == responseText) {
            steps.push('⦿ ' + i);
        } else {
            steps.push('◦ ' + i);
        }

    }
    repbody.appendParagraph(steps.join('    ')).setBold(false);
}