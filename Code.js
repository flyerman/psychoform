//import 'google-apps-script'

function onFormSubmit(event) {
    // Locate the drive folder
    var folder = DriveApp.getFolderById('1l1PNYtABFmZyZTI-4gqLGNyg1OwNMqm7');

    renderForm(event.source, event.response, folder);
}

function renderForm(form, response, folder) {
    // Create a new report
    var report = DocumentApp.create('Report assessment');
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

            // For supported question types
            case FormApp.ItemType.CHECKBOX:
                addQuestionCheckbox(repbody, formItem.asCheckboxItem(), responseItem);
                break;
            case FormApp.ItemType.CHECKBOX_GRID:
                addQuestionCheckboxGrid(repbody, formItem, responseItem);
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
            case FormApp.ItemType.PARAGRAPH_TEXT:
            case FormApp.ItemType.TEXT:
                addQuestionText(repbody, formItem, responseItem);
                break;
            case FormApp.ItemType.SCALE:
                addQuestionScale(repbody, formItem.asScaleItem(), responseItem);
                break;
            case FormApp.ItemType.DATE:
                addQuestionDate(repbody, formItem, responseItem);
                break;
            case FormApp.ItemType.TIME:
                addQuestionTime(repbody, formItem, responseItem);
                break;
            // For other elements
            default:
                break;
        }
    }

    // Save report
    report.saveAndClose();
}


// Add the question title
function addQuestionHeader(repbody, question) {
    repbody.appendParagraph('Question: ' + question.getTitle())
           .setHeading(DocumentApp.ParagraphHeading.HEADING3)
           .setBold(true);
}


function addQuestionCheckboxGrid(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    repbody.appendParagraph('ERROR: checkboxgrid not yet supported').setBold(false);
}


function addQuestionGrid(repbody, question, responseItem) {
    addQuestionHeader(repbody, question);
    var questionRows = question.getRows();
    var questionCols = question.getColumns();
    var answers = responseItem ? responseItem.getResponse() : '';

    var table = [];

    // Top header row:  ' ' , col1, col2, col3...
    var header = [''].concat(questionCols);
    table.push(header);

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
        table.push(row);
    }

    repbody.appendTable(table);
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