/**
 *
 * Generate Google Docs based on one line record from Google Spreadsheet
 *
 * License: MIT
 *
 * Copyright 2017 daimingzhong, daimingzhong@gmail.com
 */


var TEMPLATEDOC = "GoogleDocTemplate";
var TEMPLATESHEET = "GoogleSheetTemplate";
var templateDocId = [];
var templateSheetId = [];
var mode = [];
var pdfFlag = [];

function onInstall(e) {
    onOpen(e);
}

function onOpen(e) {
    if (e && e.authMode == ScriptApp.AuthMode.NONE) {
        SpreadsheetApp.getUi().createMenu('Invoice').addItem('Google Doc', 'generateDoc').addItem('Google Doc & Pdf', 'generatePdf').addToUi();
    }
    else {
        var properties = PropertiesService.getDocumentProperties();
        var workflowStarted = properties.getProperty('workflowStarted');
        if (workflowStarted) {
            SpreadsheetApp.getUi().createMenu('Invoice').addItem('Google Doc', 'generateDoc').addItem('Google Doc & Pdf', 'generatePdf').addToUi();
        } else {
            SpreadsheetApp.getUi().createMenu('Invoice').addItem('Google Doc', 'generateDoc').addItem('Google Doc & Pdf', 'generatePdf').addToUi();
        }
    }
}

function showPrompt() {
    var keyValueFileIterator = DriveApp.getFilesByName(TEMPLATEDOC);
    var templateFileIterators = DriveApp.getFilesByName(TEMPLATESHEET);
    if(keyValueFileIterator.hasNext()) {
        templateDocId[0] = keyValueFileIterator.next();
    } else {
        SpreadsheetApp.getUi().alert("You don't have GoogleDocTemplate");
        return;
    }
    if(templateFileIterators.hasNext()) {
        templateSheetId[0] = templateFileIterators.next();
    } else {
        SpreadsheetApp.getUi().alert("You don't have GoogleSheetTemplate");
        return;
    }
    var ui = SpreadsheetApp.getUi(); // Same variations.
    var result = ui.prompt(
        mode[0],
        'Enter row number please: (e.g. 2, 5-8, 7-9)',
        ui.ButtonSet.OK_CANCEL);
    var button = result.getSelectedButton();
    var row = result.getResponseText();
    if (button === ui.Button.OK) {
        var rowRange = checkInput(row);
        if (rowRange !== false) {
            for(var i = rowRange[0]; i<= rowRange[1]; i++)
                generateCustomerContract(i);
        }
        else {
            ui.alert('Invalid input: ' + row + ', please retry.');
        }
    } else if (button === ui.Button.CANCEL) {
    } else if (button === ui.Button.CLOSE) {
    }
}

function checkInput(row) {
    if (parseInt(row) == row && parseInt(row) > 0) {
        return [parseInt(row), parseInt(row)];
    }
    else {
        var tmp = row.split('-');
        if (tmp.length === 2 && parseInt(tmp[0]) == tmp[0] && parseInt(tmp[0]) > 0
            &&  parseInt(tmp[1]) == tmp[1] && parseInt(tmp[1]) >= parseInt(tmp[0])) {
            return [parseInt(tmp[0]), parseInt(tmp[1])];
        }
    }
    return false;
}

function generateDoc() {
    mode[0] = 'Google Doc';
    showPrompt();
}

function generatePdf() {
    mode[0] = 'Google Doc & Pdf';
    pdfFlag[0] = true;
    showPrompt();
}


function getKeyValueDefinition() {
    var sheets = SpreadsheetApp.openById(templateSheetId[0].getId());
    var sheet = sheets.getSheets()[0];
    var newFileName = sheets.getSheets()[1].getRange(1,1,1,1).getDisplayValue();
    var data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues();
    var array = [];
    for(i = 0; i < 2; i++) {
        for(var j = 0; j < sheet.getLastRow(); j++) {
            array.push(data[j][i]);
        }
    }
    return [array, newFileName];
}


function getRowAsArray(sheet, rowNumber, columns) {
    var dataRange = sheet.getRange(rowNumber, 1, 1, columns); // start from (row, 1), with block size(1, columns).
    var data = dataRange.getDisplayValues(); // getValue() will calculate dates from 1900, while getDisplayValue() will get the original value
    var array = [];
    for (i in data) {
        var row = data[i];
        for(var l=0; l<row.length; l++) {
            var col = row[l];
            array.push(col);
        }
    }
    return array;
}

function generateCustomerContract(row) {
    var mapping = getKeyValueDefinition()[0];
    row = (typeof row !== 'undefined')? row:3;
    var sheet = SpreadsheetApp.getActiveSheet(); // gettSheets() will return an array of all sheets
    var customerData = getRowAsArray(sheet, row, sheet.getLastColumn());
    var recentDate = Utilities.formatDate(new Date(), "America/New_York", "MM/dd/yyyy");
    var keyword = [];
    var value = [];
    for(var i = 0; i < mapping.length/2; i++) {
        keyword.push(mapping[i]);
    }
    for (i = mapping.length/2; i < mapping.length; i++) {
        value.push(customerData[mapping[i]]);
    }
    keyword.push("Recent_Date");
    value.push(recentDate);
    var newDocName = getFileName(customerData);
    removeDuplicateFile(newDocName);
    var doc =  DocumentApp.openById(templateDocId[0].makeCopy(newDocName).getId()); // return a document
    replaceParagraph(doc, keyword, value);
    doc.saveAndClose();
    if (pdfFlag[0] === true) {
        createPdf(DocumentApp.openById(doc.getId()));
    }
}

function getFileName(customerData) {
    var mapping = getKeyValueDefinition()[0];
    var newFileName = getKeyValueDefinition()[1];
    for(var i = 0; i < mapping.length/2; i++) {
        if(newFileName.indexOf(mapping[i]) !== -1) {
            reg = new RegExp("<<" + mapping[i]+ ">>", 'g');
            newFileName = newFileName.replace(reg, customerData[mapping[mapping.length/2 + i]]);
        }
    }
    return newFileName;
}


function replaceParagraph(doc, keyword, value) {
    var ps = doc.getParagraphs();
    for(i=0; i<ps.length; i++) {
        var p = ps[i];
        var text = p.getText();
        for(var j=0; j<keyword.length; j++) {
            var key = "<<" + keyword[j] + ">>";
            if(text.indexOf(key) !== -1) {
                if(!value[j]) {
                    value[j] = " ";
                }
                var newValue = text.replace(/<<.*>>/g, value[j]);
                p.setText(newValue);
            }
        }
    }
}

function createPdf(doc) {
    var newGoogleDoc = doc.getAs('application/pdf');
    var docFolder = DriveApp.getFileById(doc.getId()).getParents().next().getId();
    var pdfName = doc.getName() + ".pdf";
    removeDuplicateFile(pdfName);
    newGoogleDoc.setName(pdfName);
    var file = DriveApp.createFile(newGoogleDoc);
    var fileId = file.getId();
    moveFileId(fileId, docFolder);
}

function moveFileId(fileId, toFolderId) {
    var file = DriveApp.getFileById(fileId);
    var source_folder = DriveApp.getFileById(fileId).getParents().next();
    var folder = DriveApp.getFolderById(toFolderId);
    folder.addFile(file);
    source_folder.removeFile(file);
}

function removeDuplicateFile(docName) {
    var tmp = DriveApp.getFilesByName(docName);
    while(tmp.hasNext()) {
        var file = tmp.next();
        if (file.getParents().hasNext()) {
            var docFolder = DriveApp.getFolderById(file.getParents().next().getId());
            docFolder.removeFile(file);
        }
    }
}

function quitScript() {

}