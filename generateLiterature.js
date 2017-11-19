/**
 * This scrip if developed for the Google Apps Scripts. Its purpose is in
 *  generating a google docs document and a PDF out of this gogole docs docuemnt
 *  with the data from the literature review spreadsheet, prefilled into the
 *  template.
 *
 *
 *
 *
 *
 */


/**
 * Return spreadsheet row content as JS array with all cols in the row elements of the array.
 *
 *
 */
function gerRowData(sheet, row) {
  var maxCol = sheet.getLastColumn();
  var dataRange = sheet.getRange(row, 1, 1, maxCol);
  var data = dataRange.getValues();
  var columns = [];
  for (i in data) {
    var oneRow = data[i];
    //Logger.log("Got row", oneRow);
    for(var l=0; l<maxCol; l++) {
        var col = oneRow[l];
        columns.push(col);
    }
  }
  return columns;
}

/**
 * Move file to another folder
 *
 *
 * @return a new document with a given name from the orignal
 */
function moveFileToAnotherFolder(fileID, targetFolderID) {

  var file = DriveApp.getFileById(fileID);

  // Remove the file from all parent folders
  var parents = file.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(file);
  }

  DriveApp.getFolderById(targetFolderID).addFile(file);

}

/**
 * Duplicates a Google Apps doc in a specific folder
 *
 * @return a new document with a given name from the orignal
 */
function createDocumentDuplicate(sourceId, name, inFolder) {
  var source = DriveApp.getFileById(sourceId);
  var targetFolder = DriveApp.getFolderById(inFolder);
  var newFile = source.makeCopy(name, targetFolder);
  return DocumentApp.openById(newFile.getId());
}

/**
 * Get columns of the multidimentional array
 *
 *
 */
function getCol(matrix, col){
  var column = [];
  for(var i=0; i<matrix.length; i++){
    column.push(matrix[i][col]);
  }
  return column;
}

/**
 * Fill in one template.
 *
 *
 */
function fillOneTemplate(template, pointNames, dataPoints) {
  var templateBody = template.getBody();
  for(i = 0; i < pointNames.length; i++) {
    templateBody = templateBody.replaceText(pointNames[i], dataPoints[i])
  }
  return template.getId();
}

/**
 * Get current data in a format suitable for google google docs.
 *
 * @return Current date and time in a string.
 */
function getCurrentDate() {
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1;
  var yyyy = today.getFullYear();
  var hh = today.getHours();
  var min = today.getMinutes();
  var ss = today.getSeconds();
  var ms = today.getMilliseconds();
  if(dd<10) {dd = '0'+dd}
  if(mm<10) {mm = '0'+mm}
  today = yyyy + '-' + mm + '-' + dd +"."+ hh +"."+ min + "." + ss + "." + ms;
  return(today);
}


/**
 * Get index of the valid for prefilling rows.
 *
 * @return an array of the indexes.
 */
function getIndexesOfValidRows(dataSheet, pattern, validationColNumber) {

  if(typeof(validationColNumber) == "undefined") var validationColNumber = 19;
  if(typeof(pattern) == "undefined") var pattern = 'TRUE';

  // Getting the indexs of all valid rows with data
  var index = [];
  var array = dataSheet.getRange(1, validationColNumber, dataSheet.getMaxRows()).getValues();
  var arrayPart = getCol(array, 0);
  for(var i = 0 ;  i < arrayPart.length ; i ++ ) {
    if(arrayPart[i] == pattern) index.push(i + 1);
  }

  return index;
}

/**
 * Fill in all templates.
 *
 * @return an array of the IDs of all prefilled docs, which are storred in the folder 'tempFolderID'.
 */
function fillAllTemplates(sourceTemplateID, dataSheet, tempFolderID, nameRowNumber, validationColNumber) {

  // Col number with the statement "TRUE" that indicates that data
  // Row number where the names are. If not providede than 1.
  if(typeof(nameRowNumber) == "undefined") var nameRowNumber = 1;
  if(typeof(sourceTemplateID) == "undefined") var sourceTemplateID = LIT_REVIEW_TEMPLATE;
  if(typeof(tempFolderID) == "undefined") var tempFolderID = TEMP_FOLDER;

  var index = getIndexesOfValidRows(dataSheet, "TRUE", validationColNumber);

  // Creating a list for prefilled documents
  var listOfPrefilledDocs = [];

  // Getting the header of data
  var names = gerRowData(dataSheet, nameRowNumber);
  Logger.log("index " + index);

  var prefilRow;
  var newName;
  var newEntry;
  var newEntryID;

  // Here we loop over all valid rows with the iterator i
  for(var j = 0; j < index.length; j++) {

    prefilRow = gerRowData(dataSheet, index[j]);

    // revision name
    newName = String(prefilRow[1]) + ' ' + String(getCurrentDate());

    // Creating template duplicate
    newEntry = createDocumentDuplicate(sourceTemplateID, newName, tempFolderID);

    // Filling in template
    newEntryID = fillOneTemplate(newEntry, names, prefilRow);

    // Saving and closing it
    newEntry.saveAndClose();

    // Adding value to the list
    listOfPrefilledDocs.push(newEntryID);

    Logger.log("Prefilled document: " + newName + "; ID: " + newEntryID);

  }

  Logger.log("listOfPrefilledDocs: " + listOfPrefilledDocs);

  return listOfPrefilledDocs;

}

/**
 * Function for merging all GoogleDocs supplied in the JS aray of IDs into one base doc specified in `baseDocID`.
 */
function mergeGoogleDocs(docIDs, baseDocID) {

  //var docIDs = ['documentID_1','documentID_2','documentID_3','documentID_4'];
  var baseDoc = DocumentApp.openById(baseDocID);

  var body = baseDoc.getActiveSection();

  for (var i = 0; i < docIDs.length; ++i ) {
  Logger.log("Adding documents " + docIDs[i]);
    var otherBody = DocumentApp.openById(docIDs[i]).getActiveSection();
    var totalElements = otherBody.getNumChildren();
    for( var j = 0; j < totalElements; ++j ) {
      var element = otherBody.getChild(j).copy();
      var type = element.getType();
      if( type == DocumentApp.ElementType.PARAGRAPH )
        body.appendParagraph(element);
      else if( type == DocumentApp.ElementType.TABLE )
        body.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )
        body.appendListItem(element);
      else
        throw new Error("Unknown element type: "+type);
    }
    body.appendPageBreak();
  }
}

/**
 * Export one google doc found by ID `docID` in the the PDF in the folder `toFolderID`.
 *
 */
function exportDocToPDF(docID, toFolderID) {
  var mergedDocPDFName = "Full literature review " + String(getCurrentDate()) + ".pdf";
  var mergedDocPDF = DriveApp.createFile(DriveApp.getFileById(docID).getAs('application/pdf'));
  moveFileToAnotherFolder(mergedDocPDF.getId(), toFolderID);
  mergedDocPDF.setName(mergedDocPDFName);
}

/**
 * Main functoin that cobines all abovelisted functions and generate output.
 */
function generateLiterature() {

  Logger.clear();

  var LIT_REVIEW_DATA = "1J-m0ZZuhMPY9sKgxSVM5twEDXn_CCVPCUnIiM59R3cg";
  var LIT_REVIEW_TEMPLATE = "132P1ZdZyutKVNH_pWxCRl_WkgRddM5qsjwiq8LkdC-8";

  // In which Google Drive we toss the target documents
  var TARGET_FOLDER = "1uVjFaC301hN4h1Ha_lh_2vXOJeWR-tob";
  var TEMP_FOLDER = "13_MHJN0TKMPCiSwEyRB8EfgRNwgnZfR8";

  // Data spreadsheet
  var dataFile = SpreadsheetApp.openById(LIT_REVIEW_DATA);
  var dataSheet = dataFile.getSheets()[2];

  // Open template for prefilling
  var listOfPrefilledDocs = fillAllTemplates(LIT_REVIEW_TEMPLATE, dataSheet, TEMP_FOLDER, 1, 19);

  // Create new google document
  var mergedDoc = DocumentApp.create("Full literature review " + String(getCurrentDate()));
  var mergedDocID = mergedDoc.getId();
  moveFileToAnotherFolder(mergedDocID, TARGET_FOLDER);

  // Merge all files in one document
  mergeGoogleDocs(listOfPrefilledDocs, mergedDocID)
  mergedDoc.saveAndClose()

  // Combine all files in one PDF
  exportDocToPDF(mergedDocID, TARGET_FOLDER)
}
