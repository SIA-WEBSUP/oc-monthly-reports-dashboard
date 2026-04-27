const MENU_NAME = 'Report Tools';
const SCRIPT_ID = '1JE4hStZsGvGD7xYv63_ohZKem7zPKYS6Gvf8HLcbBk1PfQLQmXJbrPeK'
const SCRIPT_URL= 'https://script.google.com/macros/library/d/1JE4hStZsGvGD7xYv63_ohZKem7zPKYS6Gvf8HLcbBk1PfQLQmXJbrPeK/1'

function installAccessLoggerForDoc(docId) {
  const doc = DocumentApp.openById(docId);

  ScriptApp.newTrigger('logReportOpen')
    .forDocument(doc)
    .onOpen()
    .create();
}

function onOpenMenu() {
  DocumentApp.getUi()
    .createMenu('Report Tools')
    .addItem('Sort', 'sortSections')
    .addItem('Reverse Sort', 'reverseSortSections')
    .addItem('Special Sort', 'specialSortSections')
    .addItem('Desc Month, Asc Year', 'sortDescMonthAscYear')
    .addSeparator()
    .addItem('Add Blank Template for Current Month', 'addBlankTemplateForCurrentMonth')
    .addToUi();
}

function sortSections() {
  DocumentApp.getUi().alert('Sort not implemented yet.');
}

function reverseSortSections() {
  DocumentApp.getUi().alert('Reverse Sort not implemented yet.');
}

function specialSortSections() {
  DocumentApp.getUi().alert('Special sort not implemented yet.');
}

function sortDescMonthAscYear() {
  DocumentApp.getUi().alert('Desc month, asc year sort not implemented yet.');
}

function debugMonthlyTemplateElementsFromMenu() {
  debugListMonthlyTemplateElements();
}