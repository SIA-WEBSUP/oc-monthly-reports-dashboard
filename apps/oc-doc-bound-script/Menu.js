function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('O&C Reports')
    .addItem('Run Report Tools', 'runReportToolsFromLibrary')
    .addToUi();
}

function runReportToolsFromLibrary() {
  if (typeof OCTools !== 'undefined' && OCTools.runFromDoc) {
    OCTools.runFromDoc();
    return;
  }

  DocumentApp.getUi().alert('Library function OCTools.runFromDoc is not available yet.');
}
