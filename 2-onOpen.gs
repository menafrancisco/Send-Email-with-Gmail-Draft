function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ADMIN') 
    .addItem('Send emails', 'getDataAndSendEmails')
    .addToUi();
}
