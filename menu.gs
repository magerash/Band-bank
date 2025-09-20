/**
 * Функция создания меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Банк группы')
    .addItem('📝 Добавить запись', 'showDialog')
    .addToUi();
}


/**
 * Показать диалоговое окно для добавления записи
 */
function showDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('dialog')
      .evaluate()
      .setWidth(600)  // Увеличено с 500
      .setHeight(750) // Увеличено с 450
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Добавить новую запись');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка при открытии диалога: ' + error.toString());
  }
}