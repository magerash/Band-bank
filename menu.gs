/**
 * Функция создания меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Банк группы')
    .addItem('📝 Добавить запись', 'showDialog')
    .addToUi();
}