/**
 * Основной файл Google Apps Script для работы с диалоговым окном
 * Файл: Code.gs
 */

/**
 * Показать диалоговое окно для добавления записи
 */
function showDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('dialog')
      .evaluate()
      .setWidth(500)
      .setHeight(450)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Добавить новую запись');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка при открытии диалога: ' + error.toString());
  }
}

/**
 * Получить уникальные категории из именованного диапазона data_category
 */
function getCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Используем именованный диапазон data_category
    const namedRange = ss.getRangeByName("data_category");
    
    if (!namedRange) {
      throw new Error('Именованный диапазон data_category не найден');
    }
    
    // Получаем все значения из диапазона
    const categories = namedRange.getValues();
    
    // Фильтруем пустые значения и получаем уникальные
    const uniqueCategories = [...new Set(categories.flat().filter(cat => cat !== ""))];
    
    return uniqueCategories.sort();
  } catch (error) {
    console.log('Ошибка при получении категорий: ' + error.toString());
    return ['Ошибка загрузки категорий'];
  }
}

/**
 * Получить список месяцев на русском языке
 */
function getMonths() {
  return [
    'янв', 'фев', 'мар', 'апр', 'май', 'июн',
    'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'
  ];
}

/**
 * Получить текущий год
 */
function getCurrentYear() {
  return new Date().getFullYear();
}

/**
 * Сохранить новую запись в таблицу
 */
function saveRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Получаем активный лист для записи данных
    const sheet = ss.getActiveSheet();
    
    // Подготавливаем данные для записи
    const currentDate = new Date();
    
    // Преобразуем название месяца в число
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    const monthNumber = monthNames.indexOf(data.month) + 1; // +1 потому что месяцы начинаются с 1
    
    // Создаем дату в формате ДАТА(год;месяц;1)
    const monthDate = new Date(parseInt(data.year), monthNumber - 1, 1); // -1 потому что в JavaScript месяцы начинаются с 0
    
    // Данные для новой строки: [Дата, Категория, Сумма, Месяц]
    const newRow = [
      currentDate,           // Дата - текущая дата
      data.category,         // Категория - выбранная категория
      parseFloat(data.amount), // Сумма - введенная сумма как число
      monthDate             // Месяц - дата в формате ДАТА(год;месяц;1)
    ];
    
    // Находим последнюю строку с данными
    const lastRow = sheet.getLastRow();
    
    // Добавляем новую строку после последней заполненной
    const targetRow = lastRow + 1;
    
    // Записываем данные в строку (предполагаем, что данные начинаются с колонки A)
    sheet.getRange(targetRow, 1, 1, 4).setValues([newRow]);
    
    // Возвращаем успешный результат
    const monthYearDisplay = data.month + ' ' + data.year;
    return {
      success: true,
      message: `Запись успешно добавлена!\nКатегория: ${data.category}\nСумма: ${data.amount}₽\nПериод: ${monthYearDisplay}${data.comment ? '\nКомментарий: ' + data.comment : ''}`
    };
    
  } catch (error) {
    console.log('Ошибка при сохранении записи: ' + error.toString());
    return {
      success: false,
      message: 'Ошибка при сохранении: ' + error.toString()
    };
  }
}

/**
 * Включить файл HTML в основной файл (для работы с CSS и JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
