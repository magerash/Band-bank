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
/**
 * Сохранить новую запись в таблицу Банк_группы на листе Операции
 */
// function saveRecord(data) {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
    
//     // Получаем лист "Операции"
//     const sheet = ss.getSheetByName("Операции");
//     if (!sheet) {
//       throw new Error('Лист "Операции" не найден');
//     }
    
//     // Подготавливаем данные для записи
//     const currentDate = new Date();
    
//     // Преобразуем название месяца в число
//     const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
//     const monthNumber = monthNames.indexOf(data.month) + 1;
    
//     // Создаем дату для месяца
//     const monthDate = new Date(parseInt(data.year), monthNumber - 1, 1);
    
//     // Попробуем найти таблицу "Банк_группы" через именованный диапазон
//     let tableRange;
//     try {
//       tableRange = ss.getRangeByName("Банк_группы");
//     } catch (e) {
//       // Если именованный диапазон не найден, ищем таблицу вручную
//       console.log('Именованный диапазон "Банк_группы" не найден, ищем заголовки таблицы');
//     }
    
//     if (tableRange) {
//       // Работаем с именованным диапазоном как с таблицей
//       const lastRow = tableRange.getLastRow();
//       const firstColumn = tableRange.getColumn();
//       const numColumns = tableRange.getNumColumns();
      
//       // Получаем заголовки для определения порядка столбцов
//       const headers = sheet.getRange(tableRange.getRow(), firstColumn, 1, numColumns).getValues()[0];
      
//       // Находим индексы нужных столбцов
//       const dateIndex = headers.indexOf('Дата');
//       const categoryIndex = headers.indexOf('Категория');
//       const amountIndex = headers.indexOf('Сумма');
//       const monthIndex = headers.indexOf('Месяц');
//       const commentIndex = headers.indexOf('Комментарий');
      
//       // Создаем массив для новой строки
//       const newRow = new Array(numColumns).fill('');
      
//       if (dateIndex !== -1) newRow[dateIndex] = currentDate;
//       if (categoryIndex !== -1) newRow[categoryIndex] = data.category;
//       if (amountIndex !== -1) newRow[amountIndex] = parseFloat(data.amount);
//       if (monthIndex !== -1) newRow[monthIndex] = monthDate;
//       if (commentIndex !== -1 && data.comment) newRow[commentIndex] = data.comment;
      
//       // Добавляем новую строку в таблицу
//       sheet.getRange(lastRow + 1, firstColumn, 1, numColumns).setValues([newRow]);
      
//     } else {
//       // Альтернативный метод: ищем заголовки таблицы на листе
//       const dataRange = sheet.getDataRange();
//       const values = dataRange.getValues();
      
//       // Ищем строку с заголовками
//       let headerRowIndex = -1;
//       let headers = [];
      
//       for (let i = 0; i < values.length; i++) {
//         if (values[i].includes('Дата') && values[i].includes('Категория') && values[i].includes('Сумма')) {
//           headerRowIndex = i;
//           headers = values[i];
//           break;
//         }
//       }
      
//       if (headerRowIndex === -1) {
//         throw new Error('Не найдена таблица с заголовками "Дата", "Категория", "Сумма"');
//       }
      
//       // Находим индексы нужных столбцов
//       const dateIndex = headers.indexOf('Дата');
//       const categoryIndex = headers.indexOf('Категория');
//       const amountIndex = headers.indexOf('Сумма');
//       const monthIndex = headers.indexOf('Месяц');
//       const commentIndex = headers.indexOf('Комментарий');
      
//       // Находим последнюю заполненную строку в таблице
//       let lastDataRow = headerRowIndex + 1;
//       for (let i = headerRowIndex + 1; i < values.length; i++) {
//         // Проверяем, есть ли данные хотя бы в одном из основных столбцов
//         if (values[i][dateIndex] || values[i][categoryIndex] || values[i][amountIndex]) {
//           lastDataRow = i + 1; // +1 так как индексы в getRange начинаются с 1
//         }
//       }
      
//       // Создаем массив для новой строки
//       const newRow = new Array(headers.length).fill('');
      
//       if (dateIndex !== -1) newRow[dateIndex] = currentDate;
//       if (categoryIndex !== -1) newRow[categoryIndex] = data.category;
//       if (amountIndex !== -1) newRow[amountIndex] = parseFloat(data.amount);
//       if (monthIndex !== -1) newRow[monthIndex] = monthDate;
//       if (commentIndex !== -1 && data.comment) newRow[commentIndex] = data.comment;
      
//       // Добавляем новую строку после последней заполненной
//       sheet.getRange(lastDataRow + 1, 1, 1, headers.length).setValues([newRow]);
//     }
    
//     // Возвращаем успешный результат
//     const monthYearDisplay = data.month + ' ' + data.year;
//     return {
//       success: true,
//       message: `✅ Запись успешно добавлена!\n📁 Категория: ${data.category}\n💰 Сумма: ${data.amount}₽\n📅 Период: ${monthYearDisplay}${data.comment ? '\n💬 Комментарий: ' + data.comment : ''}`
//     };
    
//   } catch (error) {
//     console.log('Ошибка при сохранении записи: ' + error.toString());
//     return {
//       success: false,
//       message: 'Ошибка при сохранении: ' + error.toString()
//     };
//   }
// }


/**
 * Альтернативная функция для вставки строки в таблицу
 * Использует метод insertRowAfter для добавления строки внутрь таблицы
 */
function saveRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Операции");
    
    if (!sheet) {
      throw new Error('Лист "Операции" не найден');
    }
    
    // Ищем таблицу по заголовкам
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let headerRow = -1;
    let lastDataRow = -1;
    
    // Находим строку с заголовками
    for (let i = 0; i < values.length; i++) {
      if (values[i].includes('Дата') && values[i].includes('Категория')) {
        headerRow = i + 1; // +1 для индекса строки в Sheet
        
        // Находим последнюю заполненную строку таблицы
        for (let j = i + 1; j < values.length; j++) {
          if (values[j].some(cell => cell !== '')) {
            lastDataRow = j + 1;
          } else {
            break; // Прекращаем, если встретили полностью пустую строку
          }
        }
        break;
      }
    }
    
    if (headerRow === -1) {
      throw new Error('Таблица не найдена');
    }
    
    // Если нет данных, добавляем после заголовка
    if (lastDataRow === -1) {
      lastDataRow = headerRow;
    }
    
    // Вставляем новую строку после последней строки с данными
    sheet.insertRowAfter(lastDataRow);
    
    // Подготавливаем данные
    const currentDate = new Date();
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    const monthNumber = monthNames.indexOf(data.month) + 1;
    const monthDate = new Date(parseInt(data.year), monthNumber - 1, 1);
    
    // Получаем заголовки
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Создаем массив данных
    const newRowData = [];
    for (let i = 0; i < headers.length; i++) {
      switch(headers[i]) {
        case 'Дата':
          newRowData.push(currentDate);
          break;
        case 'Категория':
          newRowData.push(data.category);
          break;
        case 'Сумма':
          newRowData.push(parseFloat(data.amount));
          break;
        case 'Месяц':
          newRowData.push(monthDate);
          break;
        case 'Комментарий':
          newRowData.push(data.comment || '');
          break;
        default:
          newRowData.push('');
      }
    }
    
    // Записываем данные в новую строку
    sheet.getRange(lastDataRow + 1, 1, 1, newRowData.length).setValues([newRowData]);
    
    return {
      success: true,
      message: `✅ Запись добавлена в таблицу!`
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Ошибка: ' + error.toString()
    };
  }
}

/**
 * Включить файл HTML в основной файл (для работы с CSS и JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
