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

/**
 * Получить список оплаченных месяцев для категории
 */
function getPaidMonths(category) {
  try {
    if (!category) {
      return [];
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Операции");
    
    if (!sheet) {
      console.log('Лист "Операции" не найден');
      return [];
    }
    
    // Получаем все данные с листа
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Находим строку с заголовками
    let headerRow = -1;
    let categoryCol = -1;
    let monthCol = -1;
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const catIndex = row.indexOf('Категория');
      const monthIndex = row.indexOf('Месяц');
      
      if (catIndex !== -1 && monthIndex !== -1) {
        headerRow = i;
        categoryCol = catIndex;
        monthCol = monthIndex;
        break;
      }
    }
    
    if (headerRow === -1) {
      console.log('Заголовки таблицы не найдены');
      return [];
    }
    
    // Собираем оплаченные месяцы для выбранной категории
    const paidMonths = [];
    const monthNames = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    
    for (let i = headerRow + 1; i < values.length; i++) {
      const row = values[i];
      
      // Проверяем, совпадает ли категория
      if (row[categoryCol] === category && row[monthCol]) {
        // Преобразуем дату в строку месяц-год
        const monthDate = new Date(row[monthCol]);
        if (!isNaN(monthDate.getTime())) {
          const month = monthNames[monthDate.getMonth()];
          const year = monthDate.getFullYear();
          const monthYearStr = `${month} ${year}`;
          
          // Добавляем, если еще нет в списке
          if (!paidMonths.includes(monthYearStr)) {
            paidMonths.push(monthYearStr);
          }
        }
      }
    }
    
    // Сортируем по дате (сначала преобразуем обратно в даты для правильной сортировки)
    paidMonths.sort((a, b) => {
      const [monthA, yearA] = a.split(' ');
      const [monthB, yearB] = b.split(' ');
      const dateA = new Date(parseInt(yearA), monthNames.indexOf(monthA));
      const dateB = new Date(parseInt(yearB), monthNames.indexOf(monthB));
      return dateB - dateA; // Сортировка от новых к старым
    });
    
    return paidMonths;
    
  } catch (error) {
    console.log('Ошибка при получении оплаченных месяцев: ' + error.toString());
    return [];
  }
}