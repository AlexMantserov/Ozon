function getOzonProducts() {
  // Настройки
  const OZON_CLIENT_ID = "177297";
  const OZON_API_KEY = "142e8aef-9b5e-4af0-bbb7-ba08a0f6069c";
  const SPREADSHEET_ID = "1IRXrnXFyBBNt_0fBYjja-gMqOZ8frBUSBbP7TOwlg4c";
  
  try {
    console.log("Получаем товары с Ozon...");
    
    // 1. Получаем список товаров
    const productsUrl = "https://api-seller.ozon.ru/v2/product/list";
    const productsPayload = {
      "filter": {"visibility": "ALL"},
      "last_id": "",
      "limit": 1000
    };
    
    const productsOptions = {
      method: "POST",
      headers: {
        "Client-Id": OZON_CLIENT_ID,
        "Api-Key": OZON_API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(productsPayload),
      muteHttpExceptions: true
    };
    
    const productsResponse = UrlFetchApp.fetch(productsUrl, productsOptions);
    
    if (productsResponse.getResponseCode() !== 200) {
      throw new Error("Ошибка при получении списка товаров");
    }
    
    const productsData = JSON.parse(productsResponse.getContentText());
    const items = productsData.result.items;
    
    console.log(`Найдено товаров: ${items.length}`);
    
    // 2. Получаем детальную информацию
    const detailsUrl = "https://api-seller.ozon.ru/v2/product/info/list";
    const offerIds = items.map(item => item.offer_id);
    
    const detailsPayload = {
      "offer_id": offerIds
    };
    
    const detailsOptions = {
      method: "POST",
      headers: {
        "Client-Id": OZON_CLIENT_ID,
        "Api-Key": OZON_API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(detailsPayload),
      muteHttpExceptions: true
    };
    
    const detailsResponse = UrlFetchApp.fetch(detailsUrl, detailsOptions);
    
    if (detailsResponse.getResponseCode() !== 200) {
      throw new Error("Ошибка при получении детальной информации");
    }
    
    const detailsData = JSON.parse(detailsResponse.getContentText());
    
    // 3. Подготавливаем данные для таблицы
    const tableData = detailsData.result.items.map(item => {
      return [
        item.offer_id || "Нет артикула",
        item.name || item.title || "Без названия",
        item.price ? (item.price / 100).toFixed(2) : "0.00",
        item.old_price ? (item.old_price / 100).toFixed(2) : "0.00",
        new Date().toLocaleString("ru-RU")
      ];
    });
    
    // 4. Записываем в Google Таблицу
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName("Ozon товары");
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet("Ozon товары");
    }
    
    // Очищаем и заполняем таблицу
    sheet.clear();
    
    // Заголовки
    const headers = [["Артикул", "Название товара", "Текущая цена", "Цена до скидки", "Дата обновления"]];
    sheet.getRange(1, 1, 1, 5).setValues(headers);
    
    // Данные
    if (tableData.length > 0) {
      sheet.getRange(2, 1, tableData.length, 5).setValues(tableData);
    }
    
    // Форматирование
    sheet.getRange(1, 1, 1, 5)
      .setBackground("#4a86e8")
      .setFontColor("white")
      .setFontWeight("bold");
    
    sheet.autoResizeColumns(1, 5);
    sheet.setFrozenRows(1);
    
    // Формат цен
    if (tableData.length > 0) {
      sheet.getRange(2, 3, tableData.length, 2).setNumberFormat("#,##0.00 ₽");
    }
    
    console.log("Готово! Записано " + tableData.length + " товаров");
    
  } catch (error) {
    console.error("Ошибка:", error);
    SpreadsheetApp.getActiveSpreadsheet().toast("Ошибка: " + error.message, "Ошибка", 10);
  }
}

// Добавляем меню в таблицу
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ozon')
    .addItem('Получить товары', 'getOzonProducts')
    .addToUi();
}
