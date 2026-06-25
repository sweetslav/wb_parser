// ============================================================
// ТЕСТ: API ЦЕН И СКИДОК (ПОЛУЧЕНИЕ VENDORCODE)
// ============================================================

function testPricesAPI() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  const options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  // --- ТЕСТ 1: Получить первые 10 товаров ---
  Logger.log('📌 ТЕСТ 1: Получаем первые 10 товаров');
  try {
    const url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=10&offset=0';
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const goods = data.data?.listGoods || [];
      Logger.log('📊 Получено товаров: ' + goods.length);
      
      if (goods.length > 0) {
        // Показываем первый товар для примера
        const first = goods[0];
        Logger.log('🔍 Пример товара:');
        Logger.log('   nmId: ' + first.nmId);
        Logger.log('   vendorCode: ' + first.vendorCode);
        Logger.log('   price: ' + first.price);
        Logger.log('   discount: ' + first.discount);
        Logger.log('📋 Все поля: ' + Object.keys(first).join(', '));
      }
    } else {
      Logger.log('⚠️ Ошибка: ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  Utilities.sleep(500);
  
  // --- ТЕСТ 2: Получить 100 товаров (проверка пагинации) ---
  Logger.log('📌 ТЕСТ 2: Получаем 100 товаров');
  try {
    const url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=100&offset=0';
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const goods = data.data?.listGoods || [];
      Logger.log('📊 Получено товаров: ' + goods.length);
      
      // Считаем, сколько из них имеют vendorCode
      const withVendorCode = goods.filter(item => item.vendorCode).length;
      Logger.log('🏷️ С vendorCode: ' + withVendorCode);
    }
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  Utilities.sleep(500);
  
  // --- ТЕСТ 3: Проверить конкретный nmId ---
  Logger.log('📌 ТЕСТ 3: Проверяем конкретный nmId (158868824)');
  try {
    const testNmId = 158868824;
    const url = `https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=1&filterNmID=${testNmId}`;
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const goods = data.data?.listGoods || [];
      Logger.log('📊 Найдено товаров: ' + goods.length);
      
      if (goods.length > 0) {
        const item = goods[0];
        Logger.log('🔍 Результат:');
        Logger.log('   nmId: ' + item.nmId);
        Logger.log('   vendorCode: ' + item.vendorCode);
        Logger.log('   price: ' + item.price);
      } else {
        Logger.log('⚠️ Товар не найден');
      }
    }
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  Utilities.sleep(500);
  
  // --- ТЕСТ 4: Проверить, сколько всего товаров ---
  Logger.log('📌 ТЕСТ 4: Получаем первую страницу с limit=1000');
  try {
    const url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=1000&offset=0';
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const goods = data.data?.listGoods || [];
      Logger.log('📊 Получено товаров: ' + goods.length);
      
      const withVendorCode = goods.filter(item => item.vendorCode).length;
      Logger.log('🏷️ С vendorCode: ' + withVendorCode);
      
      // Проверяем структуру ответа
      Logger.log('📋 Структура data: ' + Object.keys(data.data || {}).join(', '));
    }
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
