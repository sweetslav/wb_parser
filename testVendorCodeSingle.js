// ============================================================
// ТЕСТ 1: ПОЛУЧЕНИЕ VENDORCODE ПО ОДНОМУ АРТИКУЛУ
// ============================================================
function testVendorCodeSingle() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  // Возьми ОДИН реальный nmId из твоей таблицы (например, из столбца A)
  const testNmId = 158868824; // ← ЗАМЕНИ НА РЕАЛЬНЫЙ nmId ИЗ ТВОЕЙ ТАБЛИЦЫ
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      settings: {
        cursor: {
          limit: 10
        },
        filter: {
          nmIDs: [testNmId]
        }
      }
    }),
    muteHttpExceptions: true
  };
  
  try {
    const url = 'https://content-api.wildberries.ru/content/v2/get/cards/list';
    Logger.log('📤 Запрос к API Контента для nmId: ' + testNmId);
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code !== 200) {
      Logger.log('⚠️ Ошибка: ' + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    Logger.log('📄 Полный ответ: ' + JSON.stringify(data, null, 2));
    
    const cards = data.cards || [];
    Logger.log('📊 Найдено карточек: ' + cards.length);
    
    if (cards.length > 0) {
      const card = cards[0];
      Logger.log('🔍 Карточка товара:');
      Logger.log('   nmID: ' + card.nmID);
      Logger.log('   vendorCode: ' + card.vendorCode);
      Logger.log('   supplierArticle: ' + card.supplierArticle);
      Logger.log('   article: ' + card.article);
      Logger.log('   title: ' + card.title);
      
      // Показываем все поля, чтобы понять, где лежит vendorCode
      Logger.log('📋 Все поля: ' + Object.keys(card).join(', '));
    } else {
      Logger.log('⚠️ Карточка не найдена');
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
