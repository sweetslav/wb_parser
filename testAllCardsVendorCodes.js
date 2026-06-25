// ============================================================
// ТЕСТ: ПОЛУЧИТЬ ВСЕ КАРТОЧКИ И НАЙТИ VENDORCODE
// ============================================================
function testAllCardsVendorCodes() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      settings: {
        cursor: {
          limit: 1000  // Максимум 1000 карточек за запрос
        }
      }
    }),
    muteHttpExceptions: true
  };
  
  try {
    const url = 'https://content-api.wildberries.ru/content/v2/get/cards/list';
    Logger.log('📤 Запрашиваем все карточки...');
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code !== 200) {
      Logger.log('⚠️ Ошибка: ' + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    const cards = data.cards || [];
    Logger.log('✅ Получено карточек: ' + cards.length);
    
    // Создаем словарь nmID → vendorCode
    const vendorMap = {};
    cards.forEach(card => {
      const nmId = card.nmID;
      if (nmId && card.vendorCode) {
        vendorMap[nmId] = card.vendorCode;
      }
    });
    
    Logger.log('📊 Найдено vendorCode для ' + Object.keys(vendorMap).length + ' артикулов');
    
    // Показываем первые 5 для примера
    const sample = Object.keys(vendorMap).slice(0, 5);
    sample.forEach(nmId => {
      Logger.log('   nmId: ' + nmId + ' → vendorCode: ' + vendorMap[nmId]);
    });
    
    // Проверяем конкретный артикул, который был в тесте
    const testNmId = 158868824;
    if (vendorMap[testNmId]) {
      Logger.log('✅ Найден vendorCode для ' + testNmId + ': ' + vendorMap[testNmId]);
    } else {
      Logger.log('⚠️ vendorCode для ' + testNmId + ' не найден');
    }
    
    return vendorMap;
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
    return {};
  }
}
