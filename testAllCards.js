// ============================================================
// ТЕСТ 2: ПОЛУЧЕНИЕ ВСЕХ КАРТОЧЕК (БЕЗ ФИЛЬТРА)
// ============================================================
function testAllCards() {
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
          limit: 5  // Берем только 5 карточек для теста
        }
      }
    }),
    muteHttpExceptions: true
  };
  
  try {
    const url = 'https://content-api.wildberries.ru/content/v2/get/cards/list';
    Logger.log('📤 Запрос к API Контента (без фильтра)...');
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    Logger.log('📥 Код ответа: ' + code);
    
    if (code !== 200) {
      Logger.log('⚠️ Ошибка: ' + response.getContentText());
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    const cards = data.cards || [];
    Logger.log('📊 Найдено карточек: ' + cards.length);
    
    if (cards.length > 0) {
      cards.forEach((card, index) => {
        Logger.log('🔍 Карточка ' + (index + 1) + ':');
        Logger.log('   nmID: ' + card.nmID);
        Logger.log('   vendorCode: ' + card.vendorCode);
        Logger.log('   supplierArticle: ' + card.supplierArticle);
        Logger.log('   article: ' + card.article);
        Logger.log('   title: ' + card.title);
        Logger.log('   ---');
      });
      
      // Показываем структуру первой карточки
      Logger.log('📋 Поля первой карточки: ' + Object.keys(cards[0]).join(', '));
    } else {
      Logger.log('⚠️ Карточки не найдены');
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
