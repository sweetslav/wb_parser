// ============================================================
// ТЕСТ: ДИАГНОСТИКА API КОНТЕНТА
// ============================================================

function testContentAPI() {
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
    muteHttpExceptions: true
  };
  
  // --- ТЕСТ 1: Запрос с limit=100 ---
  Logger.log('📌 ТЕСТ 1: Запрос с limit=100');
  let payload1 = {
    settings: {
      cursor: {
        limit: 100
      }
    }
  };
  
  try {
    const response1 = UrlFetchApp.fetch(
      'https://content-api.wildberries.ru/content/v2/get/cards/list',
      {
        ...options,
        payload: JSON.stringify(payload1)
      }
    );
    const data1 = JSON.parse(response1.getContentText());
    const cards1 = data1.cards || [];
    Logger.log('📊 Получено карточек: ' + cards1.length);
    Logger.log('🔍 cursor: ' + JSON.stringify(data1.cursor));
    Logger.log('📄 Первая карточка: ' + JSON.stringify(cards1[0] || {}));
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  Utilities.sleep(500);
  
  // --- ТЕСТ 2: Запрос с limit=1000 ---
  Logger.log('📌 ТЕСТ 2: Запрос с limit=1000');
  let payload2 = {
    settings: {
      cursor: {
        limit: 1000
      }
    }
  };
  
  try {
    const response2 = UrlFetchApp.fetch(
      'https://content-api.wildberries.ru/content/v2/get/cards/list',
      {
        ...options,
        payload: JSON.stringify(payload2)
      }
    );
    const data2 = JSON.parse(response2.getContentText());
    const cards2 = data2.cards || [];
    Logger.log('📊 Получено карточек: ' + cards2.length);
    Logger.log('🔍 cursor: ' + JSON.stringify(data2.cursor));
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  Utilities.sleep(500);
  
  // --- ТЕСТ 3: Запрос с сортировкой по убыванию (чтобы получить последние карточки) ---
  Logger.log('📌 ТЕСТ 3: Запрос с сортировкой по убыванию (desc)');
  let payload3 = {
    settings: {
      sort: {
        ascending: false
      },
      cursor: {
        limit: 100
      }
    }
  };
  
  try {
    const response3 = UrlFetchApp.fetch(
      'https://content-api.wildberries.ru/content/v2/get/cards/list',
      {
        ...options,
        payload: JSON.stringify(payload3)
      }
    );
    const data3 = JSON.parse(response3.getContentText());
    const cards3 = data3.cards || [];
    Logger.log('📊 Получено карточек: ' + cards3.length);
    Logger.log('🔍 cursor: ' + JSON.stringify(data3.cursor));
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
  
  // --- ТЕСТ 4: Запрос с фильтром по статусу (только активные) ---
  Logger.log('📌 ТЕСТ 4: Запрос с фильтром (только активные)');
  let payload4 = {
    settings: {
      filter: {
        withPhoto: -1,
        status: 1  // 1 - активные карточки
      },
      cursor: {
        limit: 100
      }
    }
  };
  
  try {
    const response4 = UrlFetchApp.fetch(
      'https://content-api.wildberries.ru/content/v2/get/cards/list',
      {
        ...options,
        payload: JSON.stringify(payload4)
      }
    );
    const data4 = JSON.parse(response4.getContentText());
    const cards4 = data4.cards || [];
    Logger.log('📊 Получено карточек: ' + cards4.length);
    Logger.log('🔍 cursor: ' + JSON.stringify(data4.cursor));
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
