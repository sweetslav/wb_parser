// ============================================================
// ТЕСТЫ ДЛЯ РЕКЛАМНОГО API
// ============================================================

function testAdvBalance() {
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
  
  try {
    const url = 'https://advert-api.wildberries.ru/adv/v1/balance';
    Logger.log('📤 Проверяем баланс рекламного кабинета...');
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('📥 Код: ' + response.getResponseCode());
    Logger.log('📄 Ответ: ' + response.getContentText());
    
    const countUrl = 'https://advert-api.wildberries.ru/adv/v1/promotion/count';
    Logger.log('📤 Проверяем список кампаний...');
    const countResponse = UrlFetchApp.fetch(countUrl, options);
    Logger.log('📥 Код count: ' + countResponse.getResponseCode());
    Logger.log('📄 Ответ count: ' + countResponse.getContentText());
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}

function testUpd() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  const today = new Date();
  const monthAgo = new Date(today);
  monthAgo.setDate(today.getDate() - 30);
  
  const from = Utilities.formatDate(monthAgo, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const to = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  
  const options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  try {
    const url = `https://advert-api.wildberries.ru/adv/v1/upd?from=${from}&to=${to}`;
    Logger.log('📤 Запрос к /upd: ' + url);
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('📥 Код: ' + response.getResponseCode());
    Logger.log('📄 Первые 500 символов: ' + response.getContentText().substring(0, 500));
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}

function testCampaigns() {
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
  
  try {
    const url = 'https://advert-api.wildberries.ru/api/advert/v2/adverts?statuses=9,11';
    Logger.log('📤 Запрос к /adverts...');
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('📥 Код: ' + response.getResponseCode());
    
    const data = JSON.parse(response.getContentText());
    const campaigns = data.adverts || [];
    Logger.log('📢 Найдено кампаний: ' + campaigns.length);
    
    if (campaigns.length > 0) {
      Logger.log('🔍 Пример кампании: ' + JSON.stringify(campaigns[0]).substring(0, 300));
      Logger.log('📋 Поля кампании: ' + Object.keys(campaigns[0]).join(', '));
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
