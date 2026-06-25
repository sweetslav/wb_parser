// ============================================================
// ТЕСТ ТОКЕНА WB
// ============================================================
function testToken() {
  // Берем токен из хранилища
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  
  if (!token) {
    Logger.log('❌ Токен не найден в хранилище! Запусти setWBToken()');
    return;
  }
  
  Logger.log('🔑 Токен найден. Начинаем проверку...');
  
  // Тестовый запрос к API заказов (за сегодня)
  const today = new Date();
  const dateStr = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const url = `https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=${dateStr}&limit=1`;
  
  // Пробуем разные варианты заголовков
  const options1 = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  try {
    Logger.log('📤 Пробуем вариант 1: Authorization: token');
    const response1 = UrlFetchApp.fetch(url, options1);
    const code1 = response1.getResponseCode();
    Logger.log('📥 Код ответа (вар.1): ' + code1);
    
    if (code1 === 200) {
      Logger.log('✅ УРА! Вариант 1 работает! Токен живой.');
      return;
    }
    
    // Пробуем второй вариант - с префиксом 'Bearer'
    const options2 = {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    };
    
    Logger.log('📤 Пробуем вариант 2: Authorization: Bearer token');
    const response2 = UrlFetchApp.fetch(url, options2);
    const code2 = response2.getResponseCode();
    Logger.log('📥 Код ответа (вар.2): ' + code2);
    
    if (code2 === 200) {
      Logger.log('✅ УРА! Вариант 2 работает! Токен живой.');
      // Если работает с Bearer, обновим токен в хранилище, чтобы основной код работал
      PropertiesService.getScriptProperties().setProperty('WB_TOKEN', 'Bearer ' + token);
      Logger.log('💡 Токен в хранилище обновлен с префиксом Bearer.');
      return;
    }
    
    Logger.log('❌ Оба варианта не сработали. Код 200 не получен.');
    
    // Если код 401, пробуем прочитать тело ошибки
    if (code1 === 401) {
      Logger.log('🔍 Тело ошибки (вар.1): ' + response1.getContentText());
    }
    if (code2 === 401) {
      Logger.log('🔍 Тело ошибки (вар.2): ' + response2.getContentText());
    }
    
  } catch (e) {
    Logger.log('❌ Исключение при запросе: ' + e.message);
  }
}
