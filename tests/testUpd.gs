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
