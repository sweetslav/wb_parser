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
