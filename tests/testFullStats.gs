function testFullStats() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  // Берем даты за последние 7 дней
  const today = new Date();
  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(today.getDate() - 7);
  
  const from = Utilities.formatDate(sevenDaysAgo, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const to = Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  
  // Получаем ID нескольких кампаний для теста
  const options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  try {
    // Сначала получаем список кампаний
    const advertsResponse = UrlFetchApp.fetch(WB_CAMPAIGNS_INFO_API + '?statuses=9&limit=5', options);
    const advertsData = JSON.parse(advertsResponse.getContentText());
    const campaigns = advertsData.adverts || [];
    
    if (campaigns.length === 0) {
      Logger.log('⚠️ Нет активных кампаний');
      return;
    }
    
    const ids = campaigns.map(c => c.advertId).join(',');
    Logger.log('📢 Тестовые ID кампаний: ' + ids);
    
    // Запрашиваем статистику
    const statsUrl = `https://advert-api.wildberries.ru/adv/v3/fullstats?ids=${ids}&beginDate=${from}&endDate=${to}`;
    Logger.log('📤 Запрос: ' + statsUrl);
    
    const statsResponse = UrlFetchApp.fetch(statsUrl, options);
    Logger.log('📥 Код ответа: ' + statsResponse.getResponseCode());
    
    const content = statsResponse.getContentText();
    Logger.log('📄 Тип ответа: ' + typeof content);
    Logger.log('📄 Первые 500 символов: ' + content.substring(0, 500));
    
    // Пробуем распарсить
    try {
      const parsed = JSON.parse(content);
      Logger.log('🔍 Тип после парсинга: ' + typeof parsed);
      Logger.log('🔍 Это массив? ' + Array.isArray(parsed));
      Logger.log('🔍 Структура: ' + JSON.stringify(parsed).substring(0, 300));
      
      if (Array.isArray(parsed)) {
        Logger.log('✅ Количество записей: ' + parsed.length);
      } else if (parsed.error || parsed.title) {
        Logger.log('❌ Ошибка от API: ' + JSON.stringify(parsed));
      } else {
        Logger.log('⚠️ Неизвестная структура: ' + JSON.stringify(parsed).substring(0, 300));
      }
    } catch (e) {
      Logger.log('❌ Ошибка парсинга JSON: ' + e.message);
      Logger.log('📄 Сырой ответ: ' + content);
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
