function compareAdvMethods() {
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
  
  Logger.log('📅 Период: ' + from + ' - ' + to);
  
  // 1. Метод /upd
  try {
    const updUrl = `https://advert-api.wildberries.ru/adv/v1/upd?from=${from}&to=${to}`;
    const updResponse = UrlFetchApp.fetch(updUrl, options);
    const updData = JSON.parse(updResponse.getContentText());
    
    let updTotal = 0;
    if (Array.isArray(updData)) {
      updData.forEach(item => {
        updTotal += item.updSum || 0;
      });
    }
    Logger.log('💰 /upd: ' + updTotal + ' ₽');
  } catch (e) {
    Logger.log('❌ /upd ошибка: ' + e.message);
  }
  
  // 2. Метод /fullstats
  try {
    // Получаем кампании
    const campaignsUrl = 'https://advert-api.wildberries.ru/api/advert/v2/adverts?statuses=9,11';
    const campaignsResponse = UrlFetchApp.fetch(campaignsUrl, options);
    const campaignsData = JSON.parse(campaignsResponse.getContentText());
    const campaigns = campaignsData.adverts || [];
    
    const limited = campaigns.slice(0, 50);
    const ids = limited.map(c => c.id).join(',');
    
    const statsUrl = `https://advert-api.wildberries.ru/adv/v3/fullstats?ids=${ids}&beginDate=${from}&endDate=${to}`;
    const statsResponse = UrlFetchApp.fetch(statsUrl, options);
    const statsData = JSON.parse(statsResponse.getContentText());
    
    let statsTotal = 0;
    if (Array.isArray(statsData)) {
      statsData.forEach(item => {
        statsTotal += item.sum || 0;
      });
    }
    Logger.log('💰 /fullstats (первые 50 кампаний): ' + statsTotal + ' ₽');
    
  } catch (e) {
    Logger.log('❌ /fullstats ошибка: ' + e.message);
  }
}
