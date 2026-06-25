function testCampaignsCount() {
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
    // Проверяем количество кампаний через /promotion/count
    const countUrl = 'https://advert-api.wildberries.ru/adv/v1/promotion/count';
    const countResponse = UrlFetchApp.fetch(countUrl, options);
    const countData = JSON.parse(countResponse.getContentText());
    
    Logger.log('📊 Общее количество кампаний: ' + countData.all);
    countData.adverts.forEach(group => {
      Logger.log('   Тип ' + group.type + ', статус ' + group.status + ': ' + group.count + ' шт.');
    });
    
    // Проверяем активные кампании (статус 9)
    const active = countData.adverts.filter(g => g.status === 9);
    if (active.length > 0) {
      Logger.log('✅ Активных кампаний (статус 9): ' + active.reduce((sum, g) => sum + g.count, 0));
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
