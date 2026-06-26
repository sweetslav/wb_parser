function testFullstatsCoverage() {
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
    // Получаем список активных кампаний
    const campaignsUrl = 'https://advert-api.wildberries.ru/api/advert/v2/adverts?statuses=9,11';
    const response = UrlFetchApp.fetch(campaignsUrl, options);
    const data = JSON.parse(response.getContentText());
    const campaigns = data.adverts || [];
    
    Logger.log('📢 Всего активных кампаний (статус 9 или 11): ' + campaigns.length);
    
    // Считаем, сколько из них имеют затраты (через nm_settings)
    let withNm = 0;
    campaigns.forEach(c => {
      if (c.nm_settings && c.nm_settings.length > 0) {
        withNm++;
      }
    });
    
    Logger.log('📊 Из них с артикулами (nm_settings): ' + withNm);
    Logger.log('⚠️ В fullstats мы передаем только первые 50 кампаний!');
    Logger.log('   (ограничение API — максимум 50 ID за запрос)');
    
    // Проверяем, сколько кампаний с затратами попадут в выборку
    const limited = campaigns.slice(0, 50);
    Logger.log('📊 В запрос попадает ' + limited.length + ' кампаний');
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
