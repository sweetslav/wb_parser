function testUpdWithShift() {
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) {
    Logger.log('❌ Токен не найден');
    return;
  }
  
  // Даты из таблицы (предположим, это май 2026)
  const tableDateFrom = new Date(2026, 4, 1);  // 1 мая 2026
  const tableDateTo = new Date(2026, 4, 31);   // 31 мая 2026
  
  // Сдвигаем на неделю назад для рекламы
  const advDateFrom = new Date(tableDateFrom);
  advDateFrom.setDate(advDateFrom.getDate() - 7);
  const advDateTo = new Date(tableDateTo);
  advDateTo.setDate(advDateTo.getDate() - 7);
  
  const from = Utilities.formatDate(advDateFrom, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const to = Utilities.formatDate(advDateTo, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  
  Logger.log('📅 Период в таблице: 01.05.2026 - 31.05.2026');
  Logger.log('📅 Период для рекламы (со сдвигом -7 дней): ' + from + ' - ' + to);
  
  const options = {
    method: 'GET',
    headers: { 'Authorization': token },
    muteHttpExceptions: true
  };
  
  try {
    const url = `https://advert-api.wildberries.ru/adv/v1/upd?from=${from}&to=${to}`;
    Logger.log('📤 Запрос к /upd...');
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('📥 Код: ' + response.getResponseCode());
    
    const data = JSON.parse(response.getContentText());
    if (Array.isArray(data)) {
      let totalSum = 0;
      data.forEach(item => {
        totalSum += item.updSum || 0;
      });
      Logger.log('💰 Общая сумма затрат за период: ' + totalSum + ' ₽');
      Logger.log('📊 Количество записей: ' + data.length);
      
      // Показываем первые 5 записей для примера
      if (data.length > 0) {
        Logger.log('🔍 Примеры записей:');
        data.slice(0, 5).forEach(item => {
          Logger.log('   ' + item.campName + ': ' + item.updSum + ' ₽ (' + item.updTime + ')');
        });
      }
    }
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
  }
}
