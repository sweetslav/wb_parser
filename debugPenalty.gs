function debugPenalty() {
  Logger.log('🔍 ===== ОТЛАДКА ШТРАФОВ =====');
  
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) { Logger.log('❌ Токен не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adv_effectiveness');
  const dateFrom = new Date(sheet.getRange('B2').getValue());
  const dateTo = new Date(sheet.getRange('C2').getValue());
  
  const url = 'https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed';
  const payload = {
    dateFrom: formatDate(dateFrom),
    dateTo: formatDate(dateTo),
    limit: 100000,
    rrdId: 0,
    period: 'weekly'
  };
  
  const options = {
    method: 'POST',
    headers: { 'Authorization': token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  // Ищем все строки с penalty > 0
  const penalties = data.filter(row => Number(row.penalty) > 0);
  
  Logger.log('📊 Найдено штрафов: ' + penalties.length);
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  penalties.forEach((row, i) => {
    Logger.log('Штраф ' + (i+1) + ':');
    Logger.log('   nmId: ' + row.nmId);
    Logger.log('   penalty: ' + row.penalty);
    Logger.log('   docTypeName: ' + row.docTypeName);
    Logger.log('   sellerOperName: ' + row.sellerOperName);
    Logger.log('   bonusTypeName: ' + row.bonusTypeName);
    Logger.log('   ---');
  });
  
  Logger.log('✅ ===== ОТЛАДКА ЗАВЕРШЕНА =====');
}
