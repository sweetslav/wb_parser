// ============================================================
// ТЕСТ: ПРОВЕРКА ЗАГРУЗКИ ОСТАТКОВ (ВЕРСИЯ ДЛЯ ОТЛАДКИ)
// ============================================================
function testStockDataDebug() {
  Logger.log('🔍 ===== ТЕСТ ЗАГРУЗКИ ОСТАТКОВ (ОТЛАДОЧНАЯ) =====');
  
  try {
    const SPREADSHEET_ID = '1pP1RlNjgfxcDNw9Icwep0Pl3PyJikNIeQE6bMdCDD70';
    const SHEET_NAME = 'unit расчет';
    
    Logger.log('📤 Открываем справочник...');
    const stockBook = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = stockBook.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('❌ Лист не найден');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`📊 Всего строк: ${lastRow}`);
    
    // Проверяем заголовки
    const headers = sheet.getRange(1, 1, 1, 28).getValues()[0];
    Logger.log('📋 Заголовки:');
    Logger.log(`   M (индекс 12): "${headers[12]}"`);
    Logger.log(`   Z (индекс 25): "${headers[25]}"`);
    Logger.log(`   AA (индекс 26): "${headers[26]}"`);
    Logger.log(`   AB (индекс 27): "${headers[27]}"`);
    
    // ЧИТАЕМ ДАННЫЕ ПРАВИЛЬНО - столбец M (индекс 12) для nmId
    Logger.log('📊 Читаем данные (столбцы A-AB)...');
    const range = sheet.getRange(2, 1, Math.min(20, lastRow - 1), 28);
    const values = range.getValues();
    
    Logger.log('📊 Первые 10 строк (показываем nmId из столбца M):');
    let foundCount = 0;
    
    values.forEach((row, index) => {
      const rowNum = index + 2;
      
      // ✅ ПРАВИЛЬНО: nmId в столбце M (индекс 12)
      const nmIdRaw = row[12];
      const vendorCode = row[1]; // столбец B
      const fbw = row[25] || 0;
      const fbo = row[26] || 0;
      const ourStock = row[27] || 0;
      
      // Преобразуем nmId в число
      let nmIdNum = null;
      const type = typeof nmIdRaw;
      
      if (type === 'number') {
        nmIdNum = nmIdRaw;
      } else if (type === 'string') {
        const cleaned = nmIdRaw.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      } else if (nmIdRaw && type === 'object') {
        const strValue = String(nmIdRaw);
        const cleaned = strValue.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      }
      
      if (nmIdNum && !isNaN(nmIdNum) && nmIdNum > 0) {
        foundCount++;
        Logger.log(`   ${foundCount}. Строка ${rowNum}: nmId=${nmIdNum}, vendorCode=${vendorCode}, FBW=${fbw}, FBO=${fbo}, 1С=${ourStock}`);
      } else {
        if (index < 5) {
          Logger.log(`   ⏭️ Строка ${rowNum}: nmIdRaw=${nmIdRaw} (тип: ${type}), пропущено`);
        }
      }
    });
    
    Logger.log(`✅ Найдено ${foundCount} валидных nmId в первых 20 строках`);
    
    // Теперь загружаем ВСЕ данные правильно
    Logger.log('📊 Загружаем ВСЕ остатки (правильный алгоритм)...');
    
    const allRange = sheet.getRange(2, 1, lastRow - 1, 28);
    const allValues = allRange.getValues();
    
    const stockData = {};
    let totalFound = 0;
    let totalSkipped = 0;
    
    allValues.forEach((row, index) => {
      const nmIdRaw = row[12]; // столбец M
      
      // Преобразуем nmId
      let nmIdNum = null;
      const type = typeof nmIdRaw;
      
      if (type === 'number') {
        nmIdNum = nmIdRaw;
      } else if (type === 'string') {
        const cleaned = nmIdRaw.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      } else if (nmIdRaw && type === 'object') {
        const strValue = String(nmIdRaw);
        const cleaned = strValue.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      }
      
      if (nmIdNum && !isNaN(nmIdNum) && nmIdNum > 0) {
        const fbw = Number(row[25]) || 0;
        const fbo = Number(row[26]) || 0;
        const ourStock = Number(row[27]) || 0;
        
        stockData[nmIdNum] = {
          wbStock: fbw + fbo,
          ourStock: ourStock,
          fbw: fbw,
          fbo: fbo
        };
        totalFound++;
      } else {
        totalSkipped++;
      }
    });
    
    Logger.log(`✅ Загружено: ${totalFound} товаров`);
    Logger.log(`⏭️ Пропущено: ${totalSkipped} строк`);
    
    // Проверяем конкретные nmId
    const testIds = [468667061, 366120728, 366120896, 404235933, 33705652];
    Logger.log('🔍 Проверка конкретных nmId в загруженных данных:');
    testIds.forEach(id => {
      if (stockData[id]) {
        const data = stockData[id];
        Logger.log(`   ✅ nmId=${id}: ВБ=${data.wbStock} (FBW=${data.fbw}+FBO=${data.fbo}), Наши=${data.ourStock}`);
      } else {
        Logger.log(`   ❌ nmId=${id} не найден`);
      }
    });
    
    // Товары с остатками
    const withStock = Object.keys(stockData).filter(key => 
      stockData[key].wbStock > 0 || stockData[key].ourStock > 0
    );
    Logger.log(`📦 Товаров с остатками: ${withStock.length}`);
    
    if (withStock.length > 0) {
      Logger.log('📊 Примеры товаров с остатками:');
      const sample = withStock.slice(0, 10);
      sample.forEach(key => {
        const data = stockData[key];
        Logger.log(`   nmId=${key}: ВБ=${data.wbStock} (FBW=${data.fbw}+FBO=${data.fbo}), Наши=${data.ourStock}`);
      });
    }
    
    // Общая статистика
    const totalWb = Object.values(stockData).reduce((sum, d) => sum + d.wbStock, 0);
    const totalOur = Object.values(stockData).reduce((sum, d) => sum + d.ourStock, 0);
    Logger.log(`📊 ИТОГО остатков: ВБ=${totalWb} шт., Наши=${totalOur} шт.`);
    
    Logger.log('✅ ===== ТЕСТ ЗАВЕРШЕН =====');
    
  } catch (e) {
    Logger.log('❌ Ошибка: ' + e.message);
    Logger.log(e.stack);
  }
}
