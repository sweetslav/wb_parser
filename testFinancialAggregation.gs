function testFinancialAggregation() {
  Logger.log('🔍 ===== ТЕСТ АГРЕГАЦИИ ФИНАНСОВ (v2) =====');
  
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) { Logger.log('❌ Токен не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adv_effectiveness');
  const dateFrom = new Date(sheet.getRange('B2').getValue());
  const dateTo = new Date(sheet.getRange('C2').getValue());
  
  Logger.log('📅 Период: ' + formatDate(dateFrom) + ' - ' + formatDate(dateTo));
  
  // Получаем детализацию
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
  
  Logger.log('📊 Всего строк: ' + data.length);
  
  // Агрегируем по nmId с правильными полями
  const result = {};
  let totalRetailPriceSales = 0;
  let totalRetailAmountSales = 0;
  let totalRetailPriceReturns = 0;
  let totalRetailAmountReturns = 0;
  let totalCommission = 0;
  let totalLogistics = 0;
  let totalStorage = 0;
  let totalPenalty = 0;
  let totalDeduction = 0;
  let totalAcceptance = 0;
  let totalAdditionalPayment = 0;
  let totalForPay = 0;
  let totalCashbackDiscount = 0;
  let totalCashbackAmount = 0;
  let totalCashbackCommission = 0;
  
  // Счетчик для отладки операций без nmId
  let noNmIdCount = 0;
  let storageRows = 0;
  let deductionRows = 0;
  
  data.forEach(row => {
    const nmId = row.nmId;
    
    // --- ХРАНЕНИЕ И УДЕРЖАНИЯ МОГУТ БЫТЬ БЕЗ nmId ---
    // Если нет nmId, пропускаем (это общие расходы, не привязанные к товару)
    if (!nmId) {
      noNmIdCount++;
      // Но все равно считаем общие суммы
      const storage = Number(row.paidStorage) || 0;
      const deduction = Number(row.deduction) || 0;
      const penalty = Number(row.penalty) || 0;
      const delivery = Number(row.deliveryService) || 0;
      const acceptance = Number(row.paidAcceptance) || 0;
      const additionalPayment = Number(row.additionalPayment) || 0;
      
      totalStorage += storage;
      totalDeduction += deduction;
      totalPenalty += penalty;
      totalLogistics += delivery;
      totalAcceptance += acceptance;
      totalAdditionalPayment += additionalPayment;
      
      if (storage > 0) storageRows++;
      if (deduction > 0) deductionRows++;
      
      // Для хранения и удержаний без nmId не создаем запись
      return;
    }
    
    const docType = row.docTypeName || '';
    const sellerOperName = row.sellerOperName || '';
    const quantity = Number(row.quantity) || 0;
    const retailPrice = Number(row.retailPrice) || 0;        // Розничная цена (до СПП)
    const retailAmount = Number(row.retailAmount) || 0;      // Цена после СПП
    const commission = Number(row.ppvzSalesCommission) || 0;
    const delivery = Number(row.deliveryService) || 0;
    const storage = Number(row.paidStorage) || 0;
    const penalty = Number(row.penalty) || 0;
    const deduction = Number(row.deduction) || 0;
    const acceptance = Number(row.paidAcceptance) || 0;
    const additionalPayment = Number(row.additionalPayment) || 0;
    const forPay = Number(row.forPay) || 0;
    const cashbackDiscount = Number(row.cashbackDiscount) || 0;
    const cashbackAmount = Number(row.cashbackAmount) || 0;
    const cashbackCommission = Number(row.cashbackCommissionChange) || 0;
    
    const isSale = docType === 'Продажа';
    const isReturn = docType === 'Возврат';
    const isStorage = sellerOperName === 'Хранение' || storage > 0;
    const isDeduction = sellerOperName === 'Удержания' || deduction > 0;
    const isPenalty = sellerOperName === 'Штраф' || penalty > 0;
    const isLogistics = sellerOperName === 'Логистика' || delivery > 0;
    const isAcceptance = sellerOperName === 'Обработка товара' || acceptance > 0;
    
    if (!result[nmId]) {
      result[nmId] = {
        nmId: nmId,
        salesRetailPrice: 0,    // Реализация до СПП
        salesRetailAmount: 0,   // Продажи после СПП
        returnsRetailPrice: 0,
        returnsRetailAmount: 0,
        quantitySales: 0,
        quantityReturns: 0,
        commission: 0,
        logistics: 0,
        storage: 0,
        penalty: 0,
        deduction: 0,
        acceptance: 0,
        additionalPayment: 0,
        forPay: 0,
        cashbackDiscount: 0,
        cashbackAmount: 0,
        cashbackCommission: 0
      };
    }
    
    const d = result[nmId];
    
    if (isSale) {
      d.salesRetailPrice += retailPrice;
      d.salesRetailAmount += retailAmount;
      d.quantitySales += quantity;
      d.commission += commission;
      d.forPay += forPay;
      d.cashbackDiscount += cashbackDiscount;
      d.cashbackAmount += cashbackAmount;
      d.cashbackCommission += cashbackCommission;
      
      totalRetailPriceSales += retailPrice;
      totalRetailAmountSales += retailAmount;
      totalCommission += commission;
      totalForPay += forPay;
      totalCashbackDiscount += cashbackDiscount;
      totalCashbackAmount += cashbackAmount;
      totalCashbackCommission += cashbackCommission;
      
    } else if (isReturn) {
      d.returnsRetailPrice += retailPrice;
      d.returnsRetailAmount += retailAmount;
      d.quantityReturns += quantity;
      d.commission -= commission;
      d.forPay -= forPay;
      d.cashbackDiscount -= cashbackDiscount;
      d.cashbackAmount -= cashbackAmount;
      d.cashbackCommission -= cashbackCommission;
      
      totalRetailPriceReturns += retailPrice;
      totalRetailAmountReturns += retailAmount;
      
    } else {
      // ДРУГИЕ ОПЕРАЦИИ (логистика, хранение, штрафы и т.д.)
      // ВСЕ РАСХОДЫ СУММИРУЕМ
      if (delivery > 0) {
        d.logistics += delivery;
        totalLogistics += delivery;
      }
      if (storage > 0) {
        d.storage += storage;
        totalStorage += storage;
      }
      if (penalty > 0) {
        d.penalty += penalty;
        totalPenalty += penalty;
      }
      if (deduction > 0) {
        d.deduction += deduction;
        totalDeduction += deduction;
      }
      if (acceptance > 0) {
        d.acceptance += acceptance;
        totalAcceptance += acceptance;
      }
      if (additionalPayment > 0) {
        d.additionalPayment += additionalPayment;
        totalAdditionalPayment += additionalPayment;
      }
    }
  });
  
  // --- ВЫВОД РЕЗУЛЬТАТОВ ---
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📊 ИТОГОВЫЕ СУММЫ ПО ВСЕМ ТОВАРАМ:');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('✅ Реализация до СПП (retailPrice):');
  Logger.log('   Продажи: ' + Math.round(totalRetailPriceSales) + ' ₽');
  Logger.log('   Возвраты: ' + Math.round(totalRetailPriceReturns) + ' ₽');
  Logger.log('   ИТОГО: ' + Math.round(totalRetailPriceSales - totalRetailPriceReturns) + ' ₽');
  Logger.log('');
  Logger.log('✅ Продажи после СПП (retailAmount):');
  Logger.log('   Продажи: ' + Math.round(totalRetailAmountSales) + ' ₽');
  Logger.log('   Возвраты: ' + Math.round(totalRetailAmountReturns) + ' ₽');
  Logger.log('   ИТОГО: ' + Math.round(totalRetailAmountSales - totalRetailAmountReturns) + ' ₽');
  Logger.log('');
  Logger.log('💰 Комиссия: ' + Math.round(totalCommission) + ' ₽');
  Logger.log('💰 К перечислению: ' + Math.round(totalForPay) + ' ₽');
  Logger.log('📦 Логистика: ' + Math.round(totalLogistics) + ' ₽');
  Logger.log('📦 Хранение: ' + Math.round(totalStorage) + ' ₽');
  Logger.log('📦 Штрафы: ' + Math.round(totalPenalty) + ' ₽');
  Logger.log('📦 Удержания: ' + Math.round(totalDeduction) + ' ₽');
  Logger.log('📦 Приемка: ' + Math.round(totalAcceptance) + ' ₽');
  Logger.log('📦 Корректировки: ' + Math.round(totalAdditionalPayment) + ' ₽');
  Logger.log('');
  Logger.log('💰 Кешбэк (discount): ' + Math.round(totalCashbackDiscount) + ' ₽');
  Logger.log('💰 Кешбэк (amount): ' + Math.round(totalCashbackAmount) + ' ₽');
  Logger.log('💰 Кешбэк (commission): ' + Math.round(totalCashbackCommission) + ' ₽');
  Logger.log('');
  Logger.log('📊 Строк без nmId: ' + noNmIdCount);
  Logger.log('📊 Строк с хранением: ' + storageRows);
  Logger.log('📊 Строк с удержаниями: ' + deductionRows);
  Logger.log('📊 Количество товаров в детализации: ' + Object.keys(result).length);
  
  // Показываем топ-5 по продажам
  const sorted = Object.values(result)
    .sort((a, b) => (b.salesRetailAmount - b.returnsRetailAmount) - (a.salesRetailAmount - a.returnsRetailAmount))
    .slice(0, 5);
  
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📊 ТОП-5 по продажам:');
  sorted.forEach((item, i) => {
    const realization = item.salesRetailAmount - item.returnsRetailAmount;
    const qty = item.quantitySales - item.quantityReturns;
    const price = qty > 0 ? Math.round(realization / qty) : 0;
    Logger.log((i+1) + '. nmId=' + item.nmId + ':');
    Logger.log('   Реализация до СПП: ' + Math.round(item.salesRetailPrice - item.returnsRetailPrice) + ' ₽');
    Logger.log('   Продажи после СПП: ' + Math.round(realization) + ' ₽');
    Logger.log('   Кол-во: ' + qty + ' шт');
    Logger.log('   Цена полки: ' + price + ' ₽');
    Logger.log('   Комиссия: ' + Math.round(item.commission) + ' ₽');
    Logger.log('   Логистика: ' + Math.round(item.logistics) + ' ₽');
    Logger.log('   Хранение: ' + Math.round(item.storage) + ' ₽');
    Logger.log('   Удержания: ' + Math.round(item.deduction) + ' ₽');
    Logger.log('   Штрафы: ' + Math.round(item.penalty) + ' ₽');
    Logger.log('   ---');
  });
  
  Logger.log('✅ ===== ТЕСТ ЗАВЕРШЕН =====');
}
