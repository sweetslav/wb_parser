// ============================================================
// ВЕРСИЯ 7.0 - ПОЛНАЯ ИНТЕГРАЦИЯ С ФИНАНСОВЫМИ ОТЧЕТАМИ
// ============================================================

// API эндпоинты
const WB_ORDERS_API = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders';
const WB_ANALYTICS_API = 'https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products';
const WB_CAMPAIGNS_INFO_API = 'https://advert-api.wildberries.ru/api/advert/v2/adverts';
const WB_FINANCE_API = 'https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed';

// Задержки между запросами
const DELAYS = {
  BETWEEN_PAGES: 500,
  BETWEEN_APIS: 1000,
  AFTER_ERROR: 2000,
  MAX_RETRIES: 3
};

// ============================================================
// 1. КОНФИГУРАЦИЯ КАБИНЕТОВ
// ============================================================
const CONFIGS = {
  cab1: {
    sheetName: 'adv_effectiveness',
    tokenKey: 'WB_TOKEN',
    label: 'Кабинет 1"'
  },
  cab2: {
    sheetName: 'adv_effectiveness_2',
    tokenKey: 'WB_TOKEN_2',
    label: 'Кабинет 2"'
  }
};

// ============================================================
// 2. УСТАНОВКА ТОКЕНОВ
// ============================================================
function setWBToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔑 Введите токен для первого кабинета',
    'Токен будет сохранён в защищённом хранилище.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token.length > 20) {
      PropertiesService.getScriptProperties().setProperty('WB_TOKEN', token);
      ui.alert('✅ Токен для первого кабинета сохранён!');
    } else {
      ui.alert('❌ Токен слишком короткий.');
    }
  }
}

function setWBToken2() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔑 Введите токен для второго кабинета',
    'Токен будет сохранён в защищённом хранилище.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token.length > 20) {
      PropertiesService.getScriptProperties().setProperty('WB_TOKEN_2', token);
      ui.alert('✅ Токен для второго кабинета сохранён!');
    } else {
      ui.alert('❌ Токен слишком короткий.');
    }
  }
}

// ============================================================
// 3. ГЛАВНЫЕ ФУНКЦИИ ДЛЯ ЗАПУСКА
// ============================================================
function main() { runForCabinet('cab1'); }
function mainCab2() { runForCabinet('cab2'); }

function mainAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Внимание!',
    'Будут обновлены оба кабинета.\nЭто может занять до 10 минут.\n\nПродолжить?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    runForCabinet('cab1');
    runForCabinet('cab2');
    ui.alert('✅ Оба кабинета обновлены!');
  }
}

// ============================================================
// 4. УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ДЛЯ ЛЮБОГО КАБИНЕТА
// ============================================================
function runForCabinet(cabinetId) {
  const startTime = new Date();
  const config = CONFIGS[cabinetId];
  if (!config) { Logger.log('❌ Кабинет ' + cabinetId + ' не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Лист "' + config.sheetName + '" не найден!');
    return;
  }
  
  // --- ФУНКЦИИ ДЛЯ СТАТУСА (СТРОКА 3) ---
  function setStatus(text, color) {
    try {
      sheet.getRange('A3').setValue('🔄 ' + text);
      sheet.getRange('A3').setBackground(color || '#FFF3CD');
      SpreadsheetApp.flush();
    } catch(e) {}
  }
  
  function setProgress(step, total, text) {
    try {
      const percent = Math.round((step / total) * 100);
      const bar = '█'.repeat(Math.floor(percent / 10)) + '░'.repeat(10 - Math.floor(percent / 10));
      sheet.getRange('B3').setValue(percent + '% ' + bar);
      sheet.getRange('C3').setValue(text);
      SpreadsheetApp.flush();
    } catch(e) {}
  }
  
  // --- ИНИЦИАЛИЗАЦИЯ ---
  setStatus('Запуск...', '#FFF3CD');
  setProgress(0, 7, 'Подготовка...');
  
  // Парсинг дат
  const dateFromRaw = sheet.getRange('B2').getValue();
  const dateToRaw = sheet.getRange('C2').getValue();
  let dateFrom, dateTo;
  
  try {
    dateFrom = new Date(dateFromRaw);
    dateTo = new Date(dateToRaw);
    if (isNaN(dateFrom.getTime()) || isNaN(dateTo.getTime())) throw new Error('Неверный формат даты');
    if (dateFrom > dateTo) throw new Error('Дата "с" больше даты "по"');
  } catch (e) {
    setStatus('❌ Ошибка в датах!', '#F8D7DA');
    SpreadsheetApp.getUi().alert('❌ Ошибка в датах!\n\n' + e.message);
    return;
  }
  
  const token = PropertiesService.getScriptProperties().getProperty(config.tokenKey);
  if (!token) {
    setStatus('❌ Токен не найден!', '#F8D7DA');
    SpreadsheetApp.getUi().alert('❌ Токен для ' + config.label + ' не найден!');
    return;
  }
  
  setStatus('Загрузка данных...', '#FFF3CD');
  Logger.log('🚀 ===== СТАРТ: ' + config.label + ' =====');
  
  const errors = [];
  const totalSteps = 7;
  let currentStep = 0;
  
  try {
    // --- ШАГ 1: ЗАКАЗЫ ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Заказы...');
    setStatus('Загрузка заказов (1/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 1/' + totalSteps + ' Загружаем заказы...');
    const orders = fetchWBOrders(dateFrom, token);
    if (Object.keys(orders).length === 0) errors.push('Заказы не загружены');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 2: АНАЛИТИКА ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Аналитика...');
    setStatus('Загрузка аналитики (2/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 2/' + totalSteps + ' Загружаем аналитику...');
    const detail = fetchAnalytics(dateFrom, dateTo, token);
    if (Object.keys(detail).length === 0) errors.push('Аналитика не загружена');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 3: РЕКЛАМА ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Реклама...');
    setStatus('Загрузка рекламы (3/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 3/' + totalSteps + ' Загружаем рекламу...');
    const advCosts = fetchAdvCosts(dateFrom, dateTo, token);
    if (Object.keys(advCosts.byNmId || {}).length === 0) errors.push('Реклама не загружена');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 4: VENDORCODE ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Артикулы...');
    setStatus('Загрузка артикулов (4/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 4/' + totalSteps + ' Загружаем артикулы...');
    const vendorCodes = fetchAllVendorCodes(token);
    if (Object.keys(vendorCodes).length === 0) errors.push('VendorCode не загружены');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 5: ОСТАТКИ + СЕБЕСТОИМОСТЬ ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Остатки...');
    setStatus('Загрузка остатков (5/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 5/' + totalSteps + ' Загружаем остатки...');
    const stockData = fetchStockData();
    if (Object.keys(stockData).length === 0) errors.push('Остатки не загружены');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 6: ФИНАНСЫ ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Финансы...');
    setStatus('Загрузка финансов (6/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 6/' + totalSteps + ' Загружаем финансы...');
    const financialData = fetchFinancialData(dateFrom, dateTo, token);
    if (Object.keys(financialData).length === 0) errors.push('Финансы не загружены');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ШАГ 7: ОБЪЕДИНЕНИЕ И ЗАПИСЬ ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Объединение...');
    setStatus('Объединение данных (7/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 7/' + totalSteps + ' Объединяем данные...');
    
    const merged = mergeAllData(orders, detail, advCosts, vendorCodes, stockData, financialData);
    SpreadsheetApp.flush();
    
    setStatus('Запись в таблицу...', '#FFF3CD');
    writeDataToSheet(sheet, merged);
    SpreadsheetApp.flush();
    
    // --- ФИНИШ ---
    const elapsedSeconds = Math.round((new Date() - startTime) / 1000);
    const totalItems = merged.length;
    const withAd = merged.filter(i => i.adCost > 0).length;
    const withOrders = merged.filter(i => i.ordersSum > 0).length;
    const totalAdCost = merged.reduce((sum, i) => sum + i.adCost, 0);
    const totalOrders = merged.reduce((sum, i) => sum + i.ordersSum, 0);
    const totalRealization = merged.reduce((sum, i) => sum + i.totalRealization, 0);
    
    const statusText = '✅ Готово! ' + totalItems + ' товаров, ' + elapsedSeconds + ' сек';
    setStatus(statusText, '#D4EDDA');
    sheet.getRange('B3').setValue('✅ Завершено');
    sheet.getRange('C3').setValue(new Date().toLocaleTimeString());
    SpreadsheetApp.flush();
    
    const message = '✅ ' + config.label + ' готов!\n' +
      '⏱️ ' + elapsedSeconds + ' сек\n' +
      '📦 Товаров: ' + totalItems + '\n' +
      '💰 С рекламой: ' + withAd + '\n' +
      '📦 С заказами: ' + withOrders + '\n' +
      '💰 Общая реклама: ' + Math.round(totalAdCost) + ' ₽\n' +
      '📦 Общие заказы: ' + Math.round(totalOrders) + ' ₽\n' +
      '💰 Реализация: ' + Math.round(totalRealization) + ' ₽\n' +
      (errors.length > 0 ? '\n⚠️ Ошибки: ' + errors.join(', ') : '');
    
    SpreadsheetApp.getUi().alert(message);
    Logger.log('✅ ===== ФИНИШ: ' + config.label + ' (' + elapsedSeconds + ' сек) =====');
    
  } catch (e) {
    setStatus('❌ Ошибка: ' + e.message, '#F8D7DA');
    Logger.log('❌ Ошибка: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка в ' + config.label + '!\n\n' + e.message);
  }
}

// ============================================================
// 5. ЗАПРОС ЗАКАЗОВ
// ============================================================
function fetchWBOrders(dateFrom, token) {
  const from = formatDate(dateFrom);
  let allOrders = [], page = 0, limit = 1000, hasMore = true, retryCount = 0;
  
  while (hasMore) {
    const url = WB_ORDERS_API + '?dateFrom=' + from + '&limit=' + limit + '&offset=' + (page * limit);
    const options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 429) {
        retryCount++;
        if (retryCount <= DELAYS.MAX_RETRIES) {
          Utilities.sleep(DELAYS.AFTER_ERROR * retryCount);
          continue;
        }
        break;
      }
      if (response.getResponseCode() !== 200) break;
      
      const data = JSON.parse(response.getContentText());
      if (!Array.isArray(data) || data.length === 0) break;
      
      allOrders = allOrders.concat(data);
      if (data.length < limit) break;
      page++;
      retryCount = 0;
      Utilities.sleep(DELAYS.BETWEEN_PAGES);
    } catch (e) {
      Logger.log('❌ Ошибка Orders: ' + e.message);
      break;
    }
  }
  
  const result = {};
  allOrders.forEach(item => {
    const nmId = item.nmId;
    if (!nmId || item.canceled || item.cancel) return;
    if (!result[nmId]) {
      result[nmId] = { nmId: nmId, name: item.supplierArticle || item.barcode || 'Без названия' };
    }
  });
  
  Logger.log('✅ Orders: ' + Object.keys(result).length + ' товаров');
  return result;
}

// ============================================================
// 6. ЗАПРОС К АНАЛИТИКЕ
// ============================================================
function fetchAnalytics(dateFrom, dateTo, token) {
  const from = formatDate(dateFrom), to = formatDate(dateTo);
  let allProducts = [], offset = 0, limit = 1000, hasMore = true, retryCount = 0;
  
  while (hasMore) {
    const payload = { selectedPeriod: { start: from, end: to }, limit: limit, offset: offset };
    const options = {
      method: 'POST',
      headers: { 'Authorization': token, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(WB_ANALYTICS_API, options);
      if (response.getResponseCode() === 429) {
        retryCount++;
        if (retryCount <= DELAYS.MAX_RETRIES) {
          Utilities.sleep(DELAYS.AFTER_ERROR * retryCount);
          continue;
        }
        break;
      }
      if (response.getResponseCode() !== 200) break;
      
      const data = JSON.parse(response.getContentText());
      const products = data.data?.products || [];
      if (products.length === 0) break;
      
      allProducts = allProducts.concat(products);
      if (products.length < limit) break;
      offset += limit;
      retryCount = 0;
      Utilities.sleep(DELAYS.BETWEEN_PAGES);
    } catch (e) {
      Logger.log('❌ Ошибка Analytics: ' + e.message);
      break;
    }
  }
  
  const result = {};
  allProducts.forEach(item => {
    const nmId = item.product?.nmId;
    if (!nmId) return;
    const selected = item.statistic?.selected || {};
    result[nmId] = {
      nmId: nmId,
      clicks: selected.openCount || 0,
      cart: selected.cartCount || 0,
      ordersSum: selected.orderSum || 0
    };
  });
  
  Logger.log('✅ Analytics: ' + Object.keys(result).length + ' товаров');
  return result;
}

// ============================================================
// 7. ЗАПРОС РЕКЛАМЫ
// ============================================================
function fetchAdvCosts(dateFrom, dateTo, token) {
  const from = formatDate(dateFrom), to = formatDate(dateTo);
  const options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
  let retryCount = 0;
  
  try {
    const updUrl = 'https://advert-api.wildberries.ru/adv/v1/upd?from=' + from + '&to=' + to;
    let updResponse = null, code = 0;
    
    while (retryCount < DELAYS.MAX_RETRIES) {
      updResponse = UrlFetchApp.fetch(updUrl, options);
      code = updResponse.getResponseCode();
      if (code === 429) {
        retryCount++;
        Utilities.sleep(DELAYS.AFTER_ERROR * retryCount);
        continue;
      }
      break;
    }
    
    if (code !== 200) return { byNmId: {}, byCampaign: {}, totalAdCost: 0 };
    
    const data = JSON.parse(updResponse.getContentText());
    if (!Array.isArray(data) || data.length === 0) return { byNmId: {}, byCampaign: {}, totalAdCost: 0 };
    
    const costsByCampaign = {};
    data.forEach(item => {
      const advertId = item.advertId;
      const sum = item.updSum || 0;
      if (!costsByCampaign[advertId]) costsByCampaign[advertId] = 0;
      costsByCampaign[advertId] += sum;
    });
    
    Utilities.sleep(DELAYS.BETWEEN_APIS);
    const campaignsInfo = fetchCampaignsInfo(token);
    
    const result = {};
    Object.keys(costsByCampaign).forEach(advertId => {
      const cost = costsByCampaign[advertId];
      const campaign = campaignsInfo[advertId];
      if (campaign && campaign.nmIds && campaign.nmIds.length > 0) {
        const perItem = cost / campaign.nmIds.length;
        campaign.nmIds.forEach(nmId => {
          if (!result[nmId]) result[nmId] = { nmId: nmId, adCost: 0, campaignIds: [] };
          result[nmId].adCost += perItem;
          if (!result[nmId].campaignIds.includes(advertId)) {
            result[nmId].campaignIds.push(advertId);
          }
        });
      }
    });
    
    const byNmId = {};
    Object.keys(result).forEach(nmId => {
      byNmId[nmId] = { adCost: Math.round(result[nmId].adCost * 100) / 100, hasAd: true };
    });
    
    return { byNmId: byNmId, byCampaign: costsByCampaign, totalAdCost: 0 };
    
  } catch (e) {
    Logger.log('❌ Ошибка рекламы: ' + e.message);
    return { byNmId: {}, byCampaign: {}, totalAdCost: 0 };
  }
}

// ============================================================
// 8. ИНФОРМАЦИЯ О КАМПАНИЯХ
// ============================================================
function fetchCampaignsInfo(token) {
  const options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
  let retryCount = 0;
  
  try {
    const url = WB_CAMPAIGNS_INFO_API + '?statuses=4,7,9,11';
    let response = null, code = 0;
    
    while (retryCount < DELAYS.MAX_RETRIES) {
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      if (code === 429) {
        retryCount++;
        Utilities.sleep(DELAYS.AFTER_ERROR * retryCount);
        continue;
      }
      break;
    }
    
    if (code !== 200) return {};
    const data = JSON.parse(response.getContentText());
    const campaigns = data.adverts || [];
    
    const result = {};
    campaigns.forEach(campaign => {
      const nmIds = [];
      if (campaign.nm_settings) campaign.nm_settings.forEach(item => { if (item.nm_id) nmIds.push(item.nm_id); });
      if (campaign.nms) campaign.nms.forEach(item => { if (item.nmId) nmIds.push(item.nmId); });
      if (campaign.items) campaign.items.forEach(item => { if (item.nmId) nmIds.push(item.nmId); });
      if (campaign.nmIds) campaign.nmIds.forEach(id => nmIds.push(id));
      
      const uniqueNmIds = [...new Set(nmIds)];
      if (uniqueNmIds.length > 0) {
        result[campaign.id] = { advertId: campaign.id, name: campaign.settings?.name || 'Без названия', nmIds: uniqueNmIds };
      }
    });
    
    return result;
  } catch (e) {
    return {};
  }
}

// ============================================================
// 9. ПОЛУЧЕНИЕ VENDORCODE
// ============================================================
function fetchAllVendorCodes(token) {
  const options = { method: 'GET', headers: { 'Authorization': token }, muteHttpExceptions: true };
  let allGoods = [], offset = 0, limit = 1000, hasMore = true, retryCount = 0;
  
  try {
    while (hasMore) {
      const url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=' + limit + '&offset=' + offset;
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 429) {
        retryCount++;
        if (retryCount <= DELAYS.MAX_RETRIES) {
          Utilities.sleep(DELAYS.AFTER_ERROR * retryCount);
          continue;
        }
        break;
      }
      if (response.getResponseCode() !== 200) break;
      
      const data = JSON.parse(response.getContentText());
      const goods = data.data?.listGoods || [];
      if (goods.length === 0) break;
      
      allGoods = allGoods.concat(goods);
      if (goods.length < limit) break;
      offset += limit;
      retryCount = 0;
      Utilities.sleep(DELAYS.BETWEEN_PAGES);
    }
    
    const result = {};
    allGoods.forEach(item => {
      const nmId = item.nmID;
      if (nmId && item.vendorCode) result[nmId] = item.vendorCode;
    });
    
    Logger.log('✅ VendorCode: ' + Object.keys(result).length + ' артикулов');
    return result;
  } catch (e) {
    Logger.log('❌ Ошибка vendorCode: ' + e.message);
    return {};
  }
}

// ============================================================
// 10. ПОЛУЧЕНИЕ ОСТАТКОВ И СЕБЕСТОИМОСТИ
// ============================================================
function fetchStockData() {
  try {
    const SPREADSHEET_ID = '1pP1RlNjgfxcDNw9Icwep0Pl3PyJikNIeQE6bMdCDD70';
    const SHEET_NAME = 'unit расчет';
    
    const stockBook = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = stockBook.getSheetByName(SHEET_NAME);
    if (!sheet) { Logger.log('⚠️ Лист не найден'); return {}; }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('⚠️ Нет данных'); return {}; }
    
    const range = sheet.getRange(2, 1, lastRow - 1, 29);
    const values = range.getValues();
    
    const stockData = {};
    let foundCount = 0;
    
    values.forEach(row => {
      const nmIdRaw = row[12];
      let nmIdNum = null;
      const type = typeof nmIdRaw;
      
      if (type === 'number') nmIdNum = nmIdRaw;
      else if (type === 'string') {
        const cleaned = nmIdRaw.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      } else if (nmIdRaw && type === 'object') {
        const strValue = String(nmIdRaw);
        const cleaned = strValue.replace(/[^0-9]/g, '');
        if (cleaned) nmIdNum = Number(cleaned);
      }
      
      if (!nmIdNum || isNaN(nmIdNum) || nmIdNum === 0) return;
      
      const fbw = Number(row[25]) || 0;
      const fbs = Number(row[28]) || 0;
      const costPrice = Number(row[24]) || 0;
      
      stockData[nmIdNum] = {
        fbw: fbw,
        fbs: fbs,
        total: fbw + fbs,
        costPrice: costPrice
      };
      foundCount++;
    });
    
    Logger.log('✅ Остатки: ' + foundCount + ' товаров');
    return stockData;
  } catch (e) {
    Logger.log('❌ Ошибка остатков: ' + e.message);
    return {};
  }
}

// ============================================================
// 11. ПОЛУЧЕНИЕ ФИНАНСОВЫХ ДАННЫХ
// ============================================================
function fetchFinancialData(dateFrom, dateTo, token) {
  Logger.log('📊 Загружаем финансовые данные...');
  
  try {
    let allRows = [];
    let rrdId = 0;
    let hasMore = true;
    let attempt = 0;
    const maxAttempts = 100;
    
    while (hasMore && attempt < maxAttempts) {
      attempt++;
      
      const payload = {
        dateFrom: formatDate(dateFrom),
        dateTo: formatDate(dateTo),
        limit: 100000,
        rrdId: rrdId,
        period: 'weekly'
      };
      
      const options = {
        method: 'POST',
        headers: {
          'Authorization': token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(WB_FINANCE_API, options);
      const code = response.getResponseCode();
      
      if (code === 204) {
        hasMore = false;
        break;
      }
      
      if (code !== 200) {
        Logger.log('⚠️ Ошибка финансов: ' + code);
        hasMore = false;
        break;
      }
      
      const data = JSON.parse(response.getContentText());
      
      if (!Array.isArray(data) || data.length === 0) {
        hasMore = false;
        break;
      }
      
      allRows = allRows.concat(data);
      
      if (data.length < 100000) {
        hasMore = false;
      } else {
        const lastRow = data[data.length - 1];
        rrdId = lastRow.rrdId || 0;
        if (rrdId === 0) hasMore = false;
      }
      
      Utilities.sleep(300);
    }
    
    if (allRows.length === 0) {
      Logger.log('❌ Нет финансовых данных');
      return {};
    }
    
    Logger.log('📊 Финансы: ' + allRows.length + ' строк');
    
    // Агрегируем
    const aggregated = aggregateFinancialData(allRows);
    Logger.log('✅ Финансы агрегированы: ' + Object.keys(aggregated).length + ' товаров');
    
    return aggregated;
    
  } catch (e) {
    Logger.log('❌ Ошибка финансов: ' + e.message);
    return {};
  }
}

// ============================================================
// 12. АГРЕГАЦИЯ ФИНАНСОВЫХ ДАННЫХ
// ============================================================
function aggregateFinancialData(rows) {
  const result = {};
  
  rows.forEach(row => {
    const nmId = row.nmId;
    const docType = row.docTypeName || '';
    const sellerOperName = row.sellerOperName || '';
    const quantity = Number(row.quantity) || 0;
    
    const retailPrice = Number(row.retailPrice) || 0;
    const retailAmount = Number(row.retailAmount) || 0;
    const commission = Number(row.ppvzSalesCommission) || 0;
    const delivery = Number(row.deliveryService) || 0;
    const storage = Number(row.paidStorage) || 0;
    const penalty = Number(row.penalty) || 0;
    const deduction = Number(row.deduction) || 0;
    const acceptance = Number(row.paidAcceptance) || 0;
    const additionalPayment = Number(row.additionalPayment) || 0;
    const rebillLogistic = Number(row.rebillLogisticCost) || 0;
    const forPay = Number(row.forPay) || 0;
    const cashbackDiscount = Number(row.cashbackDiscount) || 0;
    const cashbackAmount = Number(row.cashbackAmount) || 0;
    const cashbackCommission = Number(row.cashbackCommissionChange) || 0;
    
    const isSale = docType === 'Продажа';
    const isReturn = docType === 'Возврат';
    
    // Если нет nmId - суммируем общие расходы
    if (!nmId) {
      // Для общих расходов создаем специальный ключ 'total'
      if (!result['total']) {
        result['total'] = {
          nmId: 'total',
          salesRetailPrice: 0,
          salesRetailAmount: 0,
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
          rebillLogistic: 0,
          forPay: 0,
          cashbackDiscount: 0,
          cashbackAmount: 0,
          cashbackCommission: 0
        };
      }
      const d = result['total'];
      d.logistics += delivery;
      d.storage += storage;
      d.penalty += penalty;
      d.deduction += deduction;
      d.acceptance += acceptance;
      d.additionalPayment += additionalPayment;
      d.rebillLogistic += rebillLogistic;
      return;
    }
    
    if (!result[nmId]) {
      result[nmId] = {
        nmId: nmId,
        salesRetailPrice: 0,
        salesRetailAmount: 0,
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
        rebillLogistic: 0,
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
    } else if (isReturn) {
      d.returnsRetailPrice += retailPrice;
      d.returnsRetailAmount += retailAmount;
      d.quantityReturns += quantity;
      d.commission -= commission;
      d.forPay -= forPay;
      d.cashbackDiscount -= cashbackDiscount;
      d.cashbackAmount -= cashbackAmount;
      d.cashbackCommission -= cashbackCommission;
    } else {
      // Расходы (логистика, хранение, штрафы и т.д.)
      d.logistics += delivery;
      d.storage += storage;
      d.penalty += penalty;
      d.deduction += deduction;
      d.acceptance += acceptance;
      d.additionalPayment += additionalPayment;
      d.rebillLogistic += rebillLogistic;
    }
  });
  
  // Удаляем пустые
  Object.keys(result).forEach(key => {
    const d = result[key];
    if (d.salesRetailAmount === 0 && d.returnsRetailAmount === 0 && 
        d.logistics === 0 && d.storage === 0 && d.penalty === 0 && d.deduction === 0) {
      delete result[key];
    }
  });
  
  return result;
}

// ============================================================
// 13. СКЛЕИВАНИЕ ВСЕХ ДАННЫХ (ИСПРАВЛЕННАЯ)
// ============================================================
function mergeAllData(orders, detail, adv, vendorCodes, stockData, financialData) {
  const allIds = new Set();
  Object.keys(detail).forEach(id => allIds.add(Number(id)));
  Object.keys(orders).forEach(id => allIds.add(Number(id)));
  Object.keys(adv.byNmId || {}).forEach(id => allIds.add(Number(id)));
  
  if (allIds.size === 0) { Logger.log('⚠️ Нет данных для объединения'); return []; }
  
  const result = [];
  
  allIds.forEach(id => {
    const nmId = Number(id);
    const order = orders[nmId] || {};
    const det = detail[nmId] || {};
    const advData = adv.byNmId ? adv.byNmId[nmId] : null;
    const stock = stockData ? stockData[nmId] : null;
    const fin = financialData ? financialData[nmId] : null;
    
    // --- ИЗ АНАЛИТИКИ ---
    const ordersSum = det.ordersSum || 0;
    const clicks = det.clicks || 0;
    const cart = det.cart || 0;
    const vendorCode = vendorCodes[nmId] || '';
    
    // --- ИЗ РЕКЛАМЫ ---
    let adCost = 0;
    if (advData && advData.adCost > 0) {
      adCost = advData.adCost;
    }
    
    // --- ИЗ ФИНАНСОВ ---
    // ✅ ИСПРАВЛЕНО:
    // totalSales = retailPrice (Реализация до СПП)
    const totalSales = fin ? (fin.salesRetailPrice - fin.returnsRetailPrice) : 0;
    
    // totalRealization = retailAmount (Продажи после СПП)
    const totalRealization = fin ? (fin.salesRetailAmount - fin.returnsRetailAmount) : 0;
    
    const totalQuantitySales = fin ? fin.quantitySales : 0;
    const totalQuantity = fin ? (fin.quantitySales - fin.quantityReturns) : 0;
    const totalCommission = fin ? fin.commission : 0;
    const totalLogistics = fin ? fin.logistics : 0;
    const totalStorage = fin ? fin.storage : 0;
    const totalPenalty = fin ? fin.penalty : 0;
    const totalDeduction = fin ? fin.deduction : 0;
    const totalAcceptance = fin ? fin.acceptance : 0;
    const totalAdditionalPayment = fin ? fin.additionalPayment : 0;
    const totalRebillLogistic = fin ? fin.rebillLogistic : 0;
    
    // --- ИЗ СПРАВОЧНИКА ---
    const fbw = stock ? stock.fbw : 0;
    const fbs = stock ? stock.fbs : 0;
    const totalStock = stock ? stock.total : 0;
    const costPrice = stock ? stock.costPrice : 0;
    
    // --- РАСЧЕТЫ ---
    // Цена полки = Продажи после СПП / Количество продаж
    const shelfPrice = totalQuantitySales > 0 ? Math.round(totalRealization / totalQuantitySales) : 0;
    
    // Себестоимость = себест100 × Количество продаж
    const totalCostPrice = costPrice * totalQuantitySales;
    
    // Компенсации
    const totalCompensations = totalAdditionalPayment + totalRebillLogistic;
    
    // ДРР
    let drr = 0;
    if (adCost > 0 && ordersSum > 0) {
      drr = (adCost / ordersSum) * 100;
    } else if (adCost > 0 && ordersSum === 0) {
      drr = 100;
    }
    
    // CR1
    const cr1 = clicks > 0 ? (cart / clicks) * 100 : 0;
    
    // --- ВАЛОВАЯ МАРЖА ---
    // Продажи после СПП - Себестоимость - Логистика - Комиссия - Штрафы - Хранение - Удержания - Приемка - Реклама + Компенсации
    const grossMargin = totalRealization 
      - totalCostPrice 
      - totalLogistics 
      - totalCommission 
      - totalPenalty 
      - totalStorage 
      - totalDeduction 
      - totalAcceptance 
      - adCost 
      + totalCompensations;
    
    const marginPercent = totalRealization > 0 ? (grossMargin / totalRealization) * 100 : 0;
    
    result.push({
      nmId: nmId,
      vendorCode: vendorCode,
      ordersSum: ordersSum,
      adCost: adCost,
      drr: drr,
      clicks: clicks,
      cart: cart,
      cr1: cr1,
      fbw: fbw,
      fbs: fbs,
      totalStock: totalStock,
      shelfPrice: shelfPrice,
      totalSales: totalSales,              // M: Реализация до СПП (retailPrice)
      totalRealization: totalRealization,  // N: Продажи после СПП (retailAmount)
      totalQuantitySales: totalQuantitySales,
      totalCostPrice: totalCostPrice,
      totalLogistics: totalLogistics,
      totalCommission: totalCommission,
      totalPenalty: totalPenalty,
      totalStorage: totalStorage,
      totalDeduction: totalDeduction,
      totalAcceptance: totalAcceptance,
      totalCompensations: totalCompensations,
      grossMargin: grossMargin,
      marginPercent: marginPercent
    });
  });
  
  result.sort((a, b) => b.totalRealization - a.totalRealization);
  Logger.log('✅ Объединено: ' + result.length + ' товаров');
  return result;
}

// ============================================================
// 14. ЗАПИСЬ В ТАБЛИЦУ (26 СТОЛБЦОВ A-Z)
// ============================================================
function writeDataToSheet(sheet, data) {
  if (!data || data.length === 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('⚠️ Нет данных!', 'Очистить старые данные?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 5) sheet.deleteRows(5, lastRow - 4);
    }
    return;
  }
  
  function safeNum(v) { 
    if (v === undefined || v === null || v === '') return 0;
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  }
  
  const tableData = data.map(item => {
    const marginValue = Math.round(safeNum(item.marginPercent));
    const marginFormatted = String(marginValue).replace('.', ',') + '%';
    
    return [
      safeNum(item.nmId),
      item.vendorCode || '',
      Math.round(safeNum(item.ordersSum)),
      Math.round(safeNum(item.adCost)),
      safeNum(item.drr) > 0 ? Math.round(safeNum(item.drr)) + '%' : '0%',
      Math.round(safeNum(item.clicks)),
      Math.round(safeNum(item.cart)),
      safeNum(item.cr1) > 0 ? Math.round(safeNum(item.cr1)) + '%' : '0%',
      Math.round(safeNum(item.fbw)),
      Math.round(safeNum(item.fbs)),
      Math.round(safeNum(item.totalStock)),
      Math.round(safeNum(item.shelfPrice)),
      Math.round(safeNum(item.totalSales)),
      Math.round(safeNum(item.totalRealization)),
      Math.round(safeNum(item.totalQuantitySales)),
      Math.round(safeNum(item.totalCostPrice)),
      Math.round(safeNum(item.totalLogistics)),
      Math.round(safeNum(item.totalCommission)),
      Math.round(safeNum(item.totalPenalty)),
      Math.round(safeNum(item.totalStorage)),
      Math.round(safeNum(item.totalDeduction)),
      Math.round(safeNum(item.totalAcceptance)),
      Math.round(safeNum(item.adCost)),
      Math.round(safeNum(item.totalCompensations)),
      Math.round(safeNum(item.grossMargin)),
      marginFormatted
    ];
  });
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) sheet.deleteRows(5, lastRow - 4);
  
  sheet.getRange(5, 1, tableData.length, tableData[0].length).setValues(tableData);
  Logger.log('✅ Записано ' + tableData.length + ' строк, 26 столбцов');
}

// ============================================================
// 15. УТИЛИТЫ
// ============================================================
function formatDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// ============================================================
// 16. ОЧИСТКА СТАТУСА
// ============================================================
function clearStatus() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Очистить статусы?',
    'Будут очищены ячейки A3, B3, C3 на всех листах.\n\nЯчейки с датами (B2, C2) НЕ будут затронуты.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ['adv_effectiveness', 'adv_effectiveness_2'];
    
    sheets.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        sheet.getRange('A3').clearContent();
        sheet.getRange('A3').setBackground(null);
        sheet.getRange('B3').clearContent();
        sheet.getRange('C3').clearContent();
      }
    });
    
    ui.alert('✅ Статусы очищены!');
  }
}

// ============================================================
// 17. МЕНЮ
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📊 Мои отчеты');
  menu.addItem('🔄 Кабинет 1', 'main');
  menu.addItem('🔄 Кабинет 2', 'mainCab2');
  menu.addSeparator();
  menu.addItem('🔄 Обновить все', 'mainAll');
  menu.addSeparator();
  menu.addItem('🔑 Токен кабинет 1', 'setWBToken');
  menu.addItem('🔑 Токен кабинет 2', 'setWBToken2');
  menu.addSeparator();
  menu.addItem('🧹 Очистить статусы', 'clearStatus');
  menu.addToUi();
}

// ============================================================
// 18. ТЕСТОВЫЕ ФУНКЦИИ
// ============================================================
function testStockData() {
  Logger.log('🔍 Тест загрузки остатков');
  const data = fetchStockData();
  const keys = Object.keys(data);
  Logger.log('✅ Загружено: ' + keys.length + ' товаров');
  
  if (keys.length > 0) {
    const sample = keys.slice(0, 5);
    sample.forEach(key => {
      const d = data[key];
      Logger.log('nmId=' + key + ': FBW=' + d.fbw + ', FBS=' + d.fbs + ', ИТОГО=' + d.total + ', себест=' + d.costPrice);
    });
  }
}

function testFinancialAggregation() {
  Logger.log('🔍 ===== ТЕСТ АГРЕГАЦИИ ФИНАНСОВ =====');
  
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) { Logger.log('❌ Токен не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adv_effectiveness');
  const dateFrom = new Date(sheet.getRange('B2').getValue());
  const dateTo = new Date(sheet.getRange('C2').getValue());
  
  Logger.log('📅 Период: ' + formatDate(dateFrom) + ' - ' + formatDate(dateTo));
  
  const financialData = fetchFinancialData(dateFrom, dateTo, token);
  const keys = Object.keys(financialData).filter(k => k !== 'total');
  
  Logger.log('✅ Загружено: ' + keys.length + ' товаров');
  
  let totalRealization = 0;
  let totalLogistics = 0;
  let totalStorage = 0;
  let totalPenalty = 0;
  let totalDeduction = 0;
  let totalCommission = 0;
  
  keys.forEach(key => {
    const d = financialData[key];
    totalRealization += d.salesRetailAmount - d.returnsRetailAmount;
    totalLogistics += d.logistics;
    totalStorage += d.storage;
    totalPenalty += d.penalty;
    totalDeduction += d.deduction;
    totalCommission += d.commission;
  });
  
  // Добавляем общие расходы
  if (financialData['total']) {
    const t = financialData['total'];
    totalLogistics += t.logistics;
    totalStorage += t.storage;
    totalPenalty += t.penalty;
    totalDeduction += t.deduction;
  }
  
  Logger.log('📊 ИТОГО:');
  Logger.log('   Реализация: ' + Math.round(totalRealization) + ' ₽');
  Logger.log('   Логистика: ' + Math.round(totalLogistics) + ' ₽');
  Logger.log('   Хранение: ' + Math.round(totalStorage) + ' ₽');
  Logger.log('   Штрафы: ' + Math.round(totalPenalty) + ' ₽');
  Logger.log('   Удержания: ' + Math.round(totalDeduction) + ' ₽');
  Logger.log('   Комиссия: ' + Math.round(totalCommission) + ' ₽');
}

function debugPenalty() {
  Logger.log('🔍 ===== ОТЛАДКА ШТРАФОВ =====');
  
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) { Logger.log('❌ Токен не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adv_effectiveness');
  const dateFrom = new Date(sheet.getRange('B2').getValue());
  const dateTo = new Date(sheet.getRange('C2').getValue());
  
  const financialData = fetchFinancialData(dateFrom, dateTo, token);
  const keys = Object.keys(financialData).filter(k => k !== 'total');
  
  let totalPenalty = 0;
  keys.forEach(key => {
    totalPenalty += financialData[key].penalty;
  });
  if (financialData['total']) {
    totalPenalty += financialData['total'].penalty;
  }
  
  Logger.log('📊 Общая сумма штрафов: ' + totalPenalty + ' ₽');
  Logger.log('✅ ===== ОТЛАДКА ЗАВЕРШЕНА =====');
}

function debugFinancialRow(nmIdToCheck) {
  Logger.log('🔍 ===== ОТЛАДКА ДЛЯ nmId=' + nmIdToCheck + ' =====');
  
  const token = PropertiesService.getScriptProperties().getProperty('WB_TOKEN');
  if (!token) { Logger.log('❌ Токен не найден'); return; }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adv_effectiveness');
  const dateFrom = new Date(sheet.getRange('B2').getValue());
  const dateTo = new Date(sheet.getRange('C2').getValue());
  
  const financialData = fetchFinancialData(dateFrom, dateTo, token);
  
  if (!financialData[nmIdToCheck]) {
    Logger.log('❌ nmId ' + nmIdToCheck + ' не найден');
    return;
  }
  
  const d = financialData[nmIdToCheck];
  Logger.log('📊 Данные для nmId=' + nmIdToCheck + ':');
  Logger.log('   Продажи (retailPrice): ' + Math.round(d.salesRetailPrice) + ' ₽');
  Logger.log('   Продажи (retailAmount): ' + Math.round(d.salesRetailAmount) + ' ₽');
  Logger.log('   Возвраты (retailPrice): ' + Math.round(d.returnsRetailPrice) + ' ₽');
  Logger.log('   Возвраты (retailAmount): ' + Math.round(d.returnsRetailAmount) + ' ₽');
  Logger.log('   Кол-во продаж: ' + d.quantitySales + ' шт');
  Logger.log('   Кол-во возвратов: ' + d.quantityReturns + ' шт');
  Logger.log('   Комиссия: ' + Math.round(d.commission) + ' ₽');
  Logger.log('   Логистика: ' + Math.round(d.logistics) + ' ₽');
  Logger.log('   Хранение: ' + Math.round(d.storage) + ' ₽');
  Logger.log('   Штрафы: ' + Math.round(d.penalty) + ' ₽');
  Logger.log('   Удержания: ' + Math.round(d.deduction) + ' ₽');
  Logger.log('   Приемка: ' + Math.round(d.acceptance) + ' ₽');
  
  const realization = d.salesRetailAmount - d.returnsRetailAmount;
  const qty = d.quantitySales - d.quantityReturns;
  const price = qty > 0 ? Math.round(realization / qty) : 0;
  Logger.log('   Реализация: ' + Math.round(realization) + ' ₽');
  Logger.log('   Цена полки: ' + price + ' ₽');
  
  Logger.log('✅ ===== ОТЛАДКА ЗАВЕРШЕНА =====');
}

function runDebug() {
  debugFinancialRow(158868824);
}
