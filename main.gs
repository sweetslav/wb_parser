// ============================================================
// ВЕРСИЯ 6.4 - С ПРОГРЕССОМ В ТАБЛИЦЕ (СТРОКА 3)
// ============================================================

// API эндпоинты
const WB_ORDERS_API = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders';
const WB_ANALYTICS_API = 'https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products';
const WB_CAMPAIGNS_INFO_API = 'https://advert-api.wildberries.ru/api/advert/v2/adverts';

// Задержки между запросами (в миллисекундах)
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
    label: 'Кабинет 1'
  },
  cab2: {
    sheetName: 'adv_effectiveness_2',
    tokenKey: 'WB_TOKEN_2',
    label: 'Кабинет 2'
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
// 4. УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ДЛЯ ЛЮБОГО КАБИНЕТА (С ПРОГРЕССОМ)
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
  
  // --- ФУНКЦИЯ ДЛЯ ОБНОВЛЕНИЯ СТАТУСА (В СТРОКЕ 3) ---
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
  setProgress(0, 5, 'Подготовка...');
  
  // Парсинг дат (B2 и C2 не трогаем)
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
  let totalSteps = 5;
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
    
    // --- ШАГ 5: ОСТАТКИ ---
    currentStep++;
    setProgress(currentStep, totalSteps, 'Остатки...');
    setStatus('Загрузка остатков (5/' + totalSteps + ')', '#FFF3CD');
    Logger.log('📊 5/' + totalSteps + ' Загружаем остатки...');
    const stockData = fetchStockData();
    if (Object.keys(stockData).length === 0) errors.push('Остатки не загружены');
    Utilities.sleep(300);
    SpreadsheetApp.flush();
    
    // --- ОБЪЕДИНЕНИЕ ---
    setStatus('Объединение данных...', '#FFF3CD');
    setProgress(5, 5, 'Объединение...');
    Logger.log('📊 Объединяем данные...');
    const merged = mergeAllData(orders, detail, advCosts, vendorCodes, stockData);
    SpreadsheetApp.flush();
    
    // --- ЗАПИСЬ ---
    setStatus('Запись в таблицу...', '#FFF3CD');
    Logger.log('📊 Записываем данные...');
    writeDataToSheet(sheet, merged);
    SpreadsheetApp.flush();
    
    // --- ФИНИШ ---
    const elapsedSeconds = Math.round((new Date() - startTime) / 1000);
    const totalItems = merged.length;
    const withAd = merged.filter(i => i.hasAd).length;
    const withOrders = merged.filter(i => i.orders > 0).length;
    const totalAdCost = merged.reduce((sum, i) => sum + i.adCost, 0);
    const totalOrders = merged.reduce((sum, i) => sum + i.orders, 0);
    
    // Обновляем статус в строке 3
    const statusText = '✅ Готово! ' + totalItems + ' товаров, ' + elapsedSeconds + ' сек';
    setStatus(statusText, '#D4EDDA');
    sheet.getRange('B3').setValue('✅ Завершено');
    sheet.getRange('C3').setValue(new Date().toLocaleTimeString());
    SpreadsheetApp.flush();
    
    // Показываем алерт
    const message = '✅ ' + config.label + ' готов!\n' +
      '⏱️ ' + elapsedSeconds + ' сек\n' +
      '📦 Товаров: ' + totalItems + '\n' +
      '💰 С рекламой: ' + withAd + '\n' +
      '📦 С заказами: ' + withOrders + '\n' +
      '💰 Общая реклама: ' + Math.round(totalAdCost) + ' ₽\n' +
      '📦 Общие заказы: ' + Math.round(totalOrders) + ' ₽\n' +
      (errors.length > 0 ? '\n⚠️ Ошибки: ' + errors.join(', ') : '');
    
    SpreadsheetApp.getUi().alert(message);
    Logger.log('✅ ===== ФИНИШ: ' + config.label + ' =====');
    
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
// 10. ПОЛУЧЕНИЕ ОСТАТКОВ
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
      const nmIdRaw = row[12]; // M
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
      
      const fbw = Number(row[25]) || 0;   // Z
      const fbs = Number(row[28]) || 0;   // AC
      
      stockData[nmIdNum] = {
        fbw: fbw,
        fbs: fbs,
        total: fbw + fbs
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
// 11. СКЛЕИВАНИЕ ДАННЫХ
// ============================================================
function mergeAllData(orders, detail, adv, vendorCodes, stockData) {
  const allIds = new Set();
  Object.keys(detail).forEach(id => allIds.add(Number(id)));
  Object.keys(orders).forEach(id => allIds.add(Number(id)));
  Object.keys(adv.byNmId || {}).forEach(id => allIds.add(Number(id)));
  
  if (allIds.size === 0) { Logger.log('⚠️ Нет данных'); return []; }
  
  const result = [];
  allIds.forEach(id => {
    const nmId = Number(id);
    const order = orders[nmId] || {};
    const det = detail[nmId] || {};
    const advData = adv.byNmId ? adv.byNmId[nmId] : null;
    const stock = stockData ? stockData[nmId] : null;
    
    const ordersSum = det.ordersSum || 0;
    const clicks = det.clicks || 0;
    const cart = det.cart || 0;
    const vendorCode = vendorCodes[nmId] || '';
    
    let adCost = 0, hasAd = false;
    if (advData && advData.adCost > 0) { adCost = advData.adCost; hasAd = true; }
    
    const fbw = (stock && typeof stock.fbw === 'number') ? stock.fbw : 0;
    const fbs = (stock && typeof stock.fbs === 'number') ? stock.fbs : 0;
    const total = (stock && typeof stock.total === 'number') ? stock.total : 0;
    
    let drr = 0;
    if (hasAd && ordersSum > 0) drr = (adCost / ordersSum) * 100;
    else if (hasAd && ordersSum === 0) drr = 100;
    
    const cr1 = clicks > 0 ? (cart / clicks) * 100 : 0;
    
    result.push({
      nmId: nmId,
      vendorCode: vendorCode,
      orders: ordersSum,
      adCost: adCost,
      drr: drr,
      clicks: clicks,
      cart: cart,
      cr1: cr1,
      hasAd: hasAd,
      fbw: fbw,
      fbs: fbs,
      total: total
    });
  });
  
  result.sort((a, b) => b.orders - a.orders);
  Logger.log('✅ Объединено: ' + result.length + ' товаров');
  return result;
}

// ============================================================
// 12. ЗАПИСЬ В ТАБЛИЦУ
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
    const nmId = safeNum(item.nmId);
    const vendorCode = item.vendorCode || '';
    const orders = safeNum(item.orders);
    const adCost = safeNum(item.adCost);
    const clicks = safeNum(item.clicks);
    const cart = safeNum(item.cart);
    const fbw = safeNum(item.fbw);
    const fbs = safeNum(item.fbs);
    const total = safeNum(item.total);
    
    let drr = '0%';
    if (adCost > 0 && orders > 0) drr = Math.round((adCost / orders) * 10000) / 100 + '%';
    else if (adCost > 0 && orders === 0) drr = '100%';
    
    let cr1 = '0%';
    if (clicks > 0 && cart > 0) cr1 = Math.round((cart / clicks) * 10000) / 100 + '%';
    
    return [
      nmId,
      vendorCode,
      Math.round(orders * 100) / 100,
      Math.round(adCost * 100) / 100,
      drr,
      Math.round(clicks),
      Math.round(cart),
      cr1,
      '',
      '',
      Math.round(fbw),
      Math.round(fbs),
      Math.round(total)
    ];
  });
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) sheet.deleteRows(5, lastRow - 4);
  
  sheet.getRange(5, 1, tableData.length, tableData[0].length).setValues(tableData);
  Logger.log('✅ Записано ' + tableData.length + ' строк');
}

// ============================================================
// 13. УТИЛИТЫ
// ============================================================
function formatDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// ============================================================
// 14. ОЧИСТКА СТАТУСА (СТРОКА 3)
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
// 15. МЕНЮ
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📊 Мои отчеты');
  menu.addItem('🔄 ООО "Ювелир Карат на Савушкина"', 'main');
  menu.addItem('🔄 ИП "Иванова Ю.С."', 'mainCab2');
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
// 16. ТЕСТ ОСТАТКОВ
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
      Logger.log('nmId=' + key + ': FBW=' + d.fbw + ', FBS=' + d.fbs + ', ИТОГО=' + d.total);
    });
  }
}
