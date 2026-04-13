/**
 * 📊 포트폴리오 트래커 - Google Apps Script 백엔드
 *
 * ── 설정 방법 ──────────────────────────────────────────────────
 * 1. Google Sheets에서 Extensions > Apps Script 열기
 * 2. 이 코드 전체 붙여넣기 (기존 내용 교체)
 * 3. 상단 메뉴에서 initSheets 함수 선택 후 ▶ 실행 (최초 1회)
 *    → Holdings 시트, Config 시트, UserData 시트 자동 생성
 * 4. Deploy > New deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. 배포 URL 복사 → 앱의 설정 탭에 붙여넣기
 *
 * ── Config 시트 컬럼 구조 ───────────────────────────────────────
 * ticker | name | target_pct | div_per_share | div_months | currency
 * currency: KRW(기본) 또는 USD (달러 자산은 실시간 환율 자동 적용)
 *
 * ── Holdings 시트 컬럼 구조 ────────────────────────────────────
 * ticker | shares | avg_price
 *
 * ── UserData 시트 컬럼 구조 ────────────────────────────────────
 * key | value
 * ────────────────────────────────────────────────────────────────
 */

const HOLDINGS_SHEET = 'Holdings';
const CONFIG_SHEET   = 'Config';
const USERDATA_SHEET = 'UserData';

// ── 최초 시트 초기화 (1회만 실행) ───────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config 시트: ticker | name | target_pct | div_per_share | div_months | currency
  let cfg = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  cfg.clearContents();
  cfg.getRange('A1:F1').setValues([['ticker','name','target_pct','div_per_share','div_months','currency']]);
  cfg.getRange('A2:F6').setValues([
    ['0072R0', 'TIGER KRX금현물',             10, 0, '',                              'KRW'],
    ['379810', 'KODEX 미국나스닥100',          10, 0, '',                              'KRW'],
    ['379800', 'KODEX 미국S&P500',             20, 0, '',                              'KRW'],
    ['441640', 'KODEX 미국배당커버드콜액티브', 30, 0, '1,2,3,4,5,6,7,8,9,10,11,12',  'KRW'],
    ['458730', 'TIGER 미국배당다우존스',       30, 0, '1,2,3,4,5,6,7,8,9,10,11,12',  'KRW'],
  ]);
  cfg.setFrozenRows(1);

  // Holdings 시트: ticker | shares | avg_price
  let hld = ss.getSheetByName(HOLDINGS_SHEET) || ss.insertSheet(HOLDINGS_SHEET);
  hld.clearContents();
  hld.getRange('A1:C1').setValues([['ticker','shares','avg_price']]);
  hld.getRange('A2:C6').setValues([
    ['0072R0', 0, 0], ['379810', 0, 0], ['379800', 0, 0], ['441640', 0, 0], ['458730', 0, 0],
  ]);
  hld.setFrozenRows(1);

  // UserData 시트: key | value
  let ud = ss.getSheetByName(USERDATA_SHEET) || ss.insertSheet(USERDATA_SHEET);
  ud.clearContents();
  ud.getRange('A1:B1').setValues([['key','value']]);
  ud.getRange('A2:B2').setValues([['totalCashInvested', 0]]);
  ud.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ 초기화 완료!\n\nDeploy > New deployment > Web App으로 배포하세요.');
}

// ── GET 요청 핸들러 ──────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'getPortfolio';
    let data;

    if      (action === 'getPortfolio')    data = getPortfolio();
    else if (action === 'updateHoldings')  data = updateHoldings(parseParam(e.parameter.data));
    else if (action === 'setHoldings')     data = setHoldings(parseParam(e.parameter.data));
    else if (action === 'updateConfig')    data = updateConfig(parseParam(e.parameter.data));
    else if (action === 'getDividends')    data = getDividends(parseParam(e.parameter.data));
    else if (action === 'saveUserData')    data = saveUserData(parseParam(e.parameter.data));
    else throw new Error('Unknown action: ' + action);

    return respond({ success: true, data });
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

function parseParam(raw) {
  return JSON.parse(decodeURIComponent(raw || '[]'));
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 포트폴리오 조회 ──────────────────────────────────────────────
function getPortfolio() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cfgRows  = ss.getSheetByName(CONFIG_SHEET).getDataRange().getValues().slice(1);
  const hldSheet = ss.getSheetByName(HOLDINGS_SHEET);

  // Holdings 시트에 avg_price 컬럼 없으면 자동 추가
  ensureAvgPriceColumn(hldSheet);
  const hldRows = hldSheet.getDataRange().getValues().slice(1);

  // 보유수량 + 평균단가 맵
  const holdMap = {};
  const avgMap  = {};
  hldRows.forEach(r => {
    holdMap[String(r[0])] = Number(r[1]) || 0;
    avgMap[String(r[0])]  = Number(r[2]) || 0;
  });

  // USD 자산 존재 시 환율 조회
  const hasUSD = cfgRows.some(r => String(r[5]).toUpperCase() === 'USD');
  const usdkrw = hasUSD ? fetchUSDKRW() : null;

  // 실시간 가격 + 종목명 조회 (네이버)
  const naverData = {};
  cfgRows.forEach(r => {
    const ticker = String(r[0]);
    naverData[ticker] = fetchNaverData(ticker);
    Utilities.sleep(150);
  });

  // UserData 조회
  const userData = getUserDataObj();

  // 자산 목록 구성
  const assets = cfgRows.map(r => {
    const ticker        = String(r[0]);
    const cfgName       = String(r[1]);
    const target_pct    = Number(r[2]);
    const div_per_share = Number(r[3]) || 0;
    const divStr        = r[4] ? String(r[4]) : '';
    const div_months    = divStr
      ? divStr.split(',').map(s => parseInt(s.trim())).filter(n => n >= 1 && n <= 12)
      : [];
    const currency  = r[5] ? String(r[5]).toUpperCase() : 'KRW';
    const shares    = holdMap[ticker] || 0;
    const avg_price = avgMap[ticker]  || 0;

    // 네이버에서 가져온 이름 우선, 없으면 Config 시트 이름
    const name      = naverData[ticker]?.name || cfgName;
    const priceOrig = naverData[ticker]?.price || 0;
    const priceKRW  = currency === 'USD' && usdkrw ? priceOrig * usdkrw : priceOrig;
    const divPerShareKRW = currency === 'USD' && usdkrw ? div_per_share * usdkrw : div_per_share;

    return {
      ticker, name, target_pct, currency,
      div_per_share: divPerShareKRW,
      div_per_share_orig: div_per_share,
      div_months,
      shares,
      avg_price,
      price_orig: priceOrig,
      price: priceKRW,
      value: shares * priceKRW,
      current_pct: 0,
    };
  });

  const total = assets.reduce((s, a) => s + a.value, 0);
  assets.forEach(a => { a.current_pct = total > 0 ? a.value / total * 100 : 0; });

  return {
    assets,
    total_value: total,
    usdkrw: usdkrw,
    userData: userData,
    last_updated: new Date().toISOString(),
  };
}

// ── USD/KRW 환율 조회 (Naver Finance) ────────────────────────────
function fetchUSDKRW() {
  try {
    const url = 'https://polling.finance.naver.com/api/realtime?query=SERVICE_FOREX:USD';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return getFallbackUSDKRW();

    const json  = JSON.parse(res.getContentText());
    const datas = json?.result?.areas?.[0]?.datas;
    if (!datas?.length) return getFallbackUSDKRW();

    const item = datas[0];
    const raw  = item.nv || item.sv || item.basePrice || '0';
    const rate = parseFloat(String(raw).replace(/,/g, ''));
    return rate > 0 ? rate : getFallbackUSDKRW();
  } catch (e) {
    Logger.log(`[fetchUSDKRW] ${e}`);
    return getFallbackUSDKRW();
  }
}

function getFallbackUSDKRW() {
  try {
    const url = 'https://m.stock.naver.com/api/forex/FX_USDKRW';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const rate = parseFloat(json?.closePrice?.replace(/,/g,'') || '0');
    return rate > 500 ? rate : 1350;
  } catch (e) {
    return 1350;
  }
}

// ── Naver Finance 실시간 가격 + 종목명 조회 ─────────────────────
function fetchNaverData(ticker) {
  try {
    const url = `https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:${ticker}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { price: 0, name: null };

    const json  = JSON.parse(res.getContentText());
    const datas = json?.result?.areas?.[0]?.datas;
    if (!datas?.length) return { price: 0, name: null };

    const item  = datas[0];
    const raw   = item.nv || item.sv || item.pv || item.closePrice || '0';
    const price = parseFloat(String(raw).replace(/,/g, '')) || 0;
    // nm = 종목명 (네이버 API 표준 필드)
    const name  = item.nm || item.itemName || item.stockName || null;
    return { price, name };
  } catch (e) {
    Logger.log(`[fetchNaverData] ${ticker}: ${e}`);
    return { price: 0, name: null };
  }
}

// 하위 호환성용 래퍼
function fetchNaverPrice(ticker) {
  return fetchNaverData(ticker).price;
}

// ── Holdings 시트 avg_price 컬럼 자동 보완 ──────────────────────
function ensureAvgPriceColumn(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return;
  if (rows[0][2] !== 'avg_price') {
    sheet.getRange(1, 3).setValue('avg_price');
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][2]) sheet.getRange(i + 1, 3).setValue(0);
    }
  }
}

// ── 보유수량 업데이트: [{ticker, addShares}] ─────────────────────
function updateHoldings(updates) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOLDINGS_SHEET);
  ensureAvgPriceColumn(sheet);
  const rows  = sheet.getDataRange().getValues();

  const map = {};
  updates.forEach(u => { map[u.ticker] = u.addShares; });

  for (let i = 1; i < rows.length; i++) {
    const ticker = String(rows[i][0]);
    if (map[ticker] !== undefined) {
      const current = Number(rows[i][1]) || 0;
      sheet.getRange(i + 1, 2).setValue(current + map[ticker]);
    }
  }
  return { updated: updates.length };
}

// ── 보유수량 + 평균단가 절대값 설정: [{ticker, shares, avg_price}] ──
function setHoldings(updates) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOLDINGS_SHEET);
  ensureAvgPriceColumn(sheet);
  const rows  = sheet.getDataRange().getValues();

  const map = {};
  updates.forEach(u => { map[u.ticker] = u; });

  for (let i = 1; i < rows.length; i++) {
    const ticker = String(rows[i][0]);
    const u = map[ticker];
    if (!u) continue;
    sheet.getRange(i + 1, 2).setValue(Number(u.shares)    || 0);
    sheet.getRange(i + 1, 3).setValue(Number(u.avg_price) || 0);
  }
  return { updated: updates.length };
}

// ── UserData 조회 ────────────────────────────────────────────────
function getUserDataObj() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERDATA_SHEET);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues().slice(1);
  const obj  = {};
  rows.forEach(r => { if (r[0]) obj[String(r[0])] = r[1]; });
  return obj;
}

// ── UserData 저장: { totalCashInvested: number, ... } ────────────
function saveUserData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERDATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USERDATA_SHEET);
    sheet.getRange('A1:B1').setValues([['key','value']]);
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();
  for (const [key, value] of Object.entries(data)) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        rows[i][1] = value;
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, value]);
  }
  return { saved: Object.keys(data).length };
}

// ── Yahoo Finance 배당 기록 프록시 ───────────────────────────────
function getDividends(tickers) {
  const results = {};
  tickers.forEach(function(ticker) {
    const isKR    = /^\d/.test(ticker);
    const yTicker = isKR ? ticker + '.KS' : ticker.toUpperCase();
    try {
      const url = 'https://query1.finance.yahoo.com/v8/finance/chart/'
        + encodeURIComponent(yTicker) + '?events=div&range=2y&interval=1mo';
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return;
      const data = JSON.parse(res.getContentText());
      const divs = data.chart && data.chart.result && data.chart.result[0]
        && data.chart.result[0].events && data.chart.result[0].events.dividends;
      if (!divs) return;
      results[ticker] = Object.values(divs).map(function(d) {
        const dt   = new Date(d.date * 1000);
        const ym   = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
        const date = ym + '-' + String(dt.getDate()).padStart(2, '0');
        return { ym: ym, date: date, amount: d.amount, source: 'auto' };
      }).sort(function(a, b) { return a.date.localeCompare(b.date); });
    } catch(e) {
      Logger.log('[getDividends] ' + ticker + ': ' + e);
    }
  });
  return results;
}

// ── 배당/통화/비중 설정 업데이트 ─────────────────────────────────
function updateConfig(updates) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  const rows  = sheet.getDataRange().getValues();

  const map = {};
  updates.forEach(u => { map[u.ticker] = u; });

  for (let i = 1; i < rows.length; i++) {
    const ticker = String(rows[i][0]);
    const u = map[ticker];
    if (!u) continue;
    if (u.div_per_share !== undefined) sheet.getRange(i + 1, 4).setValue(u.div_per_share);
    if (u.div_months    !== undefined) sheet.getRange(i + 1, 5).setValue(u.div_months.join(','));
    if (u.currency      !== undefined) sheet.getRange(i + 1, 6).setValue(u.currency);
    if (u.target_pct    !== undefined) sheet.getRange(i + 1, 3).setValue(u.target_pct);
  }
  return { updated: updates.length };
}
