const RAW_HEADERS = [
  '記錄時間',
  '日期',
  '區域',
  '門市',
  'GU姓名',
  '膚質',
  '主困擾',
  '其他困擾',
  '眼周狀態',
  '生活情境',
  'step1',
  'step2',
  'step3',
  'step4',
  'product_1',
  'product_2',
  'product_3',
  'product_4',
  'product_summary',
  'source'
];

const PRODUCT_HEADERS = [
  '記錄時間',
  '日期',
  '區域',
  '門市',
  'GU姓名',
  '主困擾',
  '商品順序',
  '推薦商品'
];

const LIFESTYLE_HEADERS = [
  '記錄時間',
  '日期',
  '區域',
  '門市',
  'GU姓名',
  '主困擾',
  '生活情境'
];

function ensureSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function splitList(value) {
  if (!value) return [];
  return String(value)
    .split('、')
    .map(function(item) { return item.trim(); })
    .filter(Boolean);
}

function getRawRowFromParams(e) {
  return [
    new Date().toLocaleString('zh-TW'),
    e.parameter.date || '',
    e.parameter.region || '',
    e.parameter.store || '',
    e.parameter.gu || e.parameter.ba || '',
    e.parameter.q1_skin_type || '',
    e.parameter.q2_main_concern || '',
    e.parameter.q3_other_concerns || '',
    e.parameter.q4_eye || '',
    e.parameter.q5_lifestyle || '',
    e.parameter.step1 || '',
    e.parameter.step2 || '',
    e.parameter.step3 || '',
    e.parameter.step4 || '',
    e.parameter.product_1 || '',
    e.parameter.product_2 || '',
    e.parameter.product_3 || '',
    e.parameter.product_4 || '',
    e.parameter.product_summary || '',
    e.parameter.source || ''
  ];
}

function appendDerivedRows(rawRow) {
  var productSheet = ensureSheet('products_long', PRODUCT_HEADERS);
  var lifestyleSheet = ensureSheet('lifestyles_long', LIFESTYLE_HEADERS);

  var base = {
    timestamp: rawRow[0],
    date: rawRow[1],
    region: rawRow[2],
    store: rawRow[3],
    gu: rawRow[4],
    mainConcern: rawRow[6]
  };

  var products = rawRow.slice(14, 18).filter(Boolean);
  if (products.length) {
    var productRows = products.map(function(product, index) {
      return [
        base.timestamp,
        base.date,
        base.region,
        base.store,
        base.gu,
        base.mainConcern,
        index + 1,
        product
      ];
    });
    productSheet.getRange(productSheet.getLastRow() + 1, 1, productRows.length, PRODUCT_HEADERS.length).setValues(productRows);
  }

  var lifestyles = splitList(rawRow[9]);
  if (lifestyles.length) {
    var lifestyleRows = lifestyles.map(function(item) {
      return [
        base.timestamp,
        base.date,
        base.region,
        base.store,
        base.gu,
        base.mainConcern,
        item
      ];
    });
    lifestyleSheet.getRange(lifestyleSheet.getLastRow() + 1, 1, lifestyleRows.length, LIFESTYLE_HEADERS.length).setValues(lifestyleRows);
  }
}

function doGet(e) {
  var rawSheet = ensureSheet('raw_data', RAW_HEADERS);
  var rawRow = getRawRowFromParams(e);
  rawSheet.getRange(rawSheet.getLastRow() + 1, 1, 1, RAW_HEADERS.length).setValues([rawRow]);
  appendDerivedRows(rawRow);

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function rebuildDerivedSheets() {
  var rawSheet = ensureSheet('raw_data', RAW_HEADERS);
  var productSheet = ensureSheet('products_long', PRODUCT_HEADERS);
  var lifestyleSheet = ensureSheet('lifestyles_long', LIFESTYLE_HEADERS);
  productSheet.clearContents();
  lifestyleSheet.clearContents();
  productSheet.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
  lifestyleSheet.getRange(1, 1, 1, LIFESTYLE_HEADERS.length).setValues([LIFESTYLE_HEADERS]);

  var lastRow = rawSheet.getLastRow();
  if (lastRow <= 1) return;

  var values = rawSheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();
  values.forEach(function(row) {
    appendDerivedRows(row);
  });
}

function setupSkinCheckWorkbook() {
  ensureSheet('raw_data', RAW_HEADERS);
  ensureSheet('products_long', PRODUCT_HEADERS);
  ensureSheet('lifestyles_long', LIFESTYLE_HEADERS);
  buildDashboardSheet();
  buildStoreAnalysisSheet();
  buildProductAnalysisSheet();
  buildLifestyleAnalysisSheet();
  buildRegionAnalysisSheet();
  rebuildDerivedSheets();
}

function styleTitle(sheet, cellA1, title) {
  sheet.getRange(cellA1).setValue(title).setFontSize(18).setFontWeight('bold');
}

function buildDashboardSheet() {
  var sheet = ensureSheet('dashboard');
  sheet.clear();
  styleTitle(sheet, 'A1', 'Skin Check Dashboard');
  sheet.getRange('A3').setValue('總檢測數');
  sheet.getRange('B3').setFormula('=COUNTA(raw_data!A2:A)');
  sheet.getRange('D3').setValue('本週檢測數');
  sheet.getRange('E3').setFormula('=COUNTIFS(raw_data!B2:B,\">=\"&TODAY()-WEEKDAY(TODAY(),2)+1,raw_data!B2:B,\"<\"&TODAY()-WEEKDAY(TODAY(),2)+8)');
  sheet.getRange('G3').setValue('本月檢測數');
  sheet.getRange('H3').setFormula('=COUNTIFS(raw_data!B2:B,\">=\"&EOMONTH(TODAY(),-1)+1,raw_data!B2:B,\"<=\"&EOMONTH(TODAY(),0))');
  sheet.getRange('J3').setValue('使用門市數');
  sheet.getRange('K3').setFormula('=COUNTUNIQUE(FILTER(raw_data!D2:D,raw_data!D2:D<>\"\"))');
  sheet.getRange('M3').setValue('使用 GU 數');
  sheet.getRange('N3').setFormula('=COUNTUNIQUE(FILTER(raw_data!E2:E,raw_data!E2:E<>\"\"))');

  sheet.getRange('A7').setFormula("=QUERY(raw_data!C2:C,\"select C,count(C) where C<>'' group by C order by count(C) desc label C '區域', count(C) '檢測數'\",0)");
  sheet.getRange('D7').setFormula("=QUERY(raw_data!D2:D,\"select D,count(D) where D<>'' group by D order by count(D) desc label D '門市', count(D) '檢測數'\",0)");
  sheet.getRange('G7').setFormula("=QUERY(raw_data!G2:G,\"select G,count(G) where G<>'' group by G order by count(G) desc label G '主困擾', count(G) '次數'\",0)");
  sheet.getRange('J7').setFormula("=QUERY(lifestyles_long!G2:G,\"select G,count(G) where G<>'' group by G order by count(G) desc label G '生活情境', count(G) '次數'\",0)");
  sheet.getRange('M7').setFormula("=QUERY(products_long!H2:H,\"select H,count(H) where H<>'' group by H order by count(H) desc label H '推薦商品', count(H) '次數'\",0)");
  sheet.autoResizeColumns(1, 14);
}

function buildStoreAnalysisSheet() {
  var sheet = ensureSheet('store_analysis');
  sheet.clear();
  styleTitle(sheet, 'A1', 'Store Analysis');
  sheet.getRange('A3').setFormula("=QUERY(raw_data!D2:D,\"select D,count(D) where D<>'' group by D order by count(D) desc label D '門市', count(D) '檢測數'\",0)");
  sheet.getRange('D3').setFormula("=QUERY(raw_data!D2:G,\"select D,count(D) where D<>'' and G<>'' group by D pivot G label count(D) '主困擾次數'\",0)");
  sheet.getRange('D20').setFormula("=QUERY(lifestyles_long!D2:G,\"select D,count(D) where D<>'' and G<>'' group by D pivot G label count(D) '生活情境次數'\",0)");
  sheet.getRange('D37').setFormula("=QUERY(products_long!D2:H,\"select D,count(D) where D<>'' and H<>'' group by D pivot H label count(D) '推薦商品次數'\",0)");
  sheet.autoResizeColumns(1, 20);
}

function buildProductAnalysisSheet() {
  var sheet = ensureSheet('product_analysis');
  sheet.clear();
  styleTitle(sheet, 'A1', 'Product Analysis');
  sheet.getRange('A3').setFormula("=QUERY(products_long!H2:H,\"select H,count(H) where H<>'' group by H order by count(H) desc label H '推薦商品', count(H) '次數'\",0)");
  sheet.getRange('D3').setFormula("=QUERY(products_long!F2:H,\"select H,count(H) where H<>'' group by H pivot F label count(H) '主困擾次數'\",0)");
  sheet.getRange('D20').setFormula("=QUERY(products_long!C2:H,\"select H,count(H) where H<>'' group by H pivot C label count(H) '區域次數'\",0)");
  sheet.autoResizeColumns(1, 20);
}

function buildLifestyleAnalysisSheet() {
  var sheet = ensureSheet('lifestyle_analysis');
  sheet.clear();
  styleTitle(sheet, 'A1', 'Lifestyle Analysis');
  sheet.getRange('A3').setFormula("=QUERY(lifestyles_long!G2:G,\"select G,count(G) where G<>'' group by G order by count(G) desc label G '生活情境', count(G) '次數'\",0)");
  sheet.getRange('D3').setFormula("=QUERY(lifestyles_long!F2:G,\"select G,count(G) where G<>'' group by G pivot F label count(G) '主困擾次數'\",0)");
  sheet.getRange('D20').setFormula("=QUERY(lifestyles_long!C2:G,\"select G,count(G) where G<>'' group by G pivot C label count(G) '區域次數'\",0)");
  sheet.autoResizeColumns(1, 20);
}

function buildRegionAnalysisSheet() {
  var sheet = ensureSheet('region_analysis');
  sheet.clear();
  styleTitle(sheet, 'A1', 'Region Analysis');
  sheet.getRange('A3').setFormula("=QUERY(raw_data!C2:C,\"select C,count(C) where C<>'' group by C order by count(C) desc label C '區域', count(C) '檢測數'\",0)");
  sheet.getRange('D3').setFormula("=QUERY(raw_data!C2:G,\"select C,count(C) where C<>'' and G<>'' group by C pivot G label count(C) '主困擾次數'\",0)");
  sheet.getRange('D20').setFormula("=QUERY(lifestyles_long!C2:G,\"select C,count(C) where C<>'' and G<>'' group by C pivot G label count(C) '生活情境次數'\",0)");
  sheet.getRange('D37').setFormula("=QUERY(products_long!C2:H,\"select C,count(C) where C<>'' and H<>'' group by C pivot H label count(C) '推薦商品次數'\",0)");
  sheet.autoResizeColumns(1, 20);
}
