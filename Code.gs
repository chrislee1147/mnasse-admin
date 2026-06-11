// ============================================================
// 므낫세 어드민 — Google Apps Script API
// 배포: 확장 프로그램 > Apps Script > 웹앱으로 배포
//       실행: 나, 액세스: 모든 사용자
// ============================================================

const SHEETS = {
  PURCHASE: '매입_파이프라인',
  SALES: '매출_파이프라인',
  INVENTORY: '재고',
  CONTENT: '콘텐츠_트래커',
  WAITLIST: '대기리스트',
  KPI: 'KPI_일별',
  APPROVAL: '포스팅_승인큐'
};

const PURCHASE_STAGES = ['문의접수', '견적발송', '협의중', '계약완료', '입고완료', '검수완료', '재고등록', 'DROP'];
const SALES_STAGES = ['문의접수', '상담완료', '현장확인', '계약완료', '장착완료', '보관연결', 'DROP'];
const CHANNELS = ['인스타DM', '카카오DM', '카카오톡', '문자', '번개장터', '당근마켓', '네이버카페', '보관고객', '딜러소개', '지인소개', '기타'];

// ============================================================
// HTTP 핸들러
// ============================================================

function doGet(e) {
  const result = handleGet(e.parameter);
  return jsonResponse(result);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const result = handlePost(body);
  return jsonResponse(result);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGet(params) {
  try {
    switch (params.action) {
      case 'dashboard':  return getDashboard();
      case 'pipeline':   return getPipeline(params.type);
      case 'inventory':  return getInventory();
      case 'content':    return getContent();
      case 'waitlist':    return getWaitlist();
      case 'kpi':         return getKPI();
      case 'config':      return getConfig();
      case 'getApproval': return getApprovalQueue();
      default:            return { error: '알 수 없는 액션: ' + params.action };
    }
  } catch (e) {
    return { error: e.message };
  }
}

function handlePost(body) {
  try {
    switch (body.action) {
      case 'addLead':         return addLead(body.type, body.data);
      case 'updateStatus':    return updateStatus(body.sheet, body.id, body.status);
      case 'updateField':     return updateField(body.sheet, body.id, body.field, body.value);
      case 'addInventory':    return addInventory(body.data);
      case 'updateInventory': return updateField(SHEETS.INVENTORY, body.id, body.field, body.value, '재고ID');
      case 'addContent':      return addContent(body.data);
      case 'addWaitlist':     return addWaitlist(body.data);
      case 'updateWaitlist':  return updateField(SHEETS.WAITLIST, body.id, body.field, body.value, '이름');
      case 'addKPI':          return addKPI(body.data);
      case 'addApproval':     return addApproval(body.data);
      case 'approveContent':  return setApprovalStatus(body.stem, 'approved');
      case 'rejectContent':   return setApprovalStatus(body.stem, 'rejected');
      case 'markPosted':      return markPosted(body.stem, body.postId);
      case 'markRejectedDone':return setApprovalStatus(body.stem, 'rejected_done');
      default:                return { error: '알 수 없는 액션: ' + body.action };
    }
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// 유틸
// ============================================================

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => {
      obj[h] = row[j] instanceof Date ? row[j].toISOString() : row[j];
    });
    return obj;
  });
}

function generateId(prefix) {
  return prefix + '_' + Date.now();
}

function fmtDate(d) {
  return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function findRow(sheet, id, idCol) {
  idCol = idCol || '리드ID';
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf(idCol);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) return i + 1;
  }
  return -1;
}

function getColIndex(sheet, fieldName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(fieldName) + 1;
}

// ============================================================
// 읽기
// ============================================================

function getDashboard() {
  const purchase = sheetToObjects(getSheet(SHEETS.PURCHASE));
  const sales    = sheetToObjects(getSheet(SHEETS.SALES));
  const inventory = getInventory();
  const waitlist = sheetToObjects(getSheet(SHEETS.WAITLIST));
  const kpi      = sheetToObjects(getSheet(SHEETS.KPI));

  const weekAgo = new Date(Date.now() - 7 * 86400000);

  const activePurchase  = purchase.filter(r => r['파이프상태'] !== 'DROP');
  const activeSales     = sales.filter(r => r['파이프상태'] !== 'DROP');
  const activeInventory = inventory.filter(r => r['판매상태'] !== '판매완료');
  const activeWaitlist  = waitlist.filter(r => r['알림상태'] === '대기중');

  // 경고
  const alerts = [];
  activeInventory.forEach(item => {
    const days = parseInt(item['보관일수']) || 0;
    if (days >= 90) alerts.push({ type: 'danger',  msg: `[재고] ${item['모델']} — ${days}일 보관. 즉시 가격조정 필요`, id: item['재고ID'] });
    else if (days >= 60) alerts.push({ type: 'warning', msg: `[재고] ${item['모델']} — ${days}일 보관. 검토 필요`, id: item['재고ID'] });
  });
  const pendingQuote = activePurchase.filter(r => r['파이프상태'] === '문의접수');
  if (pendingQuote.length > 0) alerts.push({ type: 'info', msg: `매입 견적 미발송 ${pendingQuote.length}건` });
  const pendingConsult = activeSales.filter(r => r['파이프상태'] === '문의접수');
  if (pendingConsult.length > 0) alerts.push({ type: 'info', msg: `매출 상담 대기 ${pendingConsult.length}건` });

  // 채널별 집계
  const channelMap = {};
  [...activePurchase, ...activeSales].forEach(r => {
    const ch = r['유입채널'] || '기타';
    channelMap[ch] = (channelMap[ch] || 0) + 1;
  });

  // 퍼널 전환율
  const purchaseFunnel = calcFunnel(activePurchase, PURCHASE_STAGES.slice(0, -1), '파이프상태');
  const salesFunnel    = calcFunnel(activeSales,    SALES_STAGES.slice(0, -1),    '파이프상태');

  return {
    counts: {
      purchase: activePurchase.length,
      sales:    activeSales.length,
      waitlist: activeWaitlist.length,
      inventory: activeInventory.length,
      purchaseWeek: activePurchase.filter(r => new Date(r['날짜']) >= weekAgo).length,
      salesWeek:    activeSales.filter(r => new Date(r['날짜']) >= weekAgo).length,
      avgInventoryDays: activeInventory.length
        ? Math.round(activeInventory.reduce((s, i) => s + (parseInt(i['보관일수']) || 0), 0) / activeInventory.length)
        : 0
    },
    alerts,
    channelMap,
    purchaseFunnel,
    salesFunnel,
    recentKPI: kpi.slice(-14)
  };
}

function calcFunnel(leads, stages, statusField) {
  return stages.map(stage => ({
    stage,
    count: leads.filter(r => r[statusField] === stage).length
  }));
}

function getPipeline(type) {
  const sheetName = type === '매입' ? SHEETS.PURCHASE : SHEETS.SALES;
  return sheetToObjects(getSheet(sheetName));
}

function getInventory() {
  const items = sheetToObjects(getSheet(SHEETS.INVENTORY));
  const today = Date.now();
  return items.map(item => {
    if (item['매입일'] && item['판매상태'] !== '판매완료') {
      const d = new Date(item['매입일']);
      item['보관일수'] = Math.floor((today - d.getTime()) / 86400000);
    }
    return item;
  });
}

function getContent()  { return sheetToObjects(getSheet(SHEETS.CONTENT)); }
function getWaitlist() { return sheetToObjects(getSheet(SHEETS.WAITLIST)); }
function getKPI()      { return sheetToObjects(getSheet(SHEETS.KPI)); }

function getConfig() {
  return {
    purchaseStages: PURCHASE_STAGES,
    salesStages:    SALES_STAGES,
    channels:       CHANNELS
  };
}

// ============================================================
// 쓰기
// ============================================================

function addLead(type, data) {
  const isPurchase = type === '매입';
  const sheet = getSheet(isPurchase ? SHEETS.PURCHASE : SHEETS.SALES);
  const today = fmtDate();
  const id = generateId(isPurchase ? 'P' : 'S');

  const row = isPurchase
    ? [today, id, data.name, data.contact, data.channel,
       data.model || '', data.usage || '', data.grade || '미확인',
       data.quote || '', '문의접수', data.purchaseType || '즉시현금', '', data.memo || '', today]
    : [today, id, data.name, data.contact, data.channel,
       data.model || '', data.budget || '', data.carType || '',
       '문의접수', '', data.memo || '', today];

  sheet.appendRow(row);
  return { success: true, id };
}

function updateStatus(sheetName, id, status) {
  const sheet = getSheet(sheetName);
  const rowNum = findRow(sheet, id);
  if (rowNum < 0) return { error: 'ID 없음' };

  const statusCol   = getColIndex(sheet, '파이프상태');
  const modifiedCol = getColIndex(sheet, '최종수정일');
  if (statusCol > 0)   sheet.getRange(rowNum, statusCol).setValue(status);
  if (modifiedCol > 0) sheet.getRange(rowNum, modifiedCol).setValue(fmtDate());
  return { success: true };
}

function updateField(sheetName, id, field, value, idColName) {
  const sheet = getSheet(sheetName);
  const rowNum = findRow(sheet, id, idColName || '리드ID');
  if (rowNum < 0) return { error: 'ID 없음' };

  const col = getColIndex(sheet, field);
  if (col > 0) sheet.getRange(rowNum, col).setValue(value);

  const modifiedCol = getColIndex(sheet, '최종수정일');
  if (modifiedCol > 0) sheet.getRange(rowNum, modifiedCol).setValue(fmtDate());
  return { success: true };
}

function addInventory(data) {
  const sheet = getSheet(SHEETS.INVENTORY);
  const id = generateId('I');
  sheet.appendRow([
    id, data.model, data.grade,
    data.purchasePrice || '', data.purchaseDate || new Date(),
    data.repairDetail || '', data.repairCost || 0,
    data.salePrice || '', data.status || '검수중',
    0, data.photoUrl || '', data.listingChannel || '', data.memo || ''
  ]);
  return { success: true, id };
}

function addContent(data) {
  getSheet(SHEETS.CONTENT).appendRow([
    data.date || new Date(), data.channel, data.type,
    data.title, data.url || '', data.status || '예정',
    0, 0, data.memo || ''
  ]);
  return { success: true };
}

function addWaitlist(data) {
  getSheet(SHEETS.WAITLIST).appendRow([
    new Date(), data.name, data.contact, data.model, data.budget || '', '대기중'
  ]);
  return { success: true };
}

function addKPI(data) {
  getSheet(SHEETS.KPI).appendRow([
    data.date || new Date(),
    data.purchaseLeads || 0, data.salesLeads || 0,
    data.newWaitlist || 0, data.instaFollowers || 0,
    data.instaDM || 0, data.blogVisits || 0,
    data.cafeViews || 0, data.purchaseConvert || 0, data.salesConvert || 0
  ]);
  return { success: true };
}

// ============================================================
// 포스팅 승인 큐
// ============================================================

function getApprovalQueue() {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return [];
  const items = sheetToObjects(sheet);
  // ready 항목만 반환 (thumbnail은 용량 절약 위해 포함)
  return items.filter(r => r['status'] === 'ready');
}

function addApproval(data) {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음 — setupApprovalSheet() 실행 필요' };
  sheet.appendRow([
    data.date    || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data.stem        || '',
    data.media_type  || '',
    data.folder_type || '',
    data.file        || '',
    data.caption     || '',
    data.thumbnail   || '',
    'ready',
    '',   // post_id
    ''    // processed_at
  ]);
  return { success: true };
}

function setApprovalStatus(stem, status) {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음' };
  const rowNum = _findApprovalRow(sheet, stem);
  if (rowNum < 0) return { error: 'stem 없음: ' + stem };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol      = headers.indexOf('status') + 1;
  const processedAtCol = headers.indexOf('processed_at') + 1;

  if (statusCol > 0) sheet.getRange(rowNum, statusCol).setValue(status);
  if (processedAtCol > 0) {
    sheet.getRange(rowNum, processedAtCol)
         .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  }
  return { success: true };
}

function markPosted(stem, postId) {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음' };
  const rowNum = _findApprovalRow(sheet, stem);
  if (rowNum < 0) return { error: 'stem 없음: ' + stem };

  const headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol  = headers.indexOf('status') + 1;
  const postIdCol  = headers.indexOf('post_id') + 1;
  const procCol    = headers.indexOf('processed_at') + 1;

  if (statusCol > 0) sheet.getRange(rowNum, statusCol).setValue('posted');
  if (postIdCol > 0) sheet.getRange(rowNum, postIdCol).setValue(postId || '');
  if (procCol   > 0) sheet.getRange(rowNum, procCol).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  return { success: true };
}

function _findApprovalRow(sheet, stem) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const stemIdx = headers.indexOf('stem');
  if (stemIdx < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][stemIdx]) === String(stem)) return i + 1;
  }
  return -1;
}

// ============================================================
// 최초 세팅 — 스크립트 에디터에서 직접 실행
// ============================================================

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const configs = [
    {
      name: SHEETS.PURCHASE,
      headers: ['날짜','리드ID','이름','연락처','유입채널','RTT모델','사용기간','상태등급','견적가','파이프상태','매입유형','DROP사유','메모','최종수정일'],
      color: '#E85D5D'
    },
    {
      name: SHEETS.SALES,
      headers: ['날짜','리드ID','이름','연락처','유입채널','관심모델','예산대','보유차종','파이프상태','DROP사유','메모','최종수정일'],
      color: '#2B9E9E'
    },
    {
      name: SHEETS.INVENTORY,
      headers: ['재고ID','모델','등급','매입가','매입일','수리내역','수리비','판매가','판매상태','보관일수','사진URL','리스팅채널','메모'],
      color: '#4A7FA5'
    },
    {
      name: SHEETS.CONTENT,
      headers: ['날짜','채널','유형','제목','URL','상태','조회수','유입문의수','메모'],
      color: '#5A9E6F'
    },
    {
      name: SHEETS.WAITLIST,
      headers: ['날짜','이름','연락처','원하는모델','예산','알림상태'],
      color: '#D4A017'
    },
    {
      name: SHEETS.KPI,
      headers: ['날짜','매입문의','매출문의','대기리스트신규','인스타팔로워','인스타DM수','블로그방문','카페조회','매입전환','매출전환'],
      color: '#7B5EA7'
    },
    {
      name: SHEETS.APPROVAL,
      headers: ['날짜','stem','media_type','folder_type','file','caption','thumbnail','status','post_id','processed_at'],
      color: '#8B5CF6'
    }
  ];

  configs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) sheet = ss.insertSheet(cfg.name);
    const hRange = sheet.getRange(1, 1, 1, cfg.headers.length);
    hRange.setValues([cfg.headers])
          .setBackground(cfg.color)
          .setFontWeight('bold')
          .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, cfg.headers.length);
  });

  // 기본 시트 제거
  ['Sheet1', '시트1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > configs.length) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  Logger.log('✅ 세팅 완료');
  Logger.log('스프레드시트 ID: ' + ss.getId());
  Logger.log('다음: 배포 > 웹앱으로 배포 > URL 복사 > 어드민 index.html의 SCRIPT_URL에 붙여넣기');
}
