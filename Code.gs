// ============================================================
// 므낫세 어드민 — Google Apps Script API v2
// ============================================================

const SHEETS = {
  PIPELINE: '파이프라인',
  INVENTORY: '재고',
  CONTENT:   '콘텐츠_트래커',
  WAITLIST:  '대기리스트',
  KPI:       'KPI_일별',
  APPROVAL:  '포스팅_승인큐',
};

const PIPELINE_STAGES = ['DM유입','현장점검','흥정','매입완료','콘텐츠제작','판매중','판매완료'];
const CHANNELS = ['인스타DM','카카오DM','카카오톡','문자','번개장터','당근마켓','네이버카페','보관고객','딜러소개','지인소개','기타'];

// 모델코드 매핑 (긴 이름 먼저 — 부분매칭 순서 중요)
const MODEL_CODE_MAP = [
  ['스카이캠프 3.0 미니', 'scm3'],
  ['스카이캠프 미니 2.0', 'scm2'],
  ['스카이캠프 슬림',     'scs'],
  ['스카이캠프 3.0',      'sc3'],
  ['BDV',                 'bdv'],
  ['X-Cover',             'xcv'],
];

// 정산 항목 (추후 항목 추가 가능 — condition: always | restructured)
const SETTLEMENT_ITEMS = [
  { name: '기본(탈착·점검·보관)', amount: 20, condition: 'always' },
  { name: '구조변경 대행(턴키)',   amount: 30, condition: 'restructured' },
];

// ============================================================
// HTTP 핸들러
// ============================================================

function doGet(e) {
  return jsonResponse(handleGet(e.parameter));
}

function doPost(e) {
  return jsonResponse(handlePost(JSON.parse(e.postData.contents)));
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGet(p) {
  try {
    switch (p.action) {
      case 'dashboard':           return getDashboard();
      case 'getPipelineAll':      return getPipelineAll();
      case 'inventory':           return getInventory();
      case 'content':             return getContent();
      case 'waitlist':            return getWaitlist();
      case 'kpi':                 return getKPI();
      case 'config':              return getConfig();
      case 'getApproval':         return getApprovalQueue();
      case 'getPendingActions':   return getPendingActions();
      case 'cleanupDuplicates':   return cleanupDuplicates();
      case 'setupApprovalSheet':  return setupApprovalSheet();
      case 'getProcessTrigger':   return getProcessTrigger();
      case 'clearProcessTrigger': return clearProcessTrigger();
      case 'getScheduleSettings': return getScheduleSettings();
      case 'setupSheets':         return setupSheets();
      default: return { error: '알 수 없는 액션: ' + p.action };
    }
  } catch (e) { return { error: e.message }; }
}

function handlePost(body) {
  try {
    switch (body.action) {
      case 'addPipelineItem':      return addPipelineItem(body.data);
      case 'updatePipelineStage':  return updatePipelineStage(body.id, body.stage);
      case 'updatePipelineField':  return updatePipelineField(body.id, body.field, body.value);
      case 'addInventory':         return addInventory(body.data);
      case 'updateInventory':      return updatePipelineField(body.id, body.field, body.value);
      case 'addContent':           return addContent(body.data);
      case 'addWaitlist':          return addWaitlist(body.data);
      case 'matchWaitlist':        return matchWaitlist(body.waitlistId, body.mgmtNo);
      case 'addKPI':               return addKPI(body.data);
      case 'addApproval':          return addApproval(body.data);
      case 'approveContent':       return approveContent(body.stem, body.editedCaption);
      case 'rejectContent':        return setApprovalStatus(body.stem, 'rejected');
      case 'markPosted':           return markPosted(body.stem, body.postId);
      case 'markRejectedDone':     return setApprovalStatus(body.stem, 'rejected_done');
      case 'setProcessTrigger':    return setProcessTrigger(body.folder);
      case 'saveScheduleSettings': return saveScheduleSettings(body.data);
      default: return { error: '알 수 없는 액션: ' + body.action };
    }
  } catch (e) { return { error: e.message }; }
}

// ============================================================
// 유틸
// ============================================================

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
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

function fmtDate(d) {
  return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function findRow(sheet, id, idCol) {
  idCol = idCol || '관리번호';
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf(idCol);
  if (idIdx < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) return i + 1;
  }
  return -1;
}

function getColIndex(sheet, fieldName) {
  if (!sheet) return -1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(fieldName) + 1;
}

// ============================================================
// 관리번호 자동 채번  YYMMDD-{modelCode}-{0001}
// ============================================================

function generateMgmtNo(model) {
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyMMdd');

  const modelNorm = (model || '').replace(/\s/g, '').toLowerCase();
  let modelCode = 'etc';
  for (const [key, code] of MODEL_CODE_MAP) {
    if (modelNorm.includes(key.replace(/\s/g, '').toLowerCase())) {
      modelCode = code;
      break;
    }
  }

  const sheet = getSheet(SHEETS.PIPELINE);
  let lastSeq = 0;
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const mgmtIdx = headers.indexOf('관리번호');
    if (mgmtIdx >= 0) {
      const prefix = dateStr + '-' + modelCode + '-';
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][mgmtIdx]);
        if (id.startsWith(prefix)) {
          const seq = parseInt(id.slice(prefix.length)) || 0;
          if (seq > lastSeq) lastSeq = seq;
        }
      }
    }
  }

  return dateStr + '-' + modelCode + '-' + String(lastSeq + 1).padStart(4, '0');
}

// ============================================================
// 정산 계산
// ============================================================

function calcSettlement(restructured) {
  let total = 0;
  SETTLEMENT_ITEMS.forEach(function(item) {
    if (item.condition === 'always') total += item.amount;
    else if (item.condition === 'restructured' && restructured) total += item.amount;
  });
  return total;
}

function calcProfit(buyPrice, salePrice, restructured) {
  const buy  = parseFloat(buyPrice);
  const sell = parseFloat(salePrice);
  if (isNaN(buy) || isNaN(sell)) return '';
  return sell - buy - calcSettlement(restructured);
}

// ============================================================
// 대시보드
// ============================================================

function getDashboard() {
  const pipeline  = getPipelineAll();
  const inventory = getInventory();
  const waitlist  = sheetToObjects(getSheet(SHEETS.WAITLIST));
  const kpi       = sheetToObjects(getSheet(SHEETS.KPI));

  const weekAgo = new Date(Date.now() - 7 * 86400000);
  const buying  = pipeline.filter(function(r){ return ['DM유입','현장점검','흥정'].indexOf(r['단계']) >= 0; });
  const selling = pipeline.filter(function(r){ return ['콘텐츠제작','판매중'].indexOf(r['단계']) >= 0; });
  const done    = pipeline.filter(function(r){ return r['단계'] === '판매완료'; });
  const activeWL = waitlist.filter(function(r){ return r['알림상태'] === '대기중'; });
  const activeInv = inventory.filter(function(r){ return r['판매상태'] !== '판매완료'; });

  const alerts = [];
  activeInv.forEach(function(item) {
    const days = parseInt(item['보관일수']) || 0;
    if (days >= 90) alerts.push({ type:'danger',  msg:'[재고] ' + item['모델'] + ' — ' + days + '일 보관. 즉시 가격조정 필요' });
    else if (days >= 60) alerts.push({ type:'warning', msg:'[재고] ' + item['모델'] + ' — ' + days + '일 보관. 검토 필요' });
  });
  if (buying.length)  alerts.push({ type:'info', msg:'매입 진행 중 ' + buying.length + '건' });
  if (selling.length) alerts.push({ type:'info', msg:'판매 진행 중 ' + selling.length + '건' });

  const totalProfit = done.reduce(function(s, r){ return s + (parseFloat(r['순수익']) || 0); }, 0);

  return {
    counts: {
      purchase: buying.length,
      sales:    selling.length,
      waitlist: activeWL.length,
      inventory: activeInv.length,
      purchaseWeek: pipeline.filter(function(r){ return new Date(r['유입일']) >= weekAgo; }).length,
      salesWeek:    selling.filter(function(r){ return new Date(r['유입일']) >= weekAgo; }).length,
      avgInventoryDays: 0,
      totalProfit: totalProfit,
      done: done.length,
    },
    alerts: alerts,
    channelMap: {},
    purchaseFunnel: calcFunnel(pipeline, PIPELINE_STAGES.slice(0,4), '단계'),
    salesFunnel:    calcFunnel(pipeline, PIPELINE_STAGES.slice(4),   '단계'),
    recentKPI: kpi.slice(-14),
  };
}

function calcFunnel(items, stages, field) {
  return stages.map(function(stage) {
    return { stage: stage, count: items.filter(function(r){ return r[field] === stage; }).length };
  });
}

// ============================================================
// 파이프라인 CRUD
// ============================================================

function getPipelineAll() {
  return sheetToObjects(getSheet(SHEETS.PIPELINE));
}

function addPipelineItem(data) {
  const sheet = getSheet(SHEETS.PIPELINE);
  if (!sheet) return { error: '파이프라인 시트 없음 — setupAll() 실행 필요' };

  const mgmtNo = generateMgmtNo(data.model);
  const isRe   = data.restructured === 'Y' || data.restructured === true;
  const settle  = calcSettlement(isRe);
  const profit  = calcProfit(data.buyPrice, data.salePrice, isRe);

  sheet.appendRow([
    mgmtNo,
    data.type        || '매입',
    data.model       || '',
    data.inboundDate || fmtDate(),
    data.workDate    || '',
    data.stage       || 'DM유입',
    data.buyPrice    || '',
    data.salePrice   || '',
    isRe ? 'Y' : 'N',
    settle,
    profit,
    data.memo        || '',
    fmtDate(),
  ]);

  return { success: true, id: mgmtNo };
}

function updatePipelineStage(id, stage) {
  const sheet  = getSheet(SHEETS.PIPELINE);
  const rowNum = findRow(sheet, id, '관리번호');
  if (rowNum < 0) return { error: 'ID 없음: ' + id };

  const stageCol = getColIndex(sheet, '단계');
  const modCol   = getColIndex(sheet, '최종수정일');
  if (stageCol > 0) sheet.getRange(rowNum, stageCol).setValue(stage);
  if (modCol   > 0) sheet.getRange(rowNum, modCol).setValue(fmtDate());
  return { success: true };
}

function updatePipelineField(id, field, value) {
  const sheet  = getSheet(SHEETS.PIPELINE);
  const rowNum = findRow(sheet, id, '관리번호');
  if (rowNum < 0) return { error: 'ID 없음: ' + id };

  const col = getColIndex(sheet, field);
  if (col > 0) sheet.getRange(rowNum, col).setValue(value);

  // 가격/구조변경 변경 시 정산·순수익 자동 재계산
  if (['매입가','판매가','구조변경'].indexOf(field) >= 0) {
    const ncols   = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, ncols).getValues()[0];
    const row     = sheet.getRange(rowNum, 1, 1, ncols).getValues()[0];
    const get     = function(f){ return row[headers.indexOf(f)]; };

    const isRe = (field === '구조변경' ? value : get('구조변경')) === 'Y';
    const buy  = field === '매입가' ? value : get('매입가');
    const sell = field === '판매가' ? value : get('판매가');

    const settleCol = getColIndex(sheet, '오랜케이정산');
    const profitCol = getColIndex(sheet, '순수익');
    if (settleCol > 0) sheet.getRange(rowNum, settleCol).setValue(calcSettlement(isRe));
    if (profitCol > 0) sheet.getRange(rowNum, profitCol).setValue(calcProfit(buy, sell, isRe));
  }

  const modCol = getColIndex(sheet, '최종수정일');
  if (modCol > 0) sheet.getRange(rowNum, modCol).setValue(fmtDate());
  return { success: true };
}

// ============================================================
// 재고 (기존 유지)
// ============================================================

function getInventory() {
  const items = sheetToObjects(getSheet(SHEETS.INVENTORY));
  const today = Date.now();
  return items.map(function(item) {
    if (item['매입일'] && item['판매상태'] !== '판매완료') {
      item['보관일수'] = Math.floor((today - new Date(item['매입일']).getTime()) / 86400000);
    }
    return item;
  });
}

function addInventory(data) {
  const sheet = getSheet(SHEETS.INVENTORY);
  const id = 'I_' + Date.now();
  sheet.appendRow([
    id, data.model, data.grade,
    data.purchasePrice || '', data.purchaseDate || new Date(),
    data.repairDetail || '', data.repairCost || 0,
    data.salePrice || '', data.status || '검수중',
    0, data.photoUrl || '', data.listingChannel || '', data.memo || ''
  ]);
  return { success: true, id: id };
}

// ============================================================
// 콘텐츠 (기존 유지)
// ============================================================

function getContent() { return sheetToObjects(getSheet(SHEETS.CONTENT)); }

function addContent(data) {
  getSheet(SHEETS.CONTENT).appendRow([
    data.date || new Date(), data.channel, data.type,
    data.title, data.url || '', data.status || '예정',
    0, 0, data.memo || ''
  ]);
  return { success: true };
}

// ============================================================
// 대기리스트 (매칭 기능 추가)
// ============================================================

function getWaitlist() { return sheetToObjects(getSheet(SHEETS.WAITLIST)); }

function addWaitlist(data) {
  getSheet(SHEETS.WAITLIST).appendRow([
    new Date(), data.name, data.contact,
    data.model || '', data.budget || '', '대기중', '', data.memo || ''
  ]);
  return { success: true };
}

function matchWaitlist(waitlistName, mgmtNo) {
  const sheet  = getSheet(SHEETS.WAITLIST);
  const rowNum = findRow(sheet, waitlistName, '이름');
  if (rowNum < 0) return { error: '대기리스트 항목 없음' };
  const matchCol  = getColIndex(sheet, '매칭관리번호');
  const statusCol = getColIndex(sheet, '알림상태');
  if (matchCol  > 0) sheet.getRange(rowNum, matchCol).setValue(mgmtNo);
  if (statusCol > 0) sheet.getRange(rowNum, statusCol).setValue('매칭완료');
  return { success: true };
}

// ============================================================
// KPI (기존 유지)
// ============================================================

function getKPI() { return sheetToObjects(getSheet(SHEETS.KPI)); }

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
// 시트 자동 생성
// ============================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensureSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a1d26')
        .setFontColor('#e8eaf0');
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  ensureSheet(SHEETS.PIPELINE, [
    '관리번호','유형','모델명','유입일','작업일','단계',
    '매입가','판매가','구조변경','오랜케이정산','순수익','메모','최종수정일'
  ]);
  ensureSheet(SHEETS.WAITLIST, [
    '날짜','이름','연락처','원하는모델','예산','알림상태','매칭관리번호','메모'
  ]);
  ensureSheet(SHEETS.INVENTORY, [
    '모델','등급','매입가','판매가','수리내역','수리비','리스팅채널','판매상태','보관일수','메모'
  ]);
  ensureSheet(SHEETS.CONTENT, [
    '채널','유형','제목','발행일','URL','상태','메모'
  ]);
  ensureSheet(SHEETS.KPI, [
    '날짜','매입건수','매출건수','순수익합계','재고수','대기자수'
  ]);

  return { ok: true };
}

// ============================================================
// Config
// ============================================================

function getConfig() {
  return {
    pipelineStages:  PIPELINE_STAGES,
    channels:        CHANNELS,
    modelOptions:    MODEL_CODE_MAP.map(function(m){ return m[0]; }),
    settlementItems: SETTLEMENT_ITEMS,
  };
}

// ============================================================
// 포스팅 승인 큐 (기존 유지)
// ============================================================

function setupApprovalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.APPROVAL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.APPROVAL);
    const headers = ['날짜','stem','media_type','folder_type','file','caption','thumbnail','status','post_id','processed_at'];
    const hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]).setBackground('#8B5CF6').setFontWeight('bold').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return { success: true, sheet: SHEETS.APPROVAL };
}

function getApprovalQueue() {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return [];
  return sheetToObjects(sheet).filter(function(r){ return r['status'] === 'ready'; });
}

function cleanupDuplicates() {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '시트 없음' };
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const stemIdx = headers.indexOf('stem');
  const seen = {}, toDelete = [];
  for (let i = 1; i < data.length; i++) {
    const stem   = String(data[i][stemIdx]);
    const status = data[i][headers.indexOf('status')];
    if (seen[stem] !== undefined) {
      if (status === 'ready') toDelete.push(i + 1);
      else if (seen[stem] === 'ready') toDelete.push(seen[stem + '_row']);
    } else {
      seen[stem] = status;
      seen[stem + '_row'] = i + 1;
    }
  }
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  return { success: true, deleted: toDelete.length };
}

function getPendingActions() {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return [];
  return sheetToObjects(sheet).filter(function(r){ return r['status'] === 'approved' || r['status'] === 'rejected'; });
}

function addApproval(data) {
  const sheet = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음 — setupApprovalSheet() 실행 필요' };

  const newRow = [
    data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data.stem        || '',
    data.media_type  || '',
    data.folder_type || '',
    data.file        || '',
    data.caption     || '',
    data.thumbnail   || '',
    'ready', '', '',
  ];

  const existingRow = _findApprovalRow(sheet, data.stem);
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, newRow.length).setValues([newRow]);
    return { success: true, updated: true };
  }
  sheet.appendRow(newRow);
  return { success: true };
}

function approveContent(stem, editedCaption) {
  const sheet  = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음' };
  const rowNum = _findApprovalRow(sheet, stem);
  if (rowNum < 0) return { error: 'stem 없음: ' + stem };

  const headers       = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol     = headers.indexOf('status') + 1;
  const processedCol  = headers.indexOf('processed_at') + 1;
  const captionCol    = headers.indexOf('caption') + 1;

  if (statusCol    > 0) sheet.getRange(rowNum, statusCol).setValue('approved');
  if (processedCol > 0) sheet.getRange(rowNum, processedCol).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  if (editedCaption && captionCol > 0) sheet.getRange(rowNum, captionCol).setValue(editedCaption);
  return { success: true };
}

function setApprovalStatus(stem, status) {
  const sheet  = getSheet(SHEETS.APPROVAL);
  if (!sheet) return { error: '포스팅_승인큐 시트 없음' };
  const rowNum = _findApprovalRow(sheet, stem);
  if (rowNum < 0) return { error: 'stem 없음: ' + stem };

  const headers      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol    = headers.indexOf('status') + 1;
  const processedCol = headers.indexOf('processed_at') + 1;

  if (statusCol    > 0) sheet.getRange(rowNum, statusCol).setValue(status);
  if (processedCol > 0) sheet.getRange(rowNum, processedCol).setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  return { success: true };
}

function markPosted(stem, postId) {
  const sheet  = getSheet(SHEETS.APPROVAL);
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
  const data    = sheet.getDataRange().getValues();
  const stemIdx = data[0].indexOf('stem');
  if (stemIdx < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][stemIdx]) === String(stem)) return i + 1;
  }
  return -1;
}

// ============================================================
// 콘텐츠 생성 트리거 (어드민 → Python)
// ============================================================

function setProcessTrigger(folder) {
  PropertiesService.getScriptProperties().setProperties({
    'PROCESS_TRIGGER_FOLDER': folder || 'all',
    'PROCESS_TRIGGER_TIME':   String(Date.now()),
    'PROCESS_TRIGGER_DONE':   'false'
  });
  return { success: true, folder: folder };
}

function getProcessTrigger() {
  const p      = PropertiesService.getScriptProperties();
  const done   = p.getProperty('PROCESS_TRIGGER_DONE');
  const folder = p.getProperty('PROCESS_TRIGGER_FOLDER');
  const ts     = p.getProperty('PROCESS_TRIGGER_TIME');
  if (!folder || done === 'true') return { pending: false };
  return { pending: true, folder: folder, ts: ts };
}

function clearProcessTrigger() {
  PropertiesService.getScriptProperties().setProperty('PROCESS_TRIGGER_DONE', 'true');
  return { success: true };
}

// ============================================================
// 스케줄 설정
// ============================================================

function getScheduleSettings() {
  const raw = PropertiesService.getScriptProperties().getProperty('SCHEDULE_SETTINGS');
  if (!raw) return { active: false, days: [], time: '10:00', type: '매일' };
  try { return JSON.parse(raw); } catch(e) { return { active: false, days: [], time: '10:00', type: '매일' }; }
}

function saveScheduleSettings(data) {
  PropertiesService.getScriptProperties().setProperty('SCHEDULE_SETTINGS', JSON.stringify(data));
  return { success: true };
}

// ============================================================
// 최초 세팅 — 스크립트 에디터에서 직접 실행
// ============================================================

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const configs = [
    {
      name: SHEETS.PIPELINE,
      headers: ['관리번호','유형','모델명','유입일','작업일','단계','매입가','판매가','구조변경','오랜케이정산','순수익','메모','최종수정일'],
      color: '#E85D5D'
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
      headers: ['날짜','이름','연락처','원하는모델','예산','알림상태','매칭관리번호','메모'],
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
    },
  ];

  configs.forEach(function(cfg) {
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

  ['Sheet1','시트1'].forEach(function(name) {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > configs.length) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  Logger.log('✅ 세팅 완료');
  Logger.log('스프레드시트 ID: ' + ss.getId());
}
