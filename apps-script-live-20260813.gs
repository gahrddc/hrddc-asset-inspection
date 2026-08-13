const DB_ID = '1LJ2Zj0Aev7I3SKI7NNMnwrgWMLv3ljAfaxVoftA8xi4';
const SOURCE_ID = '1RUp8M-ciTTiCtUnjDmKgzibIYhopMt4xQ4NArDmjfFU';
const SOURCE_SHEET = 'data';
const ADMIN_EMAIL = 'ga.hrddc@gmail.com';
const TZ = 'Asia/Bangkok';

// ===== Web entry and public API =====
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.api) return apiJsonp_(params);
  const mode = String(params.mode || '');
  if (mode !== 'admin') {
    return HtmlService.createHtmlOutput('<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="refresh" content="0;url=https://gahrddc.github.io/hrddc-asset-inspection/"></head><body><a href="https://gahrddc.github.io/hrddc-asset-inspection/">เปิดหน้าตรวจครุภัณฑ์</a></body></html>').setTitle('กำลังเปิดหน้าตรวจครุภัณฑ์');
  }
  return HtmlService.createHtmlOutput(buildHtml_()).setTitle('จัดการรอบตรวจครุภัณฑ์').addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function apiJsonp_(params) {
  const callback = /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(String(params.callback || '')) ? String(params.callback) : 'callback';
  let result;
  try {
    const action = String(params.api || '');
    if (action === 'config') result = getPublicConfig('');
    else if (action === 'find') result = findAsset(String(params.q || ''), '');
    else if (action === 'save') result = saveInspection(JSON.parse(String(params.payload || '{}')));
    else throw new Error('ไม่รู้จักคำสั่ง API');
  } catch (error) {
    result = {ok:false, error:error && error.message ? error.message : String(error)};
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===== Public inspection application =====
function getPublicConfig(accessCode) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'public_config_v1';
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }
  let round, roundState = 'ACTIVE';
  try { round = getActiveRound_(accessCode); }
  catch (error) {
    const nowMs = new Date().getTime();
    const openRounds = readObjects_(db_().getSheetByName('รอบตรวจ')).map(normalizeRoundObject_).filter(r => r.status === 'เปิด').sort((a,b) => a.startMs - b.startMs);
    round = openRounds.find(r => r.startMs > nowMs) || openRounds.slice().reverse().find(r => r.endMs < nowMs);
    if (!round) throw error;
    roundState = round.startMs > nowMs ? 'UPCOMING' : 'ENDED';
  }
  const statuses = readObjects_(db_().getSheetByName('รายการสถานะ'))
    .filter(x => truthy_(x['ใช้งาน']))
    .sort((a,b) => Number(a['ลำดับ']) - Number(b['ลำดับ']))
    .map(x => ({code:String(x['รหัส']), name:String(x['ชื่อสถานะ']), noteRequired:truthy_(x['ต้องระบุหมายเหตุ']), allowCustodian:truthy_(x['ให้เปลี่ยนผู้ครอบครอง']), color:String(x['สี'] || '#1e3a5f')}));
  const result = {ok:true, round:publicRound_(round), roundState:roundState, statuses:statuses};
  cache.put(cacheKey, JSON.stringify(result), 300);
  return result;
}

function findAsset(input, accessCode) {
  const round = getActiveRound_(accessCode);
  const assetId = extractAssetId_(input);
  if (!assetId) throw new Error('อ่านรหัสจาก QR ไม่ได้ กรุณากรอก ID หรือเลขครุภัณฑ์');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'asset:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assetId)).replace(/=+$/,'');
  let asset = null;
  const cached = cache.get(cacheKey);
  if (cached) { try { asset = JSON.parse(cached); } catch (_) {} }
  if (!asset) {
    const source = sourceSheet_();
    if (!source) throw new Error('ไม่พบแท็บข้อมูลต้นฉบับ');
    const lastRow = source.getLastRow();
    const lastColumn = source.getLastColumn();
    if (lastRow < 2) throw new Error('ชีทต้นฉบับไม่มีข้อมูล');
    const rowNumber = findSourceRowFromIndex_(source, assetId);
    if (!rowNumber) throw new Error('ไม่พบครุภัณฑ์รหัส ' + assetId);
    const headers = source.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    const row = source.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0];
    const idx = headerIndexes_(headers);
    asset = {id:String(row[idx.ID]).trim(),sourceStatus:row[idx.sourceStatus]||'',assetNo:row[idx.assetNo]||'',type:row[idx.type]||'',assetNumber:row[idx.assetNumber]||'',name:row[idx.name]||'',brand:row[idx.brand]||'',serial:row[idx.serial]||'',custodian:row[idx.custodian]||'',group:row[idx.group]||''};
    cache.put(cacheKey, JSON.stringify(asset), 1800);
  }
  const photo = getAssetPhoto_(asset.assetNo || asset.assetNumber || asset.id);
  asset.photoUrl = photo ? photo.url : '';
  return {ok:true, existing:findExisting_(round, asset.id), asset:asset};
}


function normalizeIndexKey_(value) {
  return String(value || '').trim().toLowerCase();
}

function indexBucket_(key) {
  let hash = 0;
  for (let i = 0; i < key.length; i++) hash = ((hash << 5) - hash + key.charCodeAt(i)) | 0;
  return Math.abs(hash) % 16;
}

function buildSourceIndex_(source) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('source_index_ready_v2')) return;
    const lastRow = source.getLastRow();
    const values = source.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    const buckets = Array.from({length:16}, () => ({}));
    values.forEach((row, i) => {
      [row[0], row[2]].forEach(value => {
        const key = normalizeIndexKey_(value);
        if (key) buckets[indexBucket_(key)][key] = i + 2;
      });
    });
    const payload = {};
    buckets.forEach((bucket, i) => payload['source_index_v2_' + i] = JSON.stringify(bucket));
    cache.putAll(payload, 3600);
    cache.put('source_index_ready_v2', '1', 3600);
  } finally {
    lock.releaseLock();
  }
}

function findSourceRowFromIndex_(source, value) {
  const key = normalizeIndexKey_(value);
  const cache = CacheService.getScriptCache();
  const rowCacheKey = 'source_row_v3_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)
  ).replace(/=+$/,'');
  const cachedRow = Number(cache.get(rowCacheKey) || 0);
  if (cachedRow) return cachedRow;
  const lastRow = source.getLastRow();
  let found = source.getRange(2, 1, lastRow - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  if (!found) found = source.getRange(2, 3, lastRow - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  const row = found ? found.getRow() : 0;
  if (row) cache.put(rowCacheKey, String(row), 21600);
  return row;
}

// ===== Inspection persistence =====
function saveInspection(payload) {
  payload = payload || {};
  if (!String(payload.inspector || '').trim()) throw new Error('กรุณาระบุชื่อผู้ตรวจ');
  if (!String(payload.assetId || '').trim()) throw new Error('ไม่พบรหัสครุภัณฑ์');
  if (!String(payload.status || '').trim()) throw new Error('กรุณาเลือกผลการตรวจ');
  const requestId = String(payload.requestId || '').trim();
  const cache = CacheService.getScriptCache();
  const requestKey = requestId ? 'save_request_' + requestId : '';
  if (requestKey) {
    const previous = cache.get(requestKey);
    if (previous) {
      try { return JSON.parse(previous); } catch (_) {}
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (requestKey) {
      const completed = cache.get(requestKey);
      if (completed) {
        try { return JSON.parse(completed); } catch (_) {}
      }
    }
    const round = getActiveRound_(payload.accessCode);
    const sheet = db_().getSheetByName(round.sheetName);
    if (!sheet) throw new Error('ไม่พบชีทผลตรวจของรอบนี้');
    const existing = findExisting_(round, String(payload.assetId));
    if (existing) throw new Error('ครุภัณฑ์รายการนี้ตรวจแล้วในรอบนี้ โดย ' + existing.inspector + ' เมื่อ ' + existing.when);
    const statusInfo = getStatus_(payload.status);
    if (statusInfo.noteRequired && !String(payload.note || '').trim()) throw new Error('สถานะนี้ต้องระบุหมายเหตุ');
    const now = new Date();
    sheet.appendRow([
      Utilities.getUuid(), now, round.id, String(payload.assetId), String(payload.assetNo || ''), String(payload.assetName || ''),
      String(payload.oldCustodian || ''), String(payload.oldGroup || ''), String(payload.status), String(statusInfo.name),
      String(payload.newCustodian || ''), String(payload.newGroup || ''), String(payload.note || ''),
      String(payload.inspector || '').trim(), String(payload.verificationCode || ''), String(payload.qrRaw || ''), 'บันทึกแล้ว'
    ]);
    cache.remove('admin_notifications_v1');
    const result = {ok:true, message:'บันทึกผลเรียบร้อย', when:Utilities.formatDate(now,TZ,'dd/MM/yyyy HH:mm:ss')};
    if (requestKey) cache.put(requestKey, JSON.stringify(result), 21600);
    log_('ตรวจครุภัณฑ์', round.id, String(payload.assetId), statusInfo.name + ' โดย ' + payload.inspector, String(payload.inspector));
    return result;
  } finally {
    lock.releaseLock();
  }
}

// ===== Administrator operations =====
function getAdminNotifications() {
  assertAdmin_();
  const cache = CacheService.getScriptCache();
  const cached = cache.get('admin_notifications_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }
  const db = db_();
  const rounds = readObjects_(sheetOrThrow_(db, 'รอบตรวจ')).map(normalizeRoundObject_);
  let items = [];
  rounds.forEach(round => {
    const sheet = db.getSheetByName(round.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const take = Math.min(10, sheet.getLastRow() - 1);
    const rows = sheet.getRange(sheet.getLastRow() - take + 1, 1, take, 17).getValues();
    rows.forEach(r => {
      if (String(r[16] || '').trim() === 'ยกเลิก') return;
      const when = r[1] instanceof Date ? r[1] : new Date(r[1]);
      items.push({
        id:String(r[0] || ''),
        timestamp:isNaN(when.getTime()) ? 0 : when.getTime(),
        date:isNaN(when.getTime()) ? String(r[1] || '') : Utilities.formatDate(when,TZ,'dd/MM/yyyy HH:mm'),
        assetNo:String(r[4] || r[3] || ''),
        name:String(r[5] || ''),
        inspector:String(r[13] || ''),
        result:String(r[9] || ''),
        roundName:String(round.name || '')
      });
    });
  });
  items.sort((a,b)=>b.timestamp-a.timestamp);
  const result = {ok:true,items:items.slice(0,20)};
  cache.put('admin_notifications_v1', JSON.stringify(result), 45);
  return result;
}

function getAdminBootstrap() {
  assertAdmin_();
  const db = db_();
  ensureAssetPhotoSheet_();
  const rounds = readObjects_(sheetOrThrow_(db, 'รอบตรวจ')).map(normalizeRoundObject_);
  return {ok:true,email:Session.getActiveUser().getEmail(),rounds:rounds,home:getAdminHomeData_(db,rounds)};
}

function getAdminHomeData_(db, rounds) {
  const cache = CacheService.getScriptCache();
  let totalAssets = Number(cache.get('asset_total_count'));
  if (!totalAssets) {
    totalAssets = Math.max(0, sourceSheet_().getLastRow() - 1);
    cache.put('asset_total_count', String(totalAssets), 21600);
  }
  const selected = rounds.find(r=>r.status==='เปิด') || rounds[rounds.length-1] || null;
  const home = {totalAssets:totalAssets,inspected:0,damaged:0,moved:0,recent:[]};
  if (!selected) return home;
  const sheet = db.getSheetByName(selected.sheetName);
  if (!sheet || sheet.getLastRow() < 2) return home;
  const count = sheet.getLastRow() - 1;
  home.inspected = count;
  const codes = sheet.getRange(2,9,count,1).getDisplayValues();
  codes.forEach(r=>{const code=String(r[0]).toUpperCase();if(code==='DAMAGED')home.damaged++;if(code==='MOVED')home.moved++});
  const take = Math.min(5,count);
  const rows = sheet.getRange(sheet.getLastRow()-take+1,1,take,17).getDisplayValues().reverse();
  home.recent = rows.map(r=>({date:r[1],assetNo:r[4],name:r[5],inspector:r[13],result:r[9]}));
  return home;
}

function createRoundAdmin(form) {
  assertAdmin_();
  form = form || {};
  const name = String(form.name || '').trim();
  const start = parseDate_(form.start);
  const end = parseDate_(form.end);
  if (!name || !start || !end) throw new Error('กรุณาระบุชื่อรอบและวันเริ่ม-สิ้นสุด');
  if (end.getTime() < start.getTime()) throw new Error('วันสิ้นสุดต้องไม่น้อยกว่าวันเริ่ม');
  const yearBE = Number(Utilities.formatDate(start,TZ,'yyyy')) + 543;
  const db = db_();
  const roundsSheet = db.getSheetByName('รอบตรวจ');
  const id = 'R' + Utilities.formatDate(new Date(),TZ,'yyyyMMddHHmmss');
  let sheetName = 'ผลตรวจ_' + yearBE;
  if (db.getSheetByName(sheetName)) sheetName += '_' + id.slice(-4);
  const result = db.insertSheet(sheetName);
  result.getRange(1,1,1,17).setValues([['รายการ','วันเวลา','รหัสรอบ','ID ครุภัณฑ์','เลขครุภัณฑ์','ชื่อครุภัณฑ์','ผู้ครอบครองเดิม','กลุ่มเดิม','รหัสสถานะ','ผลการตรวจ','ผู้ครอบครองใหม่','กลุ่มใหม่','หมายเหตุ','ผู้ตรวจ','รหัสยืนยัน','ข้อมูล QR','สถานะรายการ']]);
  formatSheet_(result);
  roundsSheet.appendRow([id,name,start,end,String(form.status || 'ปิด'),sheetName,String(form.accessCode || ''),new Date(),Session.getActiveUser().getEmail()]);
  CacheService.getScriptCache().removeAll(['active_round','public_config_v1']);
  log_('สร้างรอบตรวจ',id,'',name,Session.getActiveUser().getEmail());
  return {ok:true,id:id,sheetName:sheetName};
}

function updateRoundAdmin(form) {
  assertAdmin_();
  form = form || {};
  const id = String(form.id || '').trim();
  const name = String(form.name || '').trim();
  const start = parseDate_(form.start);
  const end = parseDate_(form.end);
  if (!id || !name || !start || !end) throw new Error('กรุณาระบุชื่อรอบและวันเริ่ม-สิ้นสุด');
  if (end.getTime() < start.getTime()) throw new Error('วันสิ้นสุดต้องไม่น้อยกว่าวันเริ่ม');
  const sheet = sheetOrThrow_(db_(), 'รอบตรวจ');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[name, start, end]]);
      sheet.getRange(i + 1, 7).setValue(String(form.accessCode || ''));
      CacheService.getScriptCache().removeAll(['active_round','public_config_v1']);
      log_('แก้ไขรอบตรวจ', id, '', name, Session.getActiveUser().getEmail());
      return {ok:true,id:id,name:name};
    }
  }
  throw new Error('ไม่พบรอบตรวจ');
}

function setRoundStatusAdmin(roundId, status) {
  assertAdmin_();
  const sheet = db_().getSheetByName('รอบตรวจ');
  const values = sheet.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]) === String(roundId)) {
      sheet.getRange(i+1,5).setValue(String(status));
      CacheService.getScriptCache().removeAll(['active_round','public_config_v1']);
      log_('เปลี่ยนสถานะรอบ',roundId,'',String(status),Session.getActiveUser().getEmail());
      return {ok:true};
    }
  }
  throw new Error('ไม่พบรอบตรวจ');
}

function deleteRoundAdmin(roundId) {
  assertAdmin_();
  const db = db_();
  const roundsSheet = sheetOrThrow_(db, 'รอบตรวจ');
  const values = roundsSheet.getDataRange().getValues();
  let targetRow = -1;
  let sheetName = '';
  let roundName = '';
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(roundId)) {
      targetRow = i + 1;
      roundName = String(values[i][1] || '');
      sheetName = String(values[i][5] || '');
      break;
    }
  }
  if (targetRow < 0) throw new Error('ไม่พบรอบตรวจ');
  const resultSheet = sheetName ? db.getSheetByName(sheetName) : null;
  if (resultSheet) db.deleteSheet(resultSheet);
  roundsSheet.deleteRow(targetRow);
  CacheService.getScriptCache().removeAll(['active_round','public_config_v1']);
  log_('ลบรอบตรวจ', String(roundId), '', roundName, Session.getActiveUser().getEmail());
  return {ok:true, id:String(roundId), sheetName:sheetName};
}

function getDashboardAdmin(roundId) {
  assertAdmin_();
  const round = getRoundById_(String(roundId));
  const resultSheet = sheetOrThrow_(db_(), round.sheetName);
  const rows = resultSheet.getLastRow() > 1 ? resultSheet.getRange(2,1,resultSheet.getLastRow()-1,17).getDisplayValues() : [];
  const activeRows = rows.filter(r => String(r[16] || '').trim() !== 'ยกเลิก');
  const records = activeRows.slice().reverse().map(r => ({
    assetNo:String(r[4] || r[3] || ''),
    name:String(r[5] || ''),
    inspector:String(r[13] || ''),
    inspectedAt:String(r[1] || ''),
    result:String(r[9] || '')
  }));
  const inspected = activeRows.length;
  const cache = CacheService.getScriptCache();
  let total = Number(cache.get('asset_total_count'));
  if (!total) {
    total = Math.max(0, sourceSheet_().getLastRow() - 1);
    cache.put('asset_total_count', String(total), 21600);
  }
  return {ok:true,total:total,inspected:inspected,remaining:Math.max(0,total-inspected),records:records,roundName:round.name};
}

function getActiveRound_(accessCode) {
  const cache = CacheService.getScriptCache();
  let active = null;
  const cached = cache.get('active_round');
  if (cached) { try { active = JSON.parse(cached); } catch (_) {} }
  const now = new Date();
  if (!active || active.status !== 'เปิด' || now.getTime() < active.startMs || now.getTime() > active.endMs) {
    const rounds = readObjects_(db_().getSheetByName('รอบตรวจ')).map(normalizeRoundObject_);
    active = rounds.find(r => r.status === 'เปิด' && now.getTime() >= r.startMs && now.getTime() <= r.endMs);
    if (active) cache.put('active_round', JSON.stringify(active), 60);
  }
  if (!active) throw new Error('ขณะนี้ไม่มีรอบตรวจที่เปิดใช้งาน');
  if (active.accessCode && String(accessCode || '') !== active.accessCode) throw new Error('รหัสเข้ารอบตรวจไม่ถูกต้อง');
  return active;
}

function getRoundById_(id) {
  const round = readObjects_(db_().getSheetByName('รอบตรวจ')).map(normalizeRoundObject_).find(r => r.id === String(id));
  if (!round) throw new Error('ไม่พบรอบตรวจ');
  return round;
}

function normalizeRoundObject_(x) {
  const start = new Date(x['วันเริ่ม']);
  const end = new Date(x['วันสิ้นสุด']);
  return {id:String(x['รหัสรอบ']||''),name:String(x['ชื่อรอบ']||''),start:formatDateSafe_(start),end:formatDateSafe_(end),startMs:start.getTime(),endMs:end.getTime(),status:String(x['สถานะ']||''),sheetName:String(x['ชื่อชีทผล']||''),accessCode:String(x['รหัสเข้ารอบ']||'')};
}

function publicRound_(r) { return {id:r.id,name:r.name,start:r.start,end:r.end,status:r.status}; }

function findExisting_(round, assetId) {
  const sheet = db_().getSheetByName(round.sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const matches = sheet.getRange(2,4,sheet.getLastRow()-1,1).createTextFinder(String(assetId).trim()).matchEntireCell(true).findAll();
  for (let i=matches.length-1;i>=0;i--) {
    const row = sheet.getRange(matches[i].getRow(),1,1,17).getDisplayValues()[0];
    if (String(row[16]) !== 'ยกเลิก') return {when:row[1],status:row[9],inspector:row[13],note:row[12]};
  }
  return null;
}

function getStatus_(code) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'active_statuses_v1';
  let rows = null;
  const cached = cache.get(cacheKey);
  if (cached) { try { rows = JSON.parse(cached); } catch (_) {} }
  if (!rows) {
    rows = readObjects_(db_().getSheetByName('รายการสถานะ'))
      .filter(x => truthy_(x['ใช้งาน']))
      .map(x => ({code:String(x['รหัส']),name:String(x['ชื่อสถานะ']),noteRequired:truthy_(x['ต้องระบุหมายเหตุ'])}));
    cache.put(cacheKey, JSON.stringify(rows), 1800);
  }
  const found = rows.find(x => String(x.code) === String(code));
  if (!found) throw new Error('สถานะไม่ถูกต้อง');
  return {name:String(found.name),noteRequired:Boolean(found.noteRequired)};
}

function extractAssetId_(input) {
  const raw = String(input || '').trim();
  let s = raw;
  try { s = decodeURIComponent(raw); } catch (_) {}
  const m = s.match(/(?:[?#&]|^)row=([^&#]+)/i);
  if (m) {
    try { return decodeURIComponent(m[1]).trim(); } catch (_) { return String(m[1]).trim(); }
  }
  if (!s || s.length > 200 || /[\u0000-\u001F\u007F]/.test(s)) return '';
  return s;
}

function headerIndexes_(h) {
  function ix(name) { const n=h.indexOf(name); if(n<0) throw new Error('ไม่พบคอลัมน์ '+name); return n; }
  return {ID:ix('ID'),sourceStatus:ix('สถานะ'),assetNo:ix('เลขครุภัณฑ์'),type:ix('ประเภท'),assetNumber:ix('เลขสินทรัพย์'),name:ix('ชื่อครุภัณฑ์'),brand:ix('ยี่ห้อ/รุ่น'),serial:ix('Serial No.'),custodian:ix('ผู้รับผิดชอบ'),group:ix('กลุ่ม')};
}

// ===== Shared data utilities =====
function readObjects_(sheet) {
  if (!sheet || sheet.getLastRow()<2) return [];
  const v=sheet.getDataRange().getValues(), h=v[0].map(String);
  return v.slice(1).filter(r=>r.some(x=>x!=='')) .map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;});
}

function truthy_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v) === '1' || String(v) === 'ใช่'; }
function parseDate_(s) { if(!s)return null; const d=new Date(String(s)+'T00:00:00+07:00'); return isNaN(d.getTime())?null:d; }
function formatDateSafe_(d) { return isNaN(d.getTime())?'':Utilities.formatDate(d,TZ,'dd/MM/yyyy HH:mm'); }
function assertAdmin_() { const email=String(Session.getActiveUser().getEmail()||'').toLowerCase(); if(email!==ADMIN_EMAIL) throw new Error('ไม่มีสิทธิ์แอดมิน กรุณาเข้าสู่ระบบด้วย '+ADMIN_EMAIL); }
function log_(action,roundId,assetId,detail,actor) { db_().getSheetByName('ประวัติการแก้ไข').appendRow([new Date(),action,roundId,assetId,detail,actor]); }


// ===== Asset photo management =====
function ensureAssetPhotoSheet_() {
  const db = db_();
  let sheet = db.getSheetByName('รูปครุภัณฑ์');
  if (!sheet) {
    sheet = db.insertSheet('รูปครุภัณฑ์');
    sheet.getRange(1,1,1,7).setValues([['เลขครุภัณฑ์','ID ครุภัณฑ์','ชื่อครุภัณฑ์','File ID','URL รูป','อัปเดตเมื่อ','ผู้อัปเดต']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#0f5132').setFontColor('#ffffff');
    sheet.autoResizeColumns(1,7);
  }
  return sheet;
}

function getAssetPhoto_(assetNo) {
  const key = String(assetNo || '').trim();
  if (!key) return null;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'asset_photo_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)).replace(/=+$/,'');
  const cached = cache.get(cacheKey);
  if (cached) {
    if (cached === 'NONE') return null;
    try { return JSON.parse(cached); } catch (_) {}
  }
  const sheet = db_().getSheetByName('รูปครุภัณฑ์');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const found = sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(key).matchEntireCell(true).findNext();
  if (found) {
    const row = sheet.getRange(found.getRow(),1,1,7).getDisplayValues()[0];
    const result = {assetNo:key,fileId:row[3],url:row[4],updatedAt:row[5],updatedBy:row[6]};
    cache.put(cacheKey, JSON.stringify(result), 1800);
    return result;
  }
  cache.put(cacheKey, 'NONE', 300);
  return null;
}

function getSourceAssetAdmin_(input) {
  const value = extractAssetId_(input);
  if (!value) throw new Error('กรุณาระบุเลขครุภัณฑ์');
  const source = sourceSheet_();
  const rowNumber = findSourceRowFromIndex_(source, value);
  if (!rowNumber) throw new Error('ไม่พบครุภัณฑ์รหัส ' + value);
  const lastColumn = source.getLastColumn();
  const headers = source.getRange(1,1,1,lastColumn).getDisplayValues()[0];
  const row = source.getRange(rowNumber,1,1,lastColumn).getDisplayValues()[0];
  const idx = headerIndexes_(headers);
  return {id:String(row[idx.ID]).trim(),assetNo:row[idx.assetNo]||row[idx.assetNumber]||'',name:row[idx.name]||'',type:row[idx.type]||'',custodian:row[idx.custodian]||''};
}

function getAssetPhotoAdmin(input) {
  assertAdmin_();
  const asset = getSourceAssetAdmin_(input);
  return {ok:true,asset:asset,photo:getAssetPhoto_(asset.assetNo || asset.id)};
}

function getAssetPhotoListAdmin() {
  assertAdmin_();
  const source = sourceSheet_();
  if (source.getLastRow() < 2) return {ok:true,items:[],total:0,withPhoto:0,withoutPhoto:0};
  const lastColumn = source.getLastColumn();
  const headers = source.getRange(1,1,1,lastColumn).getDisplayValues()[0];
  const idx = headerIndexes_(headers);
  const rows = source.getRange(2,1,source.getLastRow()-1,lastColumn).getDisplayValues();
  const photoSheet = db_().getSheetByName('รูปครุภัณฑ์');
  const photoMap = {};
  if (photoSheet && photoSheet.getLastRow() > 1) {
    const photos = photoSheet.getRange(2,1,photoSheet.getLastRow()-1,7).getDisplayValues();
    photos.forEach(r => {
      const key = String(r[0] || '').trim();
      if (key) photoMap[key] = {fileId:r[3],url:r[4],updatedAt:r[5],updatedBy:r[6]};
    });
  }
  const items = rows.filter(r => r.some(v => v !== '')).map(r => {
    const assetNo = String(r[idx.assetNo] || r[idx.assetNumber] || '').trim();
    const id = String(r[idx.ID] || '').trim();
    const photo = photoMap[assetNo] || photoMap[id] || null;
    return {id:id,assetNo:assetNo,name:r[idx.name]||'',type:r[idx.type]||'',custodian:r[idx.custodian]||'',photo:photo};
  });
  const withPhoto = items.reduce((n,x)=>n+(x.photo&&x.photo.url?1:0),0);
  return {ok:true,items:items,total:items.length,withPhoto:withPhoto,withoutPhoto:items.length-withPhoto};
}

function getAssetPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('ASSET_PHOTO_FOLDER_ID');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (_) {} }
  const folder = DriveApp.createFolder('รูปครุภัณฑ์_ระบบตรวจประจำปี');
  props.setProperty('ASSET_PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function saveAssetPhotoAdmin(payload) {
  assertAdmin_();
  payload = payload || {};
  const asset = getSourceAssetAdmin_(payload.assetNo);
  const dataUrl = String(payload.dataUrl || '');
  const match = dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('ไฟล์รูปไม่ถูกต้อง กรุณาเลือก JPG, PNG หรือ WEBP');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5000000) throw new Error('รูปมีขนาดใหญ่เกิน 5 MB กรุณาเลือกรูปใหม่');
  const ext = match[1] === 'image/png' ? 'png' : match[1] === 'image/webp' ? 'webp' : 'jpg';
  const safeNo = String(asset.assetNo || asset.id).replace(/[^0-9A-Za-zก-๙._()-]+/g,'_');
  const blob = Utilities.newBlob(bytes, match[1], safeNo + '_' + Utilities.formatDate(new Date(),TZ,'yyyyMMdd_HHmmss') + '.' + ext);
  const file = getAssetPhotoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
  const sheet = ensureAssetPhotoSheet_();
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,7).getDisplayValues() : [];
  let targetRow = sheet.getLastRow() + 1;
  for (let i=0;i<rows.length;i++) if (String(rows[i][0]).trim() === String(asset.assetNo).trim()) { targetRow=i+2; break; }
  const actor = Session.getActiveUser().getEmail();
  sheet.getRange(targetRow,1,1,7).setValues([[asset.assetNo,asset.id,asset.name,file.getId(),url,new Date(),actor]]);
  removeAssetPhotoCache_(asset.assetNo || asset.id);
  log_('อัปเดตรูปครุภัณฑ์','',asset.id,asset.assetNo,actor);
  return {ok:true,asset:asset,photo:{assetNo:asset.assetNo,fileId:file.getId(),url:url,updatedAt:formatDateSafe_(new Date()),updatedBy:actor}};
}
function deleteAssetPhotoAdmin(input) {
  assertAdmin_();
  const asset = getSourceAssetAdmin_(input);
  const key = String(asset.assetNo || asset.id || '').trim();
  const sheet = db_().getSheetByName('รูปครุภัณฑ์');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('รายการนี้ยังไม่มีรูปให้ลบ');
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,7).getDisplayValues();
  let target = -1, fileId = '';
  for (let i=rows.length-1;i>=0;i--) {
    if (String(rows[i][0]||'').trim()===key || String(rows[i][1]||'').trim()===String(asset.id||'').trim()) { target=i+2; fileId=rows[i][3]; break; }
  }
  if (target < 0) throw new Error('รายการนี้ยังไม่มีรูปให้ลบ');
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {} }
  sheet.deleteRow(target);
  removeAssetPhotoCache_(key);
  const actor = Session.getActiveUser().getEmail();
  log_('ลบรูปครุภัณฑ์','',asset.id,key,actor);
  return {ok:true,asset:asset,photo:null};
}

function authorizePhotoDrive() {
  return getAssetPhotoFolder_().getId();
}

function removeAssetPhotoCache_(key) {
  key = String(key || '').trim();
  if (!key) return;
  const cacheKey = 'asset_photo_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)).replace(/=+$/,'');
  CacheService.getScriptCache().remove(cacheKey);
}


// ===== Data access and formatting =====
function db_() {
  return SpreadsheetApp.openById(DB_ID);
}

function sourceSheet_() {
  const sheet = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('ไม่พบแท็บข้อมูลต้นฉบับ');
  return sheet;
}

function sheetOrThrow_(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error('ไม่พบชีท ' + name);
  return sheet;
}

function formatSheet_(sheet) {
  if (!sheet) throw new Error('ไม่พบชีทสำหรับจัดรูปแบบ');
  sheet.setFrozenRows(1);
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, lastColumn);
  header.setBackground('#075f36').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(1, 1, Math.max(1, sheet.getMaxRows()), lastColumn).setVerticalAlignment('middle');
  sheet.autoResizeColumns(1, lastColumn);
}

// ===== Administrator web interface =====
function buildHtml_() {
  return `<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr@4.6.13/dist/flatpickr.min.css">
  <style>
  :root{--green:#078942;--deep:#006b38;--ink:#182238;--muted:#697386;--line:#dce4e1;--soft:#eff8f3;--bg:#f5f8f6}*{box-sizing:border-box}html{overflow-y:scroll;scrollbar-gutter:stable}body{margin:0;min-height:100vh;font-family:"Kanit",sans-serif;background:var(--bg);color:var(--ink)}.top{background:linear-gradient(120deg,#006838,#087b47);color:#fff;padding:24px 32px}.brand{max-width:1240px;margin:auto;display:flex;align-items:center;justify-content:space-between}.brand-left{display:flex;align-items:center;gap:16px}.logo{width:58px;height:58px;border:2px solid #fff;border-radius:50%;display:grid;place-items:center;font-size:20px;font-weight:700;box-shadow:0 4px 16px #003c2455}.brand h1{font-size:24px;margin:0}.brand p{margin:2px 0 0;font-size:17px}.top-action{border:1px solid #fff;border-radius:12px;padding:12px 18px;font-weight:600}.wrap{width:100%;max-width:1240px;min-height:calc(100vh - 106px);margin:0 auto;padding:28px 32px 44px}.msg{padding:14px 20px;border:1px solid #cfe7da;border-radius:12px;margin:0 0 20px;background:var(--soft);color:#145c38}.msg.error{background:#fff0f0;border-color:#ffcaca;color:#a41f1f}.admin-tabs{display:inline-flex;align-items:center;gap:4px;width:auto;margin:0 0 20px;padding:4px;border:1px solid #dce7e1;border-radius:14px;background:#eef5f1;box-shadow:0 3px 10px #0c4b2b0a}.admin-tabs .btn{min-width:190px;border-radius:10px;padding:10px 18px}.admin-tabs .secondary{background:transparent;color:#52605a}.card{background:#fff;border:1px solid var(--line);border-radius:18px;padding:24px;margin:0 0 22px;box-shadow:0 5px 18px #0c4b2b0d}.section-title{display:flex;align-items:flex-start;gap:16px;margin-bottom:22px}.step{width:48px;height:48px;flex:0 0 48px;border-radius:50%;display:grid;place-items:center;background:linear-gradient(145deg,#0da052,#00733b);color:#fff;font-size:22px;font-weight:600}.section-title h2{margin:0;color:#087a42;font-size:24px}.section-title p{margin:0;color:var(--muted)}.row{display:grid;grid-template-columns:1fr 1fr;gap:18px}.field{margin:12px 0}label{display:block;color:#087a42;font-weight:600;margin-bottom:7px}input,select,textarea,button{font:inherit}input,select,textarea{width:100%;border:1px solid #d4dde3;border-radius:12px;padding:14px 16px;background:#fff;color:var(--ink);outline:none}input:focus,select:focus,textarea:focus{border-color:#12a25a;box-shadow:0 0 0 3px #11a25a18}.btn{border:0;border-radius:10px;padding:11px 17px;font-weight:600;cursor:pointer}.primary,.active{background:linear-gradient(135deg,#08a24e,#007b3e)!important;color:#fff!important}.secondary{background:#eef2f4;color:#3d4a5e}.outline{background:#fff;color:#087a42;border:1px solid #0b9a50}.hidden{display:none!important}.round{display:grid;grid-template-columns:78px 1fr auto;gap:18px;align-items:center;border:1px solid #d8e2e5;border-radius:14px;padding:18px 22px;margin:12px 0}.round-icon{width:62px;height:62px;border-radius:50%;background:#e7f6ee;display:grid;place-items:center;color:#07934a;font-size:30px}.round-main b{font-size:19px}.round-meta{color:#58657a;margin:4px 0 10px}.round-actions{display:flex;gap:8px;flex-wrap:wrap}.pill{display:inline-block;padding:3px 10px;border-radius:999px;background:#e1f7e9;color:#087d42;font-size:13px;font-weight:600}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.num{background:#eff8f3;border-radius:12px;text-align:center;padding:14px}.num b{font-size:25px;display:block;color:#087a42}.photo-preview{display:block;width:100%;max-height:420px;object-fit:contain;margin-top:15px;border-radius:14px;background:#eef6f1}.photo-info{margin-top:14px;padding:14px;border-radius:12px;background:#eff8f3}.photo-actions{display:flex;gap:9px;flex-wrap:wrap}.file-label{display:inline-block}#homePanel,#inspectPanel,#photoPanel{width:100%;min-height:640px}.empty{text-align:center;color:#748091;padding:28px}.kebab{color:#617084;font-size:24px;align-self:start}.dash{scroll-margin-top:15px}@media(max-width:700px){.top{padding:18px 16px}.brand{align-items:flex-start}.logo{width:48px;height:48px;flex-basis:48px}.brand h1{font-size:18px}.brand p{font-size:14px}.top-action{display:none}.wrap{padding:18px 14px 34px;min-height:calc(100vh - 84px)}.row{grid-template-columns:1fr;gap:0}.card{padding:18px 15px}.section-title h2{font-size:20px}.step{width:42px;height:42px;flex-basis:42px}.admin-tabs{display:grid;grid-template-columns:1fr 1fr;width:100%}.admin-tabs .btn{min-width:0}.round{grid-template-columns:52px 1fr;padding:15px}.round-icon{width:48px;height:48px;font-size:24px}.kebab{display:none}.round-actions{grid-column:1/-1}.round-actions .btn{flex:1}.summary{grid-template-columns:1fr}#homePanel,#inspectPanel,#photoPanel{min-height:520px}.brand-left{gap:10px}}
  .action-modal{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;padding:20px;background:#0b2f2290}.action-box{width:min(430px,100%);padding:28px 24px;border-radius:20px;background:#fff;text-align:center;box-shadow:0 20px 60px #001d1244}.action-box.error{color:#a41f1f}.spinner{width:34px;height:34px;margin:0 auto 16px;border:4px solid #dcefe4;border-top-color:#078942;border-radius:50%;animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
/* System dashboard layout */
body{background:#f3f7f5}.top{position:relative;padding:19px 30px;box-shadow:0 2px 18px #053d2614}.brand{max-width:1440px}.wrap{max-width:1440px;padding:24px 28px 48px}.msg{margin-bottom:18px}#adminBody{display:grid;grid-template-columns:248px minmax(0,1fr);gap:22px;align-items:start}.admin-tabs{position:sticky;top:20px;display:flex;flex-direction:column;align-items:stretch;width:100%;margin:0;padding:14px;gap:7px;border:0;border-radius:18px;background:linear-gradient(165deg,#075f39,#064c31);box-shadow:0 12px 30px #073e2726}.admin-tabs:before{content:'เมนูจัดการ';padding:9px 12px 12px;color:#bfe5d1;font-size:13px}.admin-tabs .btn{width:100%;min-width:0;padding:13px 15px;text-align:left;border-radius:11px}.admin-tabs .active{background:#fff!important;color:#07683d!important;box-shadow:0 6px 15px #002d1c35}.admin-tabs .secondary{background:transparent;color:#e5f5ed}.admin-tabs .secondary:hover{background:#ffffff14}.page-heading{display:flex;align-items:flex-end;justify-content:space-between;gap:18px;margin:2px 0 20px}.page-heading h1{margin:2px 0 3px;color:#163b2b;font-size:30px}.page-heading p{margin:0;color:#6b7972}.eyebrow{color:#0b9150!important;font-size:13px;font-weight:600;letter-spacing:.04em}.page-badge{padding:9px 13px;border:1px solid #d7e8de;border-radius:999px;background:#fff;color:#0a7845;font-size:13px}.admin-summary{display:grid;grid-template-columns:repeat(3,1fr);gap:13px;margin-bottom:18px}.metric{position:relative;overflow:hidden;padding:17px 18px;border:1px solid #dfe9e3;border-radius:15px;background:#fff;box-shadow:0 5px 15px #073e270a}.metric:after{content:'';position:absolute;right:-18px;top:-22px;width:72px;height:72px;border-radius:50%;background:#e8f7ef}.metric span{display:block;color:#748179;font-size:13px}.metric b{display:block;margin-top:3px;color:#075f39;font-size:27px}.card{padding:22px;border-radius:17px;box-shadow:0 6px 20px #073e270b}.section-title{margin-bottom:18px}.section-title h2{font-size:21px}.step{width:42px;height:42px;flex-basis:42px;font-size:18px}.round{background:#fbfdfc}.round:hover{border-color:#9bd2b5;box-shadow:0 7px 20px #073e2710;transform:translateY(-1px);transition:.18s}.photo-actions .btn{min-width:130px}@media(max-width:850px){.wrap{padding:18px 15px 38px}#adminBody{grid-template-columns:1fr}.admin-tabs{position:static;display:grid;grid-template-columns:1fr 1fr;padding:4px;background:#eaf3ee;border:1px solid #dce7e1}.admin-tabs:before{display:none}.admin-tabs .btn{text-align:center}.admin-tabs .active{background:linear-gradient(135deg,#08a24e,#007b3e)!important;color:#fff!important}.admin-tabs .secondary{color:#52605a}.page-heading h1{font-size:25px}.admin-summary{grid-template-columns:repeat(3,1fr)}}@media(max-width:540px){.top{padding:15px}.page-heading{align-items:flex-start}.page-badge{display:none}.admin-summary{gap:7px}.metric{padding:13px 10px}.metric b{font-size:22px}.metric span{font-size:11px}.card{padding:17px 14px}}

/* Full admin shell inspired by enterprise dashboards */
body{background:#f5f7f8}.top{position:fixed;z-index:30;inset:0 auto 0 0;width:268px;height:100vh;padding:22px 18px;background:linear-gradient(180deg,#075f3b,#014e32);box-shadow:4px 0 22px #0c35231c}.brand{height:100%;display:block}.brand-left{align-items:center}.brand h1{font-size:17px}.brand p{font-size:13px;color:#d5ecdf}.top-action{display:none}.wrap{width:calc(100% - 268px);max-width:none;min-height:100vh;margin:0 0 0 268px;padding:24px 28px 44px}.msg{min-height:56px;margin:0 0 24px;padding:16px 20px;border:0;border-radius:14px;background:#fff;box-shadow:0 3px 14px #173d2b0c}#adminBody{display:block}.admin-tabs{position:fixed;z-index:31;left:14px;top:132px;width:240px;min-height:250px;margin:0;padding:12px;background:transparent;box-shadow:none}.admin-tabs:before{content:'ระบบตรวจครุภัณฑ์ประจำปี';padding:8px 10px 14px;border-top:1px solid #ffffff20;color:#d6ecdf}.admin-tabs .btn{padding:14px 15px;color:#eaf6ef}.admin-tabs .active{background:#0a9552!important;color:#fff!important;box-shadow:none}.page-heading{margin-bottom:20px}.page-heading h1{font-size:26px}.page-badge{background:#f0f7f3}.admin-summary{grid-template-columns:repeat(5,minmax(150px,1fr));gap:14px}.metric{min-height:116px;padding:19px}.metric span{max-width:125px}.metric b{font-size:29px}.danger-metric b{color:#dc2626}.danger-metric:after{background:#fff0f0}.blue-metric b{color:#3559d9}.blue-metric:after{background:#eef1ff}.recent-card{margin-bottom:20px}.recent-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}.recent-head h2{margin:0;font-size:19px}.recent-head span{color:#748179;font-size:13px}.table-wrap{overflow:auto;border:1px solid #e1e8e4;border-radius:13px}table{width:100%;border-collapse:collapse;min-width:760px}th,td{padding:13px 15px;border-bottom:1px solid #e8eeea;text-align:left;font-size:13px}th{background:#f7faf8;color:#42534a;font-weight:600}tbody tr:hover{background:#f8fbf9}.asset-link{color:#078942;font-weight:500}.result-chip{display:inline-block;padding:5px 9px;border-radius:8px;background:#e7f6ed;color:#087a42;font-size:12px}@media(max-width:1100px){.admin-summary{grid-template-columns:repeat(3,1fr)}}@media(max-width:850px){.top{position:relative;width:100%;height:auto;padding:16px}.brand{height:auto}.wrap{width:100%;margin:0;padding:16px 14px 36px}.admin-tabs{position:static;display:grid;width:100%;min-height:0;margin-bottom:18px;padding:4px;background:#eaf3ee}.admin-tabs:before{display:none}.admin-summary{grid-template-columns:repeat(2,1fr)}}@media(max-width:520px){.admin-summary{grid-template-columns:1fr 1fr}.metric{min-height:96px}.recent-head{display:block}.recent-head span{display:block;margin-top:4px}}

.photo-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px}.photo-stats>div{background:#fff;border:1px solid var(--line);border-radius:16px;padding:18px 20px;display:flex;align-items:baseline;gap:10px;box-shadow:0 6px 18px #153b2510}.photo-stats b{font-size:28px;color:var(--ink)}.photo-stats span{color:var(--muted)}.photo-stats .ok{border-left:5px solid var(--green)}.photo-stats .ok b{color:var(--green)}.photo-stats .waiting{border-left:5px solid #e18a16}.photo-stats .waiting b{color:#b86708}.photo-list-head{display:flex;justify-content:space-between;align-items:end;gap:18px;margin-bottom:18px}.photo-list-head h2{margin:0 0 3px;font-size:21px}.photo-list-head p{margin:0;color:var(--muted)}.photo-search{display:flex;gap:8px;min-width:480px}.photo-search input{flex:1}.photo-search select{width:140px}.photo-loading{min-height:170px;display:grid;place-items:center;align-content:center;gap:12px;color:var(--muted)}.asset-photo-list{display:grid;gap:10px}.photo-row{display:grid;grid-template-columns:82px minmax(240px,1fr) 150px 128px;gap:16px;align-items:center;border:1px solid #e0e8e4;border-radius:14px;padding:12px;background:#fff}.photo-thumb{width:82px;height:70px;border-radius:11px;overflow:hidden;background:#edf6f1;display:grid;place-items:center;color:#6d7e74;font-size:12px;text-align:center}.photo-thumb img{width:100%;height:100%;object-fit:cover}.photo-detail b{display:block;font-size:15px;color:#087943}.photo-detail span{display:block;color:var(--ink);margin-top:3px}.photo-detail small{display:block;color:var(--muted);margin-top:3px}.photo-state{display:inline-flex;justify-content:center;border-radius:999px;padding:7px 10px;font-size:12px;background:#fff3df;color:#a95b00}.photo-state.added{background:#e4f6eb;color:#087a42}.photo-row .btn{width:100%}#photoEditorCard{border:2px solid #b9e2cc;scroll-margin-top:18px}.photo-preview{max-width:360px;max-height:270px;object-fit:contain;border-radius:14px;border:1px solid var(--line);margin:12px 0}.photo-info{margin:10px 0;padding:12px 14px;background:#f2f8f5;border-radius:11px}@media(max-width:850px){.photo-stats{grid-template-columns:1fr 1fr 1fr}.photo-list-head{display:block}.photo-search{min-width:0;margin-top:14px}.photo-row{grid-template-columns:72px 1fr 110px}.photo-thumb{width:72px;height:64px}.photo-row .btn{grid-column:2/4}}@media(max-width:560px){.photo-stats{grid-template-columns:1fr}.photo-stats>div{padding:12px 16px}.photo-search{display:grid}.photo-search select{width:100%}.photo-row{grid-template-columns:64px 1fr;padding:10px;gap:10px}.photo-thumb{width:64px;height:58px}.photo-state{justify-self:start}.photo-row .btn{grid-column:1/3}.photo-detail span{font-size:13px}}

.photo-editor-modal{position:fixed;inset:0;z-index:1100;background:rgba(13,42,28,.55);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;padding:22px}.photo-editor-modal.hidden{display:none}.photo-editor-dialog{position:relative;width:min(680px,100%);max-height:90vh;overflow:auto;background:#fff;border:1px solid #cfe1d7;border-radius:20px;padding:26px;box-shadow:0 24px 70px rgba(3,40,21,.3)}.photo-modal-close{position:absolute;right:16px;top:14px;width:36px;height:36px;border:0;border-radius:50%;background:#eef4f1;color:#33453c;font-size:25px;line-height:1;cursor:pointer}.photo-modal-actions{display:flex!important;gap:10px!important;margin-top:16px}.photo-modal-actions .btn,.photo-modal-actions .file-label{flex:1 1 0;width:auto;min-width:0;min-height:46px;display:flex;align-items:center;justify-content:center;margin:0}.btn.danger{background:#fff0f0;border:1px solid #dc2626;color:#c72020}.btn.danger:hover{background:#dc2626;color:#fff}.photo-editor-dialog .photo-preview{display:block;max-width:100%;width:auto;margin:14px auto;max-height:330px}.photo-editor-dialog .photo-preview.hidden{display:none}@media(max-width:620px){.photo-editor-modal{padding:12px}.photo-editor-dialog{padding:22px 16px 18px;border-radius:17px}.photo-modal-actions{flex-wrap:wrap}.photo-modal-actions .btn,.photo-modal-actions .file-label{flex:1 1 calc(50% - 5px);min-height:44px}}
.report-heading{display:flex;align-items:center;justify-content:space-between;margin:24px 0 12px}.report-heading h3{margin:0;font-size:19px;color:#087a42}.report-heading span{font-size:13px;color:var(--muted)}.report-table table{min-width:820px}.report-table td:first-child{font-weight:500}.report-table td:nth-child(4){white-space:nowrap}@media(max-width:560px){.report-heading{display:block}.report-heading span{display:block;margin-top:3px}}.round-create-top{min-width:170px;align-self:center}.round-create-dialog{width:min(760px,100%)}.round-modal-actions{display:flex;gap:10px;margin-top:4px}.round-modal-actions .btn{flex:1;min-height:46px}.rounds-heading{margin-bottom:18px}@media(max-width:620px){.page-heading .round-create-top{width:100%;margin-top:14px}.round-modal-actions{display:grid;grid-template-columns:1fr 1fr}.round-create-dialog .row{grid-template-columns:1fr}}.round-status-btn{min-width:92px;border:1px solid transparent!important;color:#fff!important;font-weight:600;box-shadow:0 4px 10px rgba(0,0,0,.08)}.round-status-btn.status-open{background:#079447!important}.round-status-btn.status-open:hover{background:#067a3b!important}.round-status-btn.status-closed{background:#dc3b3b!important}.round-status-btn.status-closed:hover{background:#bd2d2d!important}.round-icon{background:#eaf7f0!important}.round-icon svg{display:block}.round-status-pill{display:inline-flex!important;align-items:center;gap:4px;padding:4px 10px!important;border-radius:999px!important;font-weight:600}
/* Admin shell redesign */
body{background:#f5f7f6!important;color:#172033!important}.top{position:fixed!important;inset:0 auto 0 0!important;width:248px!important;height:100vh!important;padding:20px 16px!important;background:linear-gradient(165deg,#075f3d 0%,#006b43 58%,#005637 100%)!important;z-index:30!important;box-shadow:none!important}.brand{display:block!important;max-width:none!important}.brand-left{display:flex!important;align-items:center!important;gap:12px!important;padding:0 4px 18px!important}.logo{width:50px!important;height:50px!important;flex:0 0 50px!important;font-size:16px!important;background:#ffffff12!important;border:2px solid #fff!important}.brand h1{font-size:17px!important;line-height:1.35!important;color:#fff!important}.brand p{font-size:13px!important;color:#d9f2e6!important}.sidebar-caption{padding:18px 8px 12px;border-top:1px solid #ffffff25;color:#d7eee3;font-size:13px;font-weight:500}.admin-topbar{position:fixed;top:0;left:248px;right:0;height:72px;background:#fff;border-bottom:1px solid #e5ebe8;display:flex;align-items:center;justify-content:space-between;padding:0 28px;z-index:25;box-shadow:0 2px 10px #10251a0a}.topbar-left{display:flex;align-items:center;gap:18px}.hamburger{font-size:22px;color:#263a30}.admin-topbar h2{font-size:21px;margin:0}.top-user{display:flex;align-items:center;gap:10px}.top-user b,.top-user small{display:block}.top-user b{font-size:13px}.top-user small{font-size:11px;color:#748179}.user-avatar{width:38px;height:38px;border-radius:50%;display:grid;place-items:center;background:#e7ecea;color:#7a8580;font-size:18px}.wrap{width:auto!important;max-width:none!important;min-height:100vh!important;margin:0 0 0 248px!important;padding:96px 28px 44px!important}.admin-tabs{position:fixed!important;left:14px!important;top:151px!important;width:220px!important;display:grid!important;gap:7px!important;padding:0!important;margin:0!important;background:transparent!important;border:0!important;border-radius:0!important;box-shadow:none!important;z-index:35!important}.admin-tabs .btn{width:100%!important;min-width:0!important;min-height:48px!important;justify-content:flex-start!important;display:flex!important;align-items:center!important;gap:13px!important;padding:11px 16px!important;border:0!important;border-radius:10px!important;background:transparent!important;color:#e5f4ec!important;font-size:14px!important;font-weight:500!important;box-shadow:none!important}.admin-tabs .btn:hover{background:#ffffff12!important}.admin-tabs .btn.active{background:#07954d!important;color:#fff!important}.nav-icon{width:25px;display:inline-grid;place-items:center;font-size:22px;line-height:1}.msg{border-radius:10px!important}.page-heading{margin-bottom:20px!important}.page-heading h1{font-size:25px!important;color:#172033!important}.page-heading p{color:#718078!important}.page-badge{border-radius:9px!important;background:#edf7f1!important;color:#087a42!important}.card{border:1px solid #e0e7e3!important;border-radius:14px!important;box-shadow:0 4px 15px #173c280b!important;padding:22px!important}.admin-summary{gap:14px!important}.metric{border:1px solid #e1e8e4!important;border-radius:13px!important;box-shadow:0 3px 12px #10291c0a!important;background:#fff!important}.metric b{color:#111a2a!important}.recent-card{margin-top:8px}.table-wrap{border-radius:11px!important}th{background:#f8faf9!important}.round{border:1px solid #dce5e0!important;border-radius:13px!important;padding:18px 20px!important;box-shadow:none!important;background:#fff!important}.round:hover{border-color:#b7d8c6!important;box-shadow:0 5px 15px #0c4b2b0b!important}.round-icon{width:58px!important;height:58px!important;flex:0 0 58px!important;border-radius:50%!important;background:#e7f6ee!important}.round-icon svg{width:28px;height:28px}.round-main>b{font-size:17px!important;color:#172033!important}.round-meta{margin-top:4px;color:#68776f!important}.round-actions{margin-top:11px!important;gap:8px!important}.round-toggle-btn{min-width:76px!important;padding:9px 15px!important;border:0!important;border-radius:9px!important;box-shadow:none!important;font-weight:600!important}.round-toggle-btn.toggle-open{background:#e4f6eb!important;color:#087b42!important}.round-toggle-btn.toggle-close{background:#fde8e8!important;color:#c52c2c!important}.round-actions .secondary{border:0!important;background:#eef2f0!important;color:#33443b!important;border-radius:9px!important}.round-status-pill{font-size:11px!important;padding:3px 8px!important}.round-status-pill.status-open{background:#e3f6eb!important;color:#087b42!important}.round-status-pill.status-closed{background:#fde8e8!important;color:#c52c2c!important}.kebab{color:#68796f!important}.btn.primary{background:#078d47!important;border-radius:9px!important;box-shadow:none!important}.photo-stats>div{border-radius:13px!important;box-shadow:none!important}.asset-photo-list .photo-row{border-radius:12px!important}.action-box,.photo-editor-dialog{border-radius:16px!important}.section-title .step{box-shadow:none!important}
@media(max-width:850px){.top{position:relative!important;width:100%!important;height:auto!important;padding:14px 16px!important}.brand-left{padding:0!important}.sidebar-caption{display:none}.admin-topbar{position:relative!important;left:0!important;height:60px!important;padding:0 16px!important}.top-user{display:none}.wrap{margin:0!important;padding:16px 14px 38px!important}.admin-tabs{position:static!important;width:100%!important;display:flex!important;overflow-x:auto!important;margin-bottom:18px!important;padding:5px!important;background:#eaf3ee!important;border-radius:12px!important}.admin-tabs .btn{color:#53635b!important;justify-content:center!important;white-space:nowrap!important}.admin-tabs .btn.active{color:#fff!important}.nav-icon{font-size:18px}.admin-summary{grid-template-columns:repeat(2,1fr)!important}.page-heading{display:block!important}}

.nav-icon svg{width:22px;height:22px;display:block}.home-heading{margin-bottom:18px!important}.home-heading h1{font-size:22px!important;margin-bottom:4px!important}.admin-summary{display:grid!important;grid-template-columns:repeat(5,minmax(0,1fr))!important;gap:14px!important}.metric{position:relative!important;display:flex!important;align-items:flex-start!important;gap:13px!important;min-height:150px!important;padding:20px 17px!important}.metric:after{display:none!important}.metric-icon{width:52px;height:52px;flex:0 0 52px;border-radius:50%;display:grid;place-items:center}.metric-icon svg{width:28px;height:28px}.metric-copy{min-width:0;padding-top:2px}.metric-copy>span{display:block!important;font-size:13px!important;line-height:1.35!important;color:#34423b!important;font-weight:500!important;max-width:none!important;min-height:35px}.metric-value{display:flex;align-items:baseline;gap:7px;margin-top:3px;white-space:nowrap}.metric-value b{font-size:27px!important;line-height:1!important;color:#111a2a!important}.metric-value em{font-size:11px;color:#69786f;font-style:normal}.metric-copy small{display:block;margin-top:19px;color:#819087;font-size:11px;white-space:nowrap}.metric-round .metric-icon{background:#e4f5eb;color:#078942}.metric-assets .metric-icon{background:#e4efff;color:#2467d9}.metric-inspected .metric-icon{background:#fff0dc;color:#f07a00}.metric-damaged .metric-icon{background:#fde8e8;color:#df2d2d}.metric-moved .metric-icon{background:#ecebff;color:#4b4fdb}.metric-damaged .metric-copy small{color:#df2d2d}.metric-moved .metric-copy small{color:#4b4fdb}.recent-card{margin-top:6px!important}.sidebar-caption{font-size:14px!important;color:#fff!important;font-weight:600!important}.round-icon svg{display:block!important;width:30px!important;height:30px!important}@media(max-width:1200px){.admin-summary{grid-template-columns:repeat(3,1fr)!important}}@media(max-width:850px){.admin-summary{grid-template-columns:repeat(2,1fr)!important}.metric{min-height:130px!important}.metric-copy small{margin-top:12px}}@media(max-width:480px){.admin-summary{grid-template-columns:1fr!important}}

/* Final topbar and date picker */
.admin-topbar{position:fixed!important;top:0!important;left:0!important;right:0!important;height:78px!important;background:linear-gradient(110deg,#075f3d,#007347)!important;border:0!important;padding:0 24px!important;color:#fff!important;z-index:50!important;box-shadow:0 3px 14px #003c2426!important}.admin-topbar .brand{display:block!important;margin:0!important;max-width:none!important}.admin-topbar .brand-left{display:flex!important;align-items:center!important;gap:12px!important;padding:0!important}.admin-topbar .logo{width:50px!important;height:50px!important;flex:0 0 50px!important;background:#ffffff12!important;border:2px solid #fff!important;color:#fff!important}.admin-topbar h1{font-size:17px!important;color:#fff!important;margin:0!important}.admin-topbar p{font-size:13px!important;color:#d9f2e6!important;margin:2px 0 0!important}.top-user-icon{width:44px;height:44px;border:1px solid #ffffff66;border-radius:50%;display:grid;place-items:center;background:#ffffff12;color:#fff;cursor:pointer}.top-user-icon svg{width:24px;height:24px}.top{top:78px!important;height:calc(100vh - 78px)!important;padding:0 16px!important}.sidebar-caption{padding:21px 8px 16px!important;margin:0!important;border-top:0!important;border-bottom:1px solid #ffffff25!important;font-size:15px!important}.admin-tabs{top:151px!important}.admin-tabs:before{display:none!important;content:none!important;border:0!important}.wrap{padding-top:102px!important}.flatpickr-calendar{font-family:"Kanit",sans-serif!important;border:1px solid #d8e5de!important;border-radius:14px!important;box-shadow:0 16px 38px #0b382333!important;overflow:hidden}.flatpickr-months{background:#078942;color:#fff;padding:7px 6px 3px}.flatpickr-months .flatpickr-month{color:#fff;fill:#fff;height:40px}.flatpickr-current-month,.flatpickr-current-month input.cur-year,.flatpickr-current-month .flatpickr-monthDropdown-months{color:#fff!important}.flatpickr-current-month .flatpickr-monthDropdown-months{background:#078942!important}.flatpickr-prev-month,.flatpickr-next-month{color:#fff!important;fill:#fff!important;top:7px!important}.flatpickr-prev-month svg,.flatpickr-next-month svg{fill:#fff!important}.flatpickr-weekdays{background:#e9f7ef}.flatpickr-weekday{color:#087a42!important;font-weight:600!important}.flatpickr-day{border-radius:8px!important}.flatpickr-day:hover{background:#e5f5ec!important;border-color:#e5f5ec!important}.flatpickr-day.selected,.flatpickr-day.startRange,.flatpickr-day.endRange{background:#078942!important;border-color:#078942!important;color:#fff!important}.flatpickr-day.today{border-color:#078942!important;color:#078942}.flatpickr-day.today.selected{color:#fff}.flatpickr-input.form-control,.date-picker+.form-control{background:#fff!important}
@media(max-width:850px){.admin-topbar{position:relative!important;height:72px!important;padding:0 14px!important}.admin-topbar .logo{width:44px!important;height:44px!important;flex-basis:44px!important}.admin-topbar h1{font-size:14px!important}.admin-topbar p{font-size:11px!important}.top{position:relative!important;top:0!important;width:100%!important;height:auto!important;padding:0 14px!important}.sidebar-caption{padding:13px 5px!important}.wrap{padding-top:16px!important}.admin-tabs{position:static!important}}
.top-account-area{display:flex;align-items:center;gap:16px}.notify-button,.account-menu{border:0;background:transparent;color:#fff;cursor:pointer;font-family:"Kanit",sans-serif}.notify-button{position:relative;width:38px;height:38px;display:grid;place-items:center;padding:0}.notify-button>svg{width:23px;height:23px}.notify-badge{position:absolute;right:1px;top:0;min-width:17px;height:17px;padding:0 4px;border-radius:999px;background:#ef3434;color:#fff;border:2px solid #087044;font:600 10px/13px "Kanit",sans-serif;text-align:center}.account-menu{display:flex;align-items:center;gap:10px;padding:5px 7px;border-radius:10px}.account-menu:hover{background:#ffffff12}.account-avatar{width:38px;height:38px;border-radius:50%;display:grid;place-items:center;background:#e8edef;color:#78858d}.account-avatar svg{width:26px;height:26px}.account-copy{text-align:left;min-width:84px}.account-copy b,.account-copy small{display:block}.account-copy b{font-size:13px;line-height:1.3;color:#fff}.account-copy small{font-size:10px;color:#d8eee3;margin-top:1px}.account-chevron{width:18px;height:18px;color:#d8eee3}@media(max-width:560px){.top-account-area{gap:5px}.account-copy,.account-chevron{display:none}.notify-button{width:34px}.account-menu{padding:3px}.account-avatar{width:36px;height:36px}}
  .admin-topbar .brand,.admin-topbar .brand-left{height:100%;display:flex!important;align-items:center!important}
  .admin-topbar .brand-left>div{display:flex;flex-direction:column;justify-content:center}
  .admin-topbar .brand-left h1,.admin-topbar .brand-left p{margin-top:0;margin-bottom:0}
  .round{display:grid!important;grid-template-columns:64px minmax(0,1fr) auto!important;align-items:center!important;gap:16px!important}
  .round-main{min-width:0}
  .round-actions{margin:0!important;display:flex!important;align-items:center!important;justify-content:flex-end!important;gap:8px!important;white-space:nowrap}
  .round-delete-btn{width:42px!important;height:42px!important;min-width:42px!important;padding:0!important;display:grid!important;place-items:center!important;background:#fff0f0!important;color:#c52c2c!important;border:1px solid #f6caca!important;border-radius:10px!important}
  .round-delete-btn svg{width:20px;height:20px}
  .round-delete-btn:hover{background:#fde0e0!important}
  .kebab{display:none!important}
  @media(max-width:700px){
    .round{grid-template-columns:54px minmax(0,1fr)!important}
    .round-actions{grid-column:1/-1!important;justify-content:flex-end!important}
  }

  .admin-topbar{background:linear-gradient(112deg,#00683f 0%,#007c4c 60%,#00643d 100%)!important;border-bottom:0!important;color:#fff!important;box-shadow:0 4px 16px #003f2830!important}
  .admin-topbar .brand-left h1,.admin-topbar .brand-left p,.admin-topbar .notify-button,.admin-topbar .account-menu,.admin-topbar .account-copy small{color:#fff!important}
  .account-avatar{background:#edf2f0!important;color:#65756e!important}
  .top{background:linear-gradient(180deg,#f8fcfa 0%,#eff8f3 100%)!important;color:#123829!important;border-right:1px solid #d9e9e1!important}
  .sidebar-caption{color:#0b5034!important;border-color:#d4e7dd!important}
  .admin-tabs{top:126px!important;gap:2px!important}
  .admin-tabs .btn{min-height:40px!important;height:40px!important;padding:7px 13px!important;gap:11px!important;border-radius:9px!important;font-size:14px!important}
  .admin-tabs .btn:not(.active){color:#174c38!important}
  .admin-tabs .btn.active{background:#dff3e8!important;color:#087948!important;box-shadow:none!important}
  .admin-tabs .nav-icon{width:22px!important}
  .admin-tabs .nav-icon svg{width:20px!important;height:20px!important}
  .top-account-area{position:relative}
  .notification-panel{position:absolute;right:210px;top:58px;width:390px;max-height:480px;background:#fff;color:#172033;border:1px solid #dce7e1;border-radius:14px;box-shadow:0 18px 46px #061b1260;overflow:hidden;z-index:120}
  .notification-panel.hidden{display:none!important}
  .notification-head{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid #e5ece8;background:#f7fbf9}
  .notification-head b{font-size:16px;color:#123829}
  .notification-head button{border:0;background:transparent;color:#078648;font-weight:600;cursor:pointer}
  .notification-list{max-height:410px;overflow:auto}
  .notification-item{padding:13px 16px;border-bottom:1px solid #edf1ef;display:grid;gap:3px}
  .notification-item:last-child{border-bottom:0}
  .notification-item.unread{background:#effaf4}
  .notification-item strong{font-size:14px;color:#153326}
  .notification-item small{color:#6e7e76}
  .notification-meta{display:flex;gap:8px;flex-wrap:wrap;font-size:12px;color:#52645b}
  .notification-result{font-weight:600;color:#078648}
  .notification-empty{padding:28px 18px;text-align:center;color:#748179}
  @media(max-width:700px){.notification-panel{position:fixed;left:12px;right:12px;top:70px;width:auto;max-height:70vh}}

  .admin-tabs{top:151px!important;height:auto!important;min-height:0!important;grid-template-rows:repeat(3,40px)!important;grid-auto-rows:40px!important;align-content:start!important;row-gap:6px!important}
  .admin-tabs>.btn{height:40px!important;min-height:40px!important;margin:0!important;align-self:stretch!important}

  .admin-tabs .btn{font-size:15.5px!important;font-weight:600!important}
  .rounds-heading,.dash .section-title{padding-left:0!important}
  .rounds-heading>.step,.dash .section-title>.step{display:none!important}

  .round-meta{margin-top:5px!important;line-height:1.45!important}
  .round-date{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .round-sheet{margin-top:2px;color:#5d7167;font-size:14px;font-weight:500}
  .round-report-btn{background:#e6f5ec!important;color:#087b42!important;border:1px solid #c8e8d6!important}
  .round-report-btn:hover{background:#d8efe2!important}

  .round-actions .btn{height:42px!important;min-height:42px!important}
  .round-report-btn,.round-edit-btn{width:42px!important;min-width:42px!important;padding:0!important;display:grid!important;place-items:center!important;border-radius:10px!important}
  .round-report-btn{background:#eef3ff!important;color:#315fc6!important;border:1px solid #d6e1fb!important}
  .round-report-btn:hover{background:#e1eaff!important}
  .round-edit-btn{background:#fff7e8!important;color:#b56a00!important;border:1px solid #f4dfb7!important}
  .round-edit-btn:hover{background:#ffefd1!important}
  .round-report-btn svg,.round-edit-btn svg{width:20px;height:20px}

/* ===== Unified premium modal system ===== */
.action-modal,.photo-editor-modal{background:rgba(15,31,25,.72)!important;backdrop-filter:blur(8px) saturate(.78)!important;-webkit-backdrop-filter:blur(8px) saturate(.78)!important;padding:22px!important}
.action-box,.photo-editor-dialog{position:relative;isolation:isolate;overflow:hidden;border:1px solid #9acdb1!important;border-radius:32px!important;background:linear-gradient(180deg,#fff 0%,#fff 72%,#f2fbf6 100%)!important;box-shadow:0 30px 90px rgba(3,26,17,.38),inset 0 1px 0 #fff!important}
.action-box{width:min(620px,100%)!important;min-height:440px;padding:58px 48px 42px!important;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#075c36!important}
.action-box:before,.photo-editor-dialog:before{content:"";position:absolute;z-index:-1;left:-8%;right:-8%;bottom:-74px;height:150px;border-radius:50% 50% 0 0/46% 46% 0 0;background:linear-gradient(135deg,rgba(18,169,91,.05),rgba(40,196,115,.17));transform:rotate(2deg)}
.action-box:after{content:"";position:absolute;z-index:-1;left:-10%;right:-4%;bottom:-106px;height:170px;border-radius:50% 50% 0 0/55% 58% 0 0;background:rgba(20,174,91,.08);transform:rotate(-4deg)}
.action-content{width:100%;text-align:center}.modal-status-icon{position:relative;display:grid;place-items:center;width:132px;height:132px;margin:0 auto 24px;border:14px solid #dff6e9;border-radius:50%;background:linear-gradient(145deg,#13b963,#078946);color:#fff;box-shadow:0 0 0 10px #edf9f2,0 14px 28px rgba(5,131,69,.22);font-size:68px;font-weight:800;line-height:1}.modal-status-icon:after{content:"";position:absolute;inset:10px;border-radius:50%;border:2px solid rgba(255,255,255,.2)}.action-box h3{margin:0 0 12px;color:#075c36;font-size:34px;font-weight:700;line-height:1.25}.action-box p{max-width:500px;margin:0 auto 26px;color:#4b5563;font-size:19px;line-height:1.7}.action-box.error{color:#a62424!important}.action-box.error .modal-status-icon{border-color:#fee2e2;background:linear-gradient(145deg,#f05252,#c92727);box-shadow:0 0 0 10px #fff1f1,0 14px 28px rgba(190,31,31,.18)}.action-box.error h3{color:#a62424}.action-box.loading .modal-status-icon{border-color:#ddf3e7;background:#f3fbf7;box-shadow:0 0 0 10px #eef9f3}.action-box.loading .modal-status-icon:before{content:"";width:68px;height:68px;border:7px solid #ccebd9;border-top-color:#078942;border-radius:50%;animation:spin .8s linear infinite}.action-box.loading .modal-status-icon:after{display:none}.action-box .spinner{display:none}.action-box>.btn{position:relative;z-index:2;min-width:210px;min-height:54px;margin-top:0;border-radius:14px!important;background:linear-gradient(135deg,#07964b,#006f39)!important;box-shadow:0 10px 22px rgba(4,119,59,.2)!important;font-size:18px}.action-box>.btn:after{content:"›";margin-left:15px;font-size:28px;line-height:0;vertical-align:-2px}.action-box .photo-actions{position:relative;z-index:2;margin-top:22px!important}.action-box .photo-actions .btn{min-width:130px;min-height:48px;border-radius:12px!important}
.photo-editor-dialog{width:min(760px,100%)!important;max-height:min(90vh,860px)!important;padding:34px 36px 30px!important}.photo-modal-close{top:18px!important;right:18px!important;width:44px!important;height:44px!important;border:2px solid #d5ddd9!important;background:#fff!important;color:#58625d!important;font-size:29px!important;box-shadow:0 4px 12px rgba(20,45,32,.08)}.photo-editor-dialog .section-title{padding-right:52px;margin-bottom:24px}.photo-editor-dialog .section-title .step{width:58px;height:58px;flex-basis:58px;background:linear-gradient(145deg,#16ad5a,#067b40);box-shadow:0 7px 16px rgba(7,137,66,.2)!important}.photo-editor-dialog .section-title h2{font-size:26px!important}.photo-editor-dialog input,.photo-editor-dialog select{min-height:52px}.photo-modal-actions,.round-modal-actions{position:relative;z-index:2;gap:12px!important;margin-top:22px!important}.photo-modal-actions .btn,.photo-modal-actions .file-label,.round-modal-actions .btn{min-height:50px!important;border-radius:12px!important}
@media(max-width:620px){.action-modal,.photo-editor-modal{padding:12px!important}.action-box{min-height:390px;padding:44px 22px 34px!important;border-radius:25px!important}.modal-status-icon{width:104px;height:104px;border-width:11px;margin-bottom:22px;font-size:54px}.action-box h3{font-size:26px}.action-box p{font-size:16px;line-height:1.6}.action-box>.btn{width:100%;min-width:0}.photo-editor-dialog{padding:26px 17px 20px!important;border-radius:25px!important}.photo-modal-close{top:12px!important;right:12px!important;width:40px!important;height:40px!important}.photo-editor-dialog .section-title{gap:12px;padding-right:38px}.photo-editor-dialog .section-title .step{width:46px;height:46px;flex-basis:46px}.photo-editor-dialog .section-title h2{font-size:21px!important}.action-box .photo-actions{display:grid!important;grid-template-columns:1fr 1fr;width:100%}.action-box .photo-actions .btn{min-width:0}}
</style></head><body><aside class="top"><div class="sidebar-caption">ระบบตรวจครุภัณฑ์</div></aside><header class="admin-topbar"><div class="brand"><div class="brand-left"><div><h1>กองบริหารทรัพยากรบุคคล</h1><p>กลุ่มบริหารทั่วไป</p></div></div></div><div class="top-account-area"><button id="notifyButton" class="notify-button" type="button" aria-label="การแจ้งเตือน" onclick="toggleNotifications()"><svg viewBox="0 0 24 24" fill="none"><path d="M6 9a6 6 0 0 1 12 0c0 6 2.5 7 2.5 7h-17S6 15 6 9Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M10 20h4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg><span id="notifyBadge" class="notify-badge" style="display:none">0</span></button><div id="notificationPanel" class="notification-panel hidden"><div class="notification-head"><b>การแจ้งเตือน</b><button type="button" onclick="markNotificationsRead()">อ่านทั้งหมดแล้ว</button></div><div id="notificationList" class="notification-list"><div class="notification-empty">กำลังโหลดข้อมูล…</div></div></div><button class="account-menu" type="button" aria-label="เมนูผู้ดูแลระบบ"><span class="account-avatar"><svg viewBox="0 0 24 24" fill="none"><circle cx="12" cy="8" r="4" fill="currentColor"/><path d="M4.5 21c.7-4.2 3.2-6.5 7.5-6.5s6.8 2.3 7.5 6.5" fill="currentColor"/></svg></span><span class="account-copy"><b>ผู้ดูแลระบบ</b><small>Administrator</small></span><svg class="account-chevron" viewBox="0 0 24 24" fill="none"><path d="m9 10 3 3 3-3" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg></button></div></header><main class="wrap">${adminHtml_()}</main><script src="https://cdn.jsdelivr.net/npm/flatpickr@4.6.13"></script><script src="https://cdn.jsdelivr.net/npm/flatpickr@4.6.13/dist/l10n/th.js"></script><script>${adminJs_()}</script></body></html>`;
}

function includeComponent_(name){return HtmlService.createHtmlOutputFromFile(name).getContent()}
function adminHtml_(){return `
  <div id="adminMsg" class="msg">กำลังตรวจสอบสิทธิ์แอดมิน…</div>
  <div id="actionModal" class="action-modal hidden"><div id="actionBox" class="action-box"><div id="actionContent"></div><button id="actionClose" type="button" class="btn primary hidden" onclick="closeActionModal()">ตกลง</button></div></div>
  <div id="adminBody" class="hidden">
    <div class="admin-tabs"><button id="homeTab" class="btn active" onclick="showAdminPanel('home')"><span class="nav-icon"><svg viewBox="0 0 24 24" fill="none"><path d="M3 10.7 12 3l9 7.7V21h-6.2v-6.2H9.2V21H3V10.7Z" fill="currentColor"/><path d="M5.5 9V4.8h3V6.4" fill="currentColor"/></svg></span><span>หน้าหลัก</span></button><button id="inspectTab" class="btn secondary" onclick="showAdminPanel('inspect')"><span class="nav-icon"><svg viewBox="0 0 24 24" fill="none"><rect x="3" y="5" width="18" height="16" rx="2" stroke="currentColor" stroke-width="1.8"/><path d="M7 3v4M17 3v4M3 10h18" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg></span><span>ตรวจครุภัณฑ์</span></button><button id="photoTab" class="btn secondary" onclick="showAdminPanel('photo')"><span class="nav-icon"><svg viewBox="0 0 24 24" fill="none"><rect x="3" y="4" width="18" height="16" rx="2" stroke="currentColor" stroke-width="1.8"/><circle cx="9" cy="10" r="2" stroke="currentColor" stroke-width="1.8"/><path d="m5 18 5-5 3 3 2-2 4 4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg></span><span>เพิ่มรูปครุภัณฑ์</span></button></div>
`+includeComponent_('AdminHome')+includeComponent_('AdminRounds')+includeComponent_('AdminPhotos')+`  </div>`}

function adminJs_(){return `
  const el=id=>document.getElementById(id);const err=e=>e&&e.message?e.message:String(e);let boot=null; function actionLoading(text){el('actionBox').className='action-box loading';el('actionContent').className='action-content';el('actionContent').innerHTML='<div class="modal-status-icon" aria-hidden="true"></div><h3>'+text+'</h3><p>กรุณารอสักครู่ ระบบกำลังดำเนินการ</p>';el('actionClose').classList.add('hidden');el('actionModal').classList.remove('hidden')} function actionResult(title,text,isError){el('actionBox').className='action-box'+(isError?' error':' success');el('actionContent').className='action-content';el('actionContent').innerHTML='<div class="modal-status-icon" aria-hidden="true">'+(isError?'!':'✓')+'</div><h3>'+title+'</h3><p>'+text+'</p>';el('actionClose').classList.remove('hidden');el('actionModal').classList.remove('hidden')} function closeActionModal(){el('actionModal').classList.add('hidden')}
  let notificationTimer=null;
  let notificationItems=[];
  let notificationLoading=false;
  function startNotificationPolling(){
    loadNotifications(true);
    if(notificationTimer)clearInterval(notificationTimer);
    notificationTimer=setInterval(()=>{if(!document.hidden)loadNotifications(false)},90000);
  }
  function loadNotifications(initial){
    if(notificationLoading)return;
    notificationLoading=true;
    google.script.run.withSuccessHandler(x=>{
      notificationLoading=false;
      notificationItems=(x&&x.items)||[];
      const newest=notificationItems[0]&&notificationItems[0].id||'';
      let seen=localStorage.getItem('assetAdminLastSeenNotification')||'';
      if(initial&&!seen&&newest){localStorage.setItem('assetAdminLastSeenNotification',newest);seen=newest}
      let unread=0;
      for(const item of notificationItems){if(item.id===seen)break;unread++}
      const badge=el('notifyBadge');
      badge.textContent=unread>99?'99+':String(unread);
      badge.style.display=unread?'grid':'none';
      renderNotifications(seen);
    }).withFailureHandler(()=>{notificationLoading=false}).getAdminNotifications();
  }
  function renderNotifications(seen){
    const list=el('notificationList');
    if(!notificationItems.length){list.innerHTML='<div class="notification-empty">ยังไม่มีรายการบันทึกใหม่</div>';return}
    let passedSeen=false;
    list.innerHTML=notificationItems.map(item=>{
      if(item.id===seen)passedSeen=true;
      return '<div class="notification-item '+(!passedSeen&&item.id!==seen?'unread':'')+'"><strong>'+esc(item.assetNo)+' · '+esc(item.name)+'</strong><div class="notification-meta"><span>'+esc(item.date)+'</span><span>ผู้ตรวจ: '+esc(item.inspector)+'</span></div><div class="notification-meta"><span class="notification-result">'+esc(item.result)+'</span><span>'+esc(item.roundName)+'</span></div></div>'
    }).join('');
  }
  function toggleNotifications(){
    const panel=el('notificationPanel');
    const opening=panel.classList.contains('hidden');
    panel.classList.toggle('hidden');
    if(opening){loadNotifications(false);setTimeout(markNotificationsRead,500)}
  }
  function markNotificationsRead(){
    const newest=notificationItems[0]&&notificationItems[0].id||'';
    if(newest)localStorage.setItem('assetAdminLastSeenNotification',newest);
    const badge=el('notifyBadge');badge.textContent='0';badge.style.display='none';
    renderNotifications(newest);
  }
  function loadAdmin(){google.script.run.withSuccessHandler(x=>{boot=x;el('adminMsg').classList.add('hidden');el('adminBody').classList.remove('hidden');renderRounds();updateMetrics();startNotificationPolling()}).withFailureHandler(e=>{el('adminMsg').classList.remove('hidden');el('adminMsg').textContent=err(e);el('adminMsg').classList.add('error')}).getAdminBootstrap()}
  function updateMetrics(){const rounds=(boot&&boot.rounds)||[];const home=(boot&&boot.home)||{};el('metricAll').textContent=rounds.length;el('metricAssets').textContent=Number(home.totalAssets||0).toLocaleString('th-TH');el('metricInspected').textContent=Number(home.inspected||0).toLocaleString('th-TH');el('metricDamaged').textContent=Number(home.damaged||0).toLocaleString('th-TH');el('metricMoved').textContent=Number(home.moved||0).toLocaleString('th-TH');const body=el('recentBody');const rows=home.recent||[];body.innerHTML=rows.length?rows.map(r=>'<tr><td>'+esc(r.date)+'</td><td>'+esc(r.name)+'</td><td class="asset-link">'+esc(r.assetNo)+'</td><td>'+esc(r.inspector)+'</td><td><span class="result-chip">'+esc(r.result)+'</span></td></tr>').join(''):'<tr><td colspan="5" class="empty">ยังไม่มีข้อมูลการตรวจ</td></tr>'}
  function esc(v){return String(v==null?'':v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
  function openRoundModal(){
    window.editingRoundId=null;
    el('roundModalTitle').textContent='สร้างรอบตรวจใหม่';
    el('roundModalSub').textContent='กรอกข้อมูลเพื่อสร้างรอบการตรวจครุภัณฑ์';
    el('roundSaveButton').textContent='สร้างรอบตรวจ';
    el('roundName').value='';el('start').value='';el('end').value='';el('roundCode').value='';el('roundStatus').value='ปิด';
    if(el('start')._flatpickr)el('start')._flatpickr.clear();
    if(el('end')._flatpickr)el('end')._flatpickr.clear();
    el('roundStatus').closest('.field').style.display='';
    el('roundCreateModal').classList.remove('hidden');document.body.style.overflow='hidden';setTimeout(()=>el('roundName').focus(),100)
  }
  function editRoundModal(r){
    window.editingRoundId=String(r.id);
    el('roundModalTitle').textContent='แก้ไขรอบตรวจ';
    el('roundModalSub').textContent='แก้ไขชื่อรอบ วันที่ตรวจ และรหัสเข้ารอบ';
    el('roundSaveButton').textContent='บันทึกการแก้ไข';
    el('roundName').value=r.name||'';
    el('roundCode').value=r.accessCode||'';
    if(el('start')._flatpickr)el('start')._flatpickr.setDate(new Date(Number(r.startMs)),true);else el('start').value=r.start||'';
    if(el('end')._flatpickr)el('end')._flatpickr.setDate(new Date(Number(r.endMs)),true);else el('end').value=r.end||'';
    el('roundStatus').closest('.field').style.display='none';
    el('roundCreateModal').classList.remove('hidden');document.body.style.overflow='hidden';setTimeout(()=>el('roundName').focus(),100)
  }
  function saveRoundForm(){if(window.editingRoundId)return updateRound();createRound()}
  function closeRoundModal(){el('roundCreateModal').classList.add('hidden');document.body.style.overflow=''}
  function createRound(){const f={name:el('roundName').value.trim(),start:el('start').value,end:el('end').value,accessCode:el('roundCode').value.trim(),status:el('roundStatus').value};if(!f.name)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาระบุชื่อรอบตรวจ',true);if(!f.start)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาเลือกวันเริ่มตรวจ',true);if(!f.end)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาเลือกวันสิ้นสุด',true);if(new Date(f.end)<new Date(f.start))return actionResult('วันที่ไม่ถูกต้อง','วันสิ้นสุดต้องไม่น้อยกว่าวันเริ่ม',true);actionLoading('กำลังสร้างรอบตรวจ');google.script.run.withSuccessHandler(x=>{closeRoundModal();el('roundName').value='';el('start').value='';el('end').value='';el('roundCode').value='';el('roundStatus').value='ปิด';actionResult('สร้างรอบตรวจสำเร็จ','สร้างชีท '+x.sheetName+' เรียบร้อยแล้ว',false);loadAdmin()}).withFailureHandler(e=>actionResult('สร้างรอบตรวจไม่สำเร็จ',err(e),true)).createRoundAdmin(f)}
  function updateRound(){
    const f={id:window.editingRoundId,name:el('roundName').value.trim(),start:el('start').value,end:el('end').value,accessCode:el('roundCode').value.trim()};
    if(!f.name)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาระบุชื่อรอบตรวจ',true);
    if(!f.start)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาเลือกวันเริ่มตรวจ',true);
    if(!f.end)return actionResult('กรอกข้อมูลไม่ครบ','กรุณาเลือกวันสิ้นสุด',true);
    if(new Date(f.end)<new Date(f.start))return actionResult('วันที่ไม่ถูกต้อง','วันสิ้นสุดต้องไม่น้อยกว่าวันเริ่ม',true);
    actionLoading('กำลังบันทึกการแก้ไข');
    google.script.run.withSuccessHandler(()=>{closeRoundModal();window.editingRoundId=null;actionResult('แก้ไขรอบตรวจสำเร็จ','บันทึกข้อมูลรอบตรวจเรียบร้อยแล้ว',false);loadAdmin()}).withFailureHandler(e=>actionResult('แก้ไขรอบตรวจไม่สำเร็จ',err(e),true)).updateRoundAdmin(f)
  }
  function renderRounds(){
    const box=el('rounds');
    box.innerHTML='';
    if(!boot.rounds.length){box.innerHTML='<div class="empty">ยังไม่มีรอบตรวจ</div>';return}
    boot.rounds.forEach(r=>{
      const d=document.createElement('div');
      d.className='round';
      d.innerHTML='<div class="round-icon"><svg width="28" height="28" viewBox="0 0 28 28" fill="none" aria-hidden="true"><rect x="4" y="6" width="20" height="18" rx="3" stroke="#078942" stroke-width="2"/><path d="M4 11H24M9 3V8M19 3V8" stroke="#078942" stroke-width="2" stroke-linecap="round"/><path d="M9 15H12M16 15H19M9 19H12M16 19H19" stroke="#078942" stroke-width="2" stroke-linecap="round"/></svg></div><div class="round-main"><b>'+r.name+'</b><div class="round-meta"><div class="round-date">'+r.start+' – '+r.end+' <span class="pill round-status-pill '+(r.status==='เปิด'?'status-open':'status-closed')+'" style="'+(r.status==='เปิด'?'background:#e3f6eb!important;color:#087b42!important':'background:#fde8e8!important;color:#c52c2c!important')+'">● '+r.status+'</span></div><div class="round-sheet">'+r.sheetName+'</div></div></div><div class="round-actions"></div>';
      const actions=d.querySelector('.round-actions');
      const isOpen=r.status==='เปิด';
      const target=isOpen?'ปิด':'เปิด';
      const toggle=document.createElement('button');
      toggle.type='button';
      toggle.className='btn round-toggle-btn '+(isOpen?'toggle-close':'toggle-open');
      toggle.textContent=(isOpen?'ปิดรอบ':'เปิดรอบ');
      toggle.onclick=()=>setStatus(r.id,target);
      actions.appendChild(toggle);
      const report=document.createElement('button');
      report.type='button';
      report.className='btn round-report-btn';
      report.title='รายงานผลการตรวจ';report.setAttribute('aria-label','รายงานผลการตรวจ '+r.name);report.innerHTML='<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M5 20V10M12 20V4M19 20v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M3 20h18" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>';
      report.onclick=()=>dashboard(r.id);
      actions.appendChild(report);
      const edit=document.createElement('button');
      edit.type='button';
      edit.className='btn round-edit-btn';
      edit.title='แก้ไขรอบตรวจ';
      edit.setAttribute('aria-label','แก้ไขรอบตรวจ '+r.name);
      edit.innerHTML='<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="m4 16.5-.8 4.3 4.3-.8L18.8 8.7l-3.5-3.5L4 16.5Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="m13.8 6.7 3.5 3.5" stroke="currentColor" stroke-width="1.8"/></svg>';
      edit.onclick=()=>editRoundModal(r);
      actions.appendChild(edit);
      const del=document.createElement('button');
      del.type='button';
      del.className='btn round-delete-btn';
      del.title='ลบรอบตรวจ';
      del.setAttribute('aria-label','ลบรอบตรวจ '+r.name);
      del.innerHTML='<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 7h16M9 7V4h6v3M7 7l1 13h8l1-13M10 11v5M14 11v5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';
      del.onclick=()=>requestDeleteRound(r.id,r.name);
      actions.appendChild(del);
      box.appendChild(d);
    })
  }
  function requestDeleteRound(id,name){
    window.pendingDeleteRound={id:String(id),name:String(name||'')};
    el('actionBox').className='action-box error';
    el('actionContent').className='action-content';
    el('actionContent').innerHTML='<div class="modal-status-icon" aria-hidden="true">!</div><h3>ลบรอบตรวจนี้หรือไม่?</h3><p>ระบบจะลบรอบ “'+esc(name)+'” และลบชีทผลตรวจของรอบนี้ด้วย การดำเนินการนี้ย้อนกลับไม่ได้</p><div class="photo-actions" style="justify-content:center"><button type="button" class="btn danger" onclick="confirmDeleteRound()">ลบรอบตรวจ</button><button type="button" class="btn secondary" onclick="cancelDeleteRound()">ยกเลิก</button></div>';
    el('actionClose').classList.add('hidden');
    el('actionModal').classList.remove('hidden');
  }
  function cancelDeleteRound(){window.pendingDeleteRound=null;closeActionModal()}
  function confirmDeleteRound(){
    const target=window.pendingDeleteRound;
    if(!target)return;
    actionLoading('กำลังลบรอบตรวจ');
    google.script.run.withSuccessHandler(x=>{
      window.pendingDeleteRound=null;
      boot.rounds=(boot.rounds||[]).filter(r=>String(r.id)!==String(target.id));
      renderRounds();
      updateMetrics();
      el('dash').classList.add('hidden');
      actionResult('ลบรอบตรวจสำเร็จ','ลบรอบตรวจและชีท '+(x.sheetName||'')+' เรียบร้อยแล้ว',false);
      loadAdmin();
    }).withFailureHandler(e=>actionResult('ลบรอบตรวจไม่สำเร็จ',err(e),true)).deleteRoundAdmin(target.id);
  }
  function setStatus(id,s){actionLoading('กำลัง'+s+'รอบตรวจ');google.script.run.withSuccessHandler(()=>{const changed=boot.rounds.find(r=>String(r.id)===String(id));if(changed)changed.status=s;renderRounds();updateMetrics();actionResult(s+'รอบตรวจสำเร็จ','สถานะรอบตรวจอัปเดตเรียบร้อยแล้ว',false);loadAdmin()}).withFailureHandler(e=>actionResult('เปลี่ยนสถานะไม่สำเร็จ',err(e),true)).setRoundStatusAdmin(String(id),String(s))}
  function dashboard(id){el('dash').classList.remove('hidden');el('dash').innerHTML='<div class="card"><div class="spinner"></div><h3 style="text-align:center">กำลังโหลดรายงาน</h3><p style="text-align:center;color:#6b7972">กรุณารอสักครู่</p></div>';el('dash').scrollIntoView({behavior:'smooth',block:'start'});let finished=false;const timer=setTimeout(()=>{if(finished)return;finished=true;el('dash').innerHTML='<div class="msg error"><b>โหลดรายงานใช้เวลานานเกินไป</b><br>กรุณาลองใหม่อีกครั้ง</div>'},20000);google.script.run.withSuccessHandler(x=>{if(finished)return;finished=true;clearTimeout(timer);const rows=x.records||[];const tableRows=rows.length?rows.map(r=>'<tr><td class="asset-link">'+esc(r.assetNo)+'</td><td>'+esc(r.name)+'</td><td>'+esc(r.inspector)+'</td><td>'+esc(r.inspectedAt)+'</td><td><span class="result-chip">'+esc(r.result)+'</span></td></tr>').join(''):'<tr><td colspan="5" class="empty">ยังไม่มีผลการตรวจในรอบนี้</td></tr>';el('dash').innerHTML='<div class="section-title"><div><h2>รายงานผลการตรวจ</h2><p>'+esc(x.roundName||'รายละเอียดผลการตรวจครุภัณฑ์')+'</p></div></div><div class="summary"><div class="num"><b>'+Number(x.total||0).toLocaleString('th-TH')+'</b>ทั้งหมด</div><div class="num"><b>'+Number(x.inspected||0).toLocaleString('th-TH')+'</b>ตรวจแล้ว</div><div class="num"><b>'+Number(x.remaining||0).toLocaleString('th-TH')+'</b>คงเหลือ</div></div><div class="report-heading"><h3>รายการผลตรวจ</h3><span>จำนวน '+rows.length.toLocaleString('th-TH')+' รายการ</span></div><div class="table-wrap report-table"><table><thead><tr><th>เลขครุภัณฑ์</th><th>รายการ</th><th>ผู้ตรวจ</th><th>วันที่ตรวจ</th><th>สถานะ</th></tr></thead><tbody>'+tableRows+'</tbody></table></div>'}).withFailureHandler(e=>{if(finished)return;finished=true;clearTimeout(timer);el('dash').innerHTML='<div class="msg error"><b>โหลดรายงานไม่สำเร็จ</b><br>'+esc(err(e))+'</div>'}).getDashboardAdmin(String(id))}
  let photoData='',photoItems=[],photoListLoaded=false;
  function showAdminPanel(name){['home','inspect','photo'].forEach(x=>{el(x+'Panel').classList.toggle('hidden',x!==name);el(x+'Tab').className='btn '+(x===name?'active':'secondary')});const titles={home:'หน้าหลัก',inspect:'ตรวจครุภัณฑ์',photo:'เพิ่มรูปครุภัณฑ์'};if(el('topbarTitle'))el('topbarTitle').textContent=titles[name]||'ระบบจัดการ';if(name==='photo'&&!photoListLoaded)loadPhotoList();window.scrollTo({top:0,behavior:'smooth'})}
  function loadPhotoList(){el('photoListLoading').classList.remove('hidden');el('photoList').classList.add('hidden');el('photoEmpty').classList.add('hidden');google.script.run.withSuccessHandler(x=>{photoItems=x.items||[];photoListLoaded=true;el('photoTotal').textContent=Number(x.total||0).toLocaleString('th-TH');el('photoAdded').textContent=Number(x.withPhoto||0).toLocaleString('th-TH');el('photoMissing').textContent=Number(x.withoutPhoto||0).toLocaleString('th-TH');el('photoListLoading').classList.add('hidden');filterPhotoList()}).withFailureHandler(e=>{el('photoListLoading').innerHTML='<div class="msg error"><b>โหลดรายการครุภัณฑ์ไม่สำเร็จ</b><br>'+esc(err(e))+'<br><button class="btn outline" onclick="loadPhotoList()">ลองใหม่</button></div>'}).getAssetPhotoListAdmin()}
  function filterPhotoList(){const q=(el('photoSearch').value||'').trim().toLowerCase(),status=el('photoStatusFilter').value;const rows=photoItems.map((x,i)=>({x:x,i:i})).filter(o=>{const x=o.x,has=!!(x.photo&&x.photo.url);if(status==='added'&&!has)return false;if(status==='missing'&&has)return false;return !q||[x.assetNo,x.id,x.name,x.type,x.custodian].join(' ').toLowerCase().includes(q)});const box=el('photoList');el('photoEmpty').classList.toggle('hidden',rows.length>0);box.classList.toggle('hidden',rows.length===0);box.innerHTML=rows.map(o=>{const x=o.x,has=!!(x.photo&&x.photo.url);return '<div class="photo-row"><div class="photo-thumb">'+(has?'<img src="'+esc(x.photo.url)+'" alt="รูป '+esc(x.assetNo)+'" loading="lazy">':'ยังไม่มีรูป')+'</div><div class="photo-detail"><b>'+esc(x.assetNo||x.id)+'</b><span>'+esc(x.name||'ไม่ระบุชื่อรายการ')+'</span><small>'+esc(x.type||'ไม่ระบุประเภท')+(x.custodian?' · ผู้ครอง '+esc(x.custodian):'')+'</small></div><div><span class="photo-state '+(has?'added':'')+'">'+(has?'เพิ่มแล้ว':'ยังไม่เพิ่ม')+'</span>'+(has&&x.photo.updatedAt?'<small style="display:block;margin-top:5px;color:#77847d">'+esc(x.photo.updatedAt)+'</small>':'')+'</div><button class="btn '+(has?'outline':'primary')+'" onclick="openPhotoEditor('+o.i+')">'+(has?'แก้ไขรูป':'เพิ่มรูป')+'</button></div>'}).join('')}
  function openPhotoEditor(index){const x=photoItems[index];if(!x)return;photoData='';el('photoAssetNo').value=x.assetNo||x.id;el('photoEditorTitle').textContent=x.photo&&x.photo.url?'แก้ไขรูปครุภัณฑ์ภาพจริง':'เพิ่มรูปครุภัณฑ์ภาพจริง';el('photoEditorSub').textContent=(x.assetNo||x.id)+' · '+(x.name||'ไม่ระบุชื่อรายการ');el('photoInfo').classList.remove('hidden');el('photoInfo').innerHTML='<b>'+esc(x.name||'ไม่ระบุชื่อรายการ')+'</b><br><small>'+(x.photo&&x.photo.url?'รูปด้านล่างคือรูปที่ใช้อยู่ในหน้าตรวจ':'รายการนี้ยังไม่มีรูปภาพจริง')+'</small>';if(x.photo&&x.photo.url){el('photoPreview').src=x.photo.url;el('photoPreview').classList.remove('hidden')}else{el('photoPreview').classList.add('hidden');el('photoPreview').removeAttribute('src')}el('photoFile').value='';el('photoSave').classList.add('hidden');el('photoDelete').classList.toggle('hidden',!(x.photo&&x.photo.url));el('photoEditorCard').classList.remove('hidden');document.body.style.overflow='hidden'}
  function closePhotoEditor(){el('photoEditorCard').classList.add('hidden');document.body.style.overflow='';photoData='';el('photoFile').value=''}
  function requestDeletePhoto(){const no=el('photoAssetNo').value.trim();el('actionBox').className='action-box error';el('actionContent').className='action-content';el('actionContent').innerHTML='<div class="modal-status-icon" aria-hidden="true">!</div><h3>ลบรูปครุภัณฑ์นี้หรือไม่?</h3><p>รูปของ '+esc(no)+' จะถูกนำออกจากหน้าตรวจ และสามารถเพิ่มรูปใหม่ภายหลังได้</p><div class="photo-actions" style="justify-content:center"><button class="btn danger" onclick="confirmDeletePhoto()">ลบรูป</button><button class="btn secondary" onclick="closeActionModal()">ยกเลิก</button></div>';el('actionClose').classList.add('hidden');el('actionModal').classList.remove('hidden')}
  function confirmDeletePhoto(){const no=el('photoAssetNo').value.trim();actionLoading('กำลังลบรูปครุภัณฑ์');google.script.run.withSuccessHandler(()=>{const item=photoItems.find(v=>String(v.assetNo||v.id)===String(no));if(item)item.photo=null;const added=photoItems.filter(v=>v.photo&&v.photo.url).length;el('photoAdded').textContent=added.toLocaleString('th-TH');el('photoMissing').textContent=(photoItems.length-added).toLocaleString('th-TH');filterPhotoList();closePhotoEditor();actionResult('ลบรูปสำเร็จ','นำรูปออกจากรายการครุภัณฑ์และหน้าตรวจแล้ว',false)}).withFailureHandler(e=>actionResult('ลบรูปไม่สำเร็จ',err(e),true)).deleteAssetPhotoAdmin(no)}
  function preparePhoto(event){const file=event.target.files&&event.target.files[0];event.target.value='';if(!file)return;if(!/^image\\/(jpeg|png|webp)$/.test(file.type))return actionResult('ไฟล์รูปไม่ถูกต้อง','กรุณาเลือกไฟล์ JPG, PNG หรือ WEBP',true);actionLoading('กำลังแปลงรูปเป็น WebP');const reader=new FileReader();reader.onload=()=>{const img=new Image();img.onload=()=>{const max=1600,scale=Math.min(1,max/Math.max(img.width,img.height));const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(img.width*scale));canvas.height=Math.max(1,Math.round(img.height*scale));const ctx=canvas.getContext('2d');ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';ctx.drawImage(img,0,0,canvas.width,canvas.height);canvas.toBlob(blob=>{if(!blob||blob.type!=='image/webp')return actionResult('แปลงรูปไม่สำเร็จ','อุปกรณ์นี้ไม่รองรับการแปลงไฟล์ WebP กรุณาใช้เบราว์เซอร์เวอร์ชันล่าสุด',true);const webpReader=new FileReader();webpReader.onload=()=>{photoData=webpReader.result;el('photoPreview').src=photoData;el('photoPreview').classList.remove('hidden');el('photoSave').classList.remove('hidden');closeActionModal()};webpReader.onerror=()=>actionResult('อ่านรูปไม่สำเร็จ','กรุณาเลือกรูปใหม่อีกครั้ง',true);webpReader.readAsDataURL(blob)},'image/webp',.82)};img.onerror=()=>actionResult('อ่านรูปไม่สำเร็จ','กรุณาเลือกรูปใหม่อีกครั้ง',true);img.src=reader.result};reader.onerror=()=>actionResult('อ่านรูปไม่สำเร็จ','กรุณาเลือกรูปใหม่อีกครั้ง',true);reader.readAsDataURL(file)}
  function uploadPhoto(){const no=el('photoAssetNo').value.trim();if(!no)return actionResult('ไม่พบเลขครุภัณฑ์','กรุณาเลือกรายการครุภัณฑ์อีกครั้ง',true);if(!photoData)return actionResult('ยังไม่ได้เลือกรูป','กรุณาถ่ายรูปหรือเลือกรูปก่อนบันทึก',true);actionLoading('กำลังบันทึกรูปครุภัณฑ์');el('photoSave').disabled=true;google.script.run.withSuccessHandler(x=>{el('photoSave').disabled=false;const item=photoItems.find(v=>String(v.assetNo||v.id)===String(no));if(item)item.photo=x.photo;photoData='';el('photoFile').value='';el('photoSave').classList.add('hidden');photoListLoaded=true;const added=photoItems.filter(v=>v.photo&&v.photo.url).length;el('photoAdded').textContent=added.toLocaleString('th-TH');el('photoMissing').textContent=(photoItems.length-added).toLocaleString('th-TH');filterPhotoList();if(item)openPhotoEditor(photoItems.indexOf(item));actionResult('บันทึกรูปสำเร็จ','รูปภาพจริงถูกอัปเดตและจะใช้ในหน้าตรวจทันที',false)}).withFailureHandler(e=>{el('photoSave').disabled=false;actionResult('บันทึกรูปไม่สำเร็จ',err(e),true)}).saveAssetPhotoAdmin({assetNo:no,dataUrl:photoData})}
  function initDatePickers(){if(typeof flatpickr!=='function')return;const locale=(flatpickr.l10ns&&flatpickr.l10ns.th)||'default';flatpickr('#start',{locale:locale,dateFormat:'Y-m-d',altInput:true,altFormat:'j F Y',disableMobile:true,monthSelectorType:'static'});flatpickr('#end',{locale:locale,dateFormat:'Y-m-d',altInput:true,altFormat:'j F Y',disableMobile:true,monthSelectorType:'static'})}
  initDatePickers();loadAdmin();`}

function repairStatusOverviewNow() {
  refreshAssetInspectionOverview();
}

function repairInspectionStatusesNow() {
  const db = db_();
  const statusSheet = db.getSheetByName('รายการสถานะ');
  const statusRows = [
    ['FOUND','ใช้งานอยู่',false,false,'#079447',1,true],
    ['DAMAGED','ชำรุด',true,false,'#dc2626',2,true],
    ['MOVED','เปลี่ยนผู้ครอบครอง',false,true,'#2563eb',3,true]
  ];
  if (statusSheet.getLastRow() > 1) statusSheet.getRange(2,1,statusSheet.getLastRow()-1,statusSheet.getMaxColumns()).clearContent();
  statusSheet.getRange(2,1,statusRows.length,statusRows[0].length).setValues(statusRows);
  const names = {FOUND:'ใช้งานอยู่',DAMAGED:'ชำรุด',MOVED:'เปลี่ยนผู้ครอบครอง'};
  const rounds = db.getSheetByName('รอบตรวจ');
  if (rounds && rounds.getLastRow() > 1) {
    const sheetNames = rounds.getRange(2,6,rounds.getLastRow()-1,1).getDisplayValues().flat().filter(String);
    sheetNames.forEach(function(sheetName) {
      const result = db.getSheetByName(sheetName);
      if (!result || result.getLastRow() < 2) return;
      const codes = result.getRange(2,9,result.getLastRow()-1,1).getDisplayValues();
      result.getRange(2,10,codes.length,1).setValues(codes.map(function(row) { return [names[String(row[0]).trim()] || String(row[0]).trim()]; }));
    });
  }
  refreshAssetInspectionOverview();
  SpreadsheetApp.flush();
}
