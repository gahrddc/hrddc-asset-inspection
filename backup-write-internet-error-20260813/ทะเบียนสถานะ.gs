const ASSET_OVERVIEW_SHEET_ = 'สถานะการตรวจครุภัณฑ์';

function setupAssetInspectionOverview() {
  refreshAssetInspectionOverview();
  return 'สร้างทะเบียนสถานะการตรวจเรียบร้อยแล้ว';
}

function refreshAssetInspectionOverview() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const db = SpreadsheetApp.openById(DB_ID);
    let overview = db.getSheetByName(ASSET_OVERVIEW_SHEET_);
    const isNew = !overview;
    if (!overview) overview = db.insertSheet(ASSET_OVERVIEW_SHEET_);

    if (isNew || overview.getLastRow() < 2) {
      const source = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET);
      if (!source) throw new Error('ไม่พบชีทข้อมูลต้นฉบับ');
      const sourceLastRow = source.getLastRow();
      const rows = sourceLastRow > 1
        ? source.getRange(2, 1, sourceLastRow - 1, 15).getDisplayValues()
        : [];
      const assets = rows
        .filter(function(row) { return String(row[2] || '').trim() !== ''; })
        .map(function(row) {
          return [row[2], row[5], row[13], row[14]];
        });

      overview.clear();
      overview.getRange(1, 1, 1, 4).setValues([[
        'เลขครุภัณฑ์',
        'ชื่อครุภัณฑ์',
        'ผู้รับผิดชอบ',
        'กลุ่ม'
      ]]);
      if (assets.length) {
        overview.getRange(2, 1, assets.length, 4).setValues(assets);
      }
      formatAssetOverviewBase_(overview, Math.max(assets.length + 1, 2));
    }

    setDynamicRoundStatusFormula_(overview);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function setDynamicRoundStatusFormula_(sheet) {
  if (sheet.getMaxColumns() > 4) {
    sheet.getRange(1, 5, sheet.getMaxRows(), sheet.getMaxColumns() - 4).clearContent();
  }
  const formula = '=LET(' +
    'starts,FILTER(\'รอบตรวจ\'!C2:C,\'รอบตรวจ\'!A2:A<>""),' +
    'resultSheets,FILTER(\'รอบตรวจ\'!F2:F,\'รอบตรวจ\'!A2:A<>""),' +
    'assetIds,FILTER(A2:A,A2:A<>""),' +
    'headers,TRANSPOSE("สถานะการตรวจ "&DAY(starts)&" "&CHOOSE(MONTH(starts),"มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม")&" "&(YEAR(starts)+543)),' +
    'statuses,MAKEARRAY(ROWS(assetIds),ROWS(starts),LAMBDA(r,c,' +
      'LET(matched,IFERROR(MATCH(INDEX(assetIds,r,1),INDIRECT("\'"&INDEX(resultSheets,c,1)&"\'!E:E"),0),0),' +
      'IF(matched=0,"ยังไม่ตรวจ",INDEX(INDIRECT("\'"&INDEX(resultSheets,c,1)&"\'!J:J"),matched))))),' +
    'VSTACK(headers,statuses))';

  sheet.getRange('E1').setFormula(formula);
  sheet.setFrozenRows(1);
  const maxStatusColumns = Math.min(50, sheet.getMaxColumns() - 4);
  if (maxStatusColumns > 0) {
    const headerRange = sheet.getRange(1, 5, 1, maxStatusColumns);
    headerRange.setBackground('#0b7a47').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    for (let column = 5; column < 5 + maxStatusColumns; column++) sheet.setColumnWidth(column, 210);
    const statusRange = sheet.getRange(2, 5, Math.max(sheet.getMaxRows() - 1, 1), maxStatusColumns);
    const keptRules = sheet.getConditionalFormatRules().filter(function(rule) {
      return !rule.getRanges().some(function(range) { return range.getColumn() >= 5; });
    });
    const rules = [
      ['ใช้งานอยู่','#dff3e7','#0b7a47'],
      ['ชำรุด','#fde7e7','#c62828'],
      ['เปลี่ยนผู้ครอบครอง','#e8f1ff','#2563b8'],
      ['ยังไม่ตรวจ','#f3f4f6','#6b7280']
    ].map(function(item) {
      return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(item[0]).setBackground(item[1]).setFontColor(item[2]).setRanges([statusRange]).build();
    });
    sheet.setConditionalFormatRules(keptRules.concat(rules));
  }
}

function formatAssetOverviewBase_(sheet, lastRow) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#0b7a47')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(1, 1, lastRow, 4)
    .setFontFamily('Kanit')
    .setVerticalAlignment('middle');
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 150);
  sheet.setRowHeight(1, 36);
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).setWrap(true);
  sheet.getRange(1, 1, lastRow, 4)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN);
}
