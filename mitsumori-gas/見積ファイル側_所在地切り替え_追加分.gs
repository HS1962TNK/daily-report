// ★このファイルは「見積ひな形」の Apps Script（コード.gs）の一番下に「貼り足す」ための追加分です。
//   既存のコード（CSV一括作成の BI_ で始まる関数など）は消さずに、この内容を末尾に足してください。
//   あわせて onOpen() のメニューに次の1行を足すこと（「この見積を複製」の行のすぐ下）：
//     .addItem('所在地を担当部署に合わせる', 'applyOfficeAddressMenu')
//   ※既存のコードに function onEdit がすでにある場合は、二重にならないよう知らせてください。

// ── 所在地（本社／大阪営業所）の切り替え（2026-10-07） ──
// 担当部署(J10)が「大阪営業所」なら大阪営業所の所在地、それ以外の部署はすべて本社の所在地を表紙に入れる。
// 中身は母艦の「所在地設定」シートで決める（A列=セル番地／B列=本社／C列=大阪営業所）。
// 母艦が見積を作るときに、このファイルの「所在地設定」シート（非表示）へ写しを入れておくので、
// 担当部署を選び直したときは onEdit でその写しを使う（単純トリガーは他のファイルを開けないため）。
// ※母艦コード.gsの同名の処理と同じ判定。片方を直したらもう片方も直すこと。
const OFFICE_SHEET = '所在地設定';
const OFFICE_DEPT_CELL = 'J10';

function normalizeOfficeName_(s) {
  return String(s || '').replace(/[\s　]/g, '');
}

function readOfficeTable_(sheet) {
  if (!sheet) return null;
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return null;
  return sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
}

function officeColumnFor_(table, dept) {
  const d = normalizeOfficeName_(dept);
  if (d) {
    for (let c = 2; c < table[0].length; c++) {
      const h = normalizeOfficeName_(table[0][c]);
      if (h && h === d) return c;
    }
  }
  return 1;
}

function applyOfficeAddress_(cover, table) {
  if (!cover || !table) return '';
  const col = officeColumnFor_(table, cover.getRange(OFFICE_DEPT_CELL).getValue());
  for (let r = 1; r < table.length; r++) {
    const a1 = String(table[r][0] || '').trim().toUpperCase();
    if (!/^[A-Z]{1,2}[0-9]{1,3}$/.test(a1)) continue;
    const cell = cover.getRange(a1);
    const v = table[r][col];
    if (String(cell.getDisplayValue()) !== String(v)) cell.setValue(v);
  }
  return normalizeOfficeName_(table[0][col]) || '本社';
}

// 表紙の担当部署を選び直したら、所在地を切り替える（単純トリガー）。
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== '見積書') return;
    const dept = sheet.getRange(OFFICE_DEPT_CELL);
    const r = e.range;
    if (dept.getRow() < r.getRow() || dept.getRow() > r.getLastRow() ||
        dept.getColumn() < r.getColumn() || dept.getColumn() > r.getLastColumn()) return;
    applyOfficeAddress_(sheet, readOfficeTable_(e.source.getSheetByName(OFFICE_SHEET)));
  } catch (err) {
    // 所在地の切り替えに失敗しても入力そのものは止めない
  }
}

// メニュー用。母艦の最新の所在地設定を取り込んでから表紙に反映する（古い見積ファイルにも使える）。
function applyOfficeAddressMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cover = ss.getSheetByName('見積書');
  if (!cover) { ui.alert('見積書シートが見つかりません。'); return; }
  let table = null;
  try {
    table = readOfficeTable_(SpreadsheetApp.openById(MASTER_SPREADSHEET_ID).getSheetByName(OFFICE_SHEET));
    if (table) {
      const local = ss.getSheetByName(OFFICE_SHEET) || ss.insertSheet(OFFICE_SHEET);
      local.clear();
      local.getRange(1, 1, table.length, table[0].length).setNumberFormat('@').setValues(table);
      local.hideSheet();
      ss.setActiveSheet(cover);
    }
  } catch (e) {}
  if (!table) table = readOfficeTable_(ss.getSheetByName(OFFICE_SHEET));
  if (!table) {
    ui.alert('「' + OFFICE_SHEET + '」が見つかりません。\n母艦のメニュー「見積書メーカー」→「所在地設定シートを作る（初回のみ）」で作ってください。');
    return;
  }
  const office = applyOfficeAddress_(cover, table);
  ui.alert('表紙の所在地を「' + office + '」の内容にしました。');
}
