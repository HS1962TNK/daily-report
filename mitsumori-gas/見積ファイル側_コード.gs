const MASTER_SPREADSHEET_ID = '18XrMR3ZdcuvpS-dmkvF_LArDsOWaUHrF2IK6P8SHb-w';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('見積機能')
    .addItem('見積一覧に登録・更新', 'updateQuoteIndex')
    .addItem('この見積を複製（新規保存）', 'duplicateThisQuote')
    .addSeparator()
    .addItem('CSVから一括作成', 'BI_showImportDialog')
    .addToUi();
}

function updateQuoteIndex() {
  const ui = SpreadsheetApp.getUi();
  try {
    const thisSs = SpreadsheetApp.getActiveSpreadsheet();
    const masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const indexSheet = masterSs.getSheetByName('見積一覧');
    if (!indexSheet) { ui.alert('見積一覧シートが見つかりません。'); return; }
    
    const cover = thisSs.getSheetByName('見積書');
    if (!cover) { ui.alert('見積書シートが見つかりません。'); return; }
    
    const quoteNo = cover.getRange('L2').getValue();
    const quoteDate = cover.getRange('L3').getValue();
    const client = cover.getRange('B5').getValue();
    const jobName = cover.getRange('C9').getValue();
    const department = cover.getRange('I10').getValue();
    const person = cover.getRange('L10').getValue();
    
    const data = indexSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == quoteNo) {
        indexSheet.getRange(i+1, 2).setValue(quoteDate);
        indexSheet.getRange(i+1, 3).setValue(client);
        indexSheet.getRange(i+1, 4).setValue(jobName);
        indexSheet.getRange(i+1, 5).setValue(department);
        indexSheet.getRange(i+1, 6).setValue(person);
        found = true;
        break;
      }
    }
    if (found) {
      ui.alert('見積一覧を更新しました。');
    } else {
      ui.alert('見積NO「' + quoteNo + '」が見積一覧に見つかりません。');
    }
  } catch (e) {
    ui.alert('エラーが発生しました：' + e.message);
  }
}

function duplicateThisQuote() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('この見積を複製', '新しい見積NO（例：260708-002）を入力してください', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const newNo = res.getResponseText();
  if (!newNo || !newNo.trim()) { ui.alert('見積NOが入力されていません。'); return; }
  
  try {
    const thisSs = SpreadsheetApp.getActiveSpreadsheet();
    const folder = thisSs.getParents().next();
    const newFile = thisSs.makeCopy('見積' + newNo.trim(), folder);
    const newSs = SpreadsheetApp.openById(newFile.getId());
    const cover = newSs.getSheetByName('見積書');
    if (cover) { cover.getRange('L2').setValue(newNo.trim()); }
    ui.alert('新しい見積ファイルを作成しました。\nファイル名：見積' + newNo.trim());
  } catch (e) {
    ui.alert('複製に失敗しました：' + e.message);
      }
      }

function redesignCoverSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('見積書');
  sh.setColumnWidth(1,16);
  sh.setColumnWidth(2,133);
  sh.setColumnWidth(3,213);
  sh.setColumnWidth(4,75);
  sh.setColumnWidth(5,66);
  sh.setColumnWidth(7,213);
  sh.setColumnWidth(8,126);
  sh.setColumnWidth(9,66);
  sh.setColumnWidth(10,127);
  sh.setColumnWidth(11,74);
  sh.setColumnWidth(12,66);
  sh.setColumnWidth(13,102);
  sh.setColumnWidth(14,68);
  sh.setColumnWidth(15,65);
  sh.setColumnWidth(16,66);
  sh.getRange(11,1,20,16).clearContent();
  sh.getRange(11,1,17,16).breakApart();
  sh.getRange('B11').setValue('名 称');
  sh.getRange('D11').setValue('品 質・規 格');
  sh.getRange('H11').setValue('数 量');
  sh.getRange('I11').setValue('呼 称');
  sh.getRange('J11').setValue('単 価');
  sh.getRange('K11').setValue('金 額');
  sh.getRange('D12').setValue('別紙内訳書 1');
  sh.getRange('H12').setValue(1);
  sh.getRange('I12').setValue('式');
  sh.getRange('K12').setFormula("='内訳書1'!M19");
  sh.getRange('D13').setValue('別紙内訳書 2');
  sh.getRange('H13').setValue(1);
  sh.getRange('I13').setValue('式');
  sh.getRange('K13').setFormula("='内訳書2'!M19");
  sh.getRange('D14').setValue('別紙内訳書 3');
  sh.getRange('H14').setValue(1);
  sh.getRange('I14').setValue('式');
  sh.getRange('K14').setFormula("='内訳書3'!M19");
  sh.getRange('D15').setValue('別紙内訳書 4');
  sh.getRange('H15').setValue(1);
  sh.getRange('I15').setValue('式');
  sh.getRange('K15').setFormula("='内訳書4'!M19");
  sh.getRange('D16').setValue('別紙内訳書 5');
  sh.getRange('H16').setValue(1);
  sh.getRange('I16').setValue('式');
  sh.getRange('K16').setFormula("='内訳書5'!M19");
  sh.getRange('B22').setValue('法定福利費');
  sh.getRange('H22').setValue(1);
  sh.getRange('I22').setValue('式');
  sh.getRange('K22').setFormula("=ROUND(SUM(K12:K21)*'設定'!D4,0)");
  sh.getRange('B23').setValue('諸経費');
  sh.getRange('H23').setValue(1);
  sh.getRange('I23').setValue('式');
  sh.getRange('K23').setFormula("=ROUND(SUM(K12:K21)*'設定'!D5,0)");
  sh.getRange('B24').setValue('見積条件');
  sh.getRange('I24').setValue('小 計');
  sh.getRange('K24').setFormula('=SUM(K12:K23)');
  sh.getRange('B25').setValue('・\n・\n・\n・\n・\n・\n・');
  sh.getRange('I25').setValue('端 数 調 整');
  sh.getRange('K25').setFormula('=ROUNDDOWN(K24,-4)-K24');
  sh.getRange('I26').setValue('消 費 税');
  sh.getRange('K26').setFormula('=ROUND((K24+K25)*0.1,0)');
  sh.getRange('I27').setValue('合 計（税込）');
  sh.getRange('K27').setFormula('=K24+K25+K26');
  sh.getRange('B11:C11').merge();
  sh.getRange('D11:G11').merge();
  sh.getRange('K11:M11').merge();
  sh.getRange('B12:C12').merge();
  sh.getRange('D12:G12').merge();
  sh.getRange('K12:M12').merge();
  sh.getRange('B13:C13').merge();
  sh.getRange('D13:G13').merge();
  sh.getRange('K13:M13').merge();
  sh.getRange('B14:C14').merge();
  sh.getRange('D14:G14').merge();
  sh.getRange('K14:M14').merge();
  sh.getRange('B15:C15').merge();
  sh.getRange('D15:G15').merge();
  sh.getRange('K15:M15').merge();
  sh.getRange('B16:C16').merge();
  sh.getRange('D16:G16').merge();
  sh.getRange('K16:M16').merge();
  sh.getRange('B17:C17').merge();
  sh.getRange('D17:G17').merge();
  sh.getRange('K17:M17').merge();
  sh.getRange('B18:C18').merge();
  sh.getRange('D18:G18').merge();
  sh.getRange('K18:M18').merge();
  sh.getRange('B19:C19').merge();
  sh.getRange('D19:G19').merge();
  sh.getRange('K19:M19').merge();
  sh.getRange('B20:C20').merge();
  sh.getRange('D20:G20').merge();
  sh.getRange('K20:M20').merge();
  sh.getRange('B21:C21').merge();
  sh.getRange('D21:G21').merge();
  sh.getRange('K21:M21').merge();
  sh.getRange('B22:C22').merge();
  sh.getRange('D22:G22').merge();
  sh.getRange('K22:M22').merge();
  sh.getRange('B23:C23').merge();
  sh.getRange('D23:G23').merge();
  sh.getRange('K23:M23').merge();
  sh.getRange('I24:J24').merge();
  sh.getRange('K24:M24').merge();
  sh.getRange('B25:H27').merge();
  sh.getRange('I25:J25').merge();
  sh.getRange('K25:M25').merge();
  sh.getRange('I26:J26').merge();
  sh.getRange('K26:M26').merge();
  sh.getRange('I27:J27').merge();
  sh.getRange('K27:M27').merge();
  sh.getRange('J12:J23').setNumberFormat('#,##0');
  sh.getRange('K12:K27').setNumberFormat('#,##0');
  sh.getRange(1,1,200,26).setFontFamily('MS PGothic');
}

function redesignBreakdownSheet_(n) {
var ss = SpreadsheetApp.getActive();
var sh = ss.getSheetByName('内訳書' + n);
sh.setColumnWidth(1,26);
sh.setColumnWidth(2,85);
sh.setColumnWidth(3,66);
sh.setColumnWidth(4,81);
sh.setColumnWidth(5,75);
sh.setColumnWidth(6,66);
sh.setColumnWidth(8,154);
sh.setColumnWidth(9,137);
sh.setColumnWidth(10,61);
sh.setColumnWidth(11,66);
sh.setColumnWidth(12,61);
sh.setColumnWidth(13,75);
sh.setColumnWidth(14,112);
sh.setColumnWidth(15,65);
sh.setColumnWidth(16,66);
sh.getRange(1,1,25,16).clearContent();
sh.getRange(1,1,22,16).breakApart();
sh.getRange('B1').setValue('内 訳 書　' + n);
sh.getRange('B1').setFontSize(16);
sh.getRange('B1:N2').merge();
sh.getRange('B4').setValue('名 称');
sh.getRange('E4').setValue('品 質・規 格');
sh.getRange('I4').setValue('数 量');
sh.getRange('J4').setValue('呼 称');
sh.getRange('K4').setValue('単 価');
sh.getRange('M4').setValue('金 額');
sh.getRange('B4:D4').merge();
sh.getRange('E4:H4').merge();
sh.getRange('K4:L4').merge();
sh.getRange('M4:N4').merge();
for (var r = 5; r <= 18; r++) {
sh.getRange('B'+r+':D'+r).merge();
sh.getRange('E'+r+':H'+r).merge();
sh.getRange('K'+r+':L'+r).merge();
sh.getRange('M'+r+':N'+r).merge();
sh.getRange('M'+r).setFormula('=IF(OR(I'+r+'="",K'+r+'=""),0,I'+r+'*K'+r+')');
}
sh.getRange('B19').setValue('小 計');
sh.getRange('B19:L19').merge();
sh.getRange('M19').setFormula('=SUM(M5:N18)');
sh.getRange('M19:N19').merge();
sh.getRange('K5:L18').setNumberFormat('#,##0');
sh.getRange('M5:N19').setNumberFormat('#,##0');
sh.getRange(1,1,200,26).setFontFamily('MS PGothic');
}

function redesign2026_ApplyAll() {
redesignCoverSheet_();
redesignBreakdownSheet_(1);
redesignBreakdownSheet_(2);
redesignBreakdownSheet_(3);
redesignBreakdownSheet_(4);
redesignBreakdownSheet_(5);
SpreadsheetApp.getActive().toast('レイアウト調整が完了しました');
}


function migrateRecord2_() {
  var TEMPLATE_ID = '1JBKTU9zgI8c_329JDYtPcVSxUeHj8OOm8ge680_acYk';
  var tpl = DriveApp.getFileById(TEMPLATE_ID);
  var parents = tpl.getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var newName = '見積書_25090065_タカラ住販株式会社_沢口ビル外壁塗装工事';
  var copy = tpl.makeCopy(newName, folder);
  var ss = SpreadsheetApp.openById(copy.getId());
  var cover = ss.getSheetByName('見積書');

  cover.getRange('L2').setValue('25090065');
  cover.getRange('L3').setValue(new Date(2025,8,2));
  cover.getRange('B5').setValue('タカラ住販株式会社');
  cover.getRange('F5').setValue('御中');
  cover.getRange('C9').setValue('沢口ビル外壁塗装工事');
  cover.getRange('C10').setValue('東京都足立区中央本町1-18-9');
  cover.getRange('I10').setValue('リフォーム事業部');
  cover.getRange('L10').setValue('小島 明生');

  var categories = ['仮設工事','シーリング工事','塗装工事','ベランダ防水工事','雑工事'];
  for (var i=0;i<categories.length;i++){
    cover.getRange('B'+(12+i)).setValue(categories[i]);
  }

  cover.getRange('M22').setValue(150000);
  cover.getRange('M23').setValue(140000);
  cover.getRange('M25').setValue(-45300);

  var data1 = [
    ['外部足場架払','単管ブラケット足場 架け払い',469.2,850],
    ['外部養生','飛散防止メッシュシート',469.2,150],
    ['道路占有許可費','東側道路部、書類作成費、申請手間含む',1,100000],
    ['防護管取付費','近接高圧電線、引込電線',1,150000],
    ['ガードマン費','交通誘導（足場仮設・解体時）',4,23000]
  ];
  var data2 = [
    ['ALC開口シール','（20×15）ノンブリード 既存撤去後、新規シール',112.4,1100],
    ['下端笠木シール','（20×15）ノンブリード 既存撤去後、新規シール',11.7,1100],
    ['ＡＬＣ板縦目地シール','（20×10）ノンブリード 増打ち',584.8,650],
    ['サッシ皿上シール','（10×10）変成シリコーン 既存撤去後、新規シール',25.8,900],
    ['雑シール','顎上変シリ、貫通部等',1,30000]
  ];
  var data3 = [
    ['外壁下地補修','ひび割れ補修 Uカット処理 シール充填',1,80000],
    ['高圧洗浄','15Mps',388.8,150],
    ['各所養生','非塗装部養生',1,70000],
    ['外壁塗装','下塗り、上塗り（1回目）、上塗り（2回目） 計 3回塗り',297.3,2100],
    ['出窓天端','錆止め、上塗り（1回目）、上塗り（2回目） 計 3回塗り W1800×D250',3,3000],
    ['フード','錆止め、上塗り（1回目）、上塗り（2回目） 計 3回塗り W400×H1200×D300',1,2500],
    ['ベランダタラップ','錆止め、上塗り（1回目）、上塗り（2回目） 計 3回塗り W400×H4700',1,8000],
    ['目隠し枠','錆止め、上塗り（1回目）、上塗り（2回目） 計 3回塗り W1300×H600',1,3000],
    ['電力盤','錆止め、上塗り（1回目）、上塗り（2回目） 計 3回塗り W700×H1400×D200 ※閉切施工',1,3500],
    ['軒天ボード','上塗り（1回目）、上塗り（2回目） 計 2回塗り ベランダ軒 8.4㎡',1,20000],
    ['塩ビパイプ','上塗り（1回目）、上塗り（2回目） 計 2回塗り 径75mm',42,900],
    ['フード','上塗り（1回目）、上塗り（2回目） 計 2回塗り',1,1000]
  ];
  var data4 = [
    ['下地清掃','高圧水洗浄',1,20000],
    ['下地処理','緩衝目地処理、クラック処理、入隅シール',1,30000],
    ['笠木シール処理','アルミ手摺支柱根元、笠木継ぎ目部シール処理',1,60000],
    ['室外機処理','エアコン室外機吊り上げ、洗濯機移動共',1,40000],
    ['ウレタン塗膜防水','プライマー、ウレタン防水材×２、トップコート',1,150000]
  ];
  var data5 = [
    ['支持金具交換','ステンレス製 不良部取付足部含む 80A×20ケ、60A×1ケ、50A×2ケ、40A×１ケ',24,3200],
    ['目隠し波板交換','廊下波板 ポリカーボネート 2尺×2枚',1,40000],
    ['発生廃材処分費','中間処分・運搬',1,40000]
  ];

  function fillSheet(sheetName, rows){
    var sh = ss.getSheetByName(sheetName);
    for (var r=0;r<rows.length;r++){
      var row = 5+r;
      sh.getRange('B'+row).setValue(rows[r][0]);
      sh.getRange('E'+row).setValue(rows[r][1]);
      sh.getRange('I'+row).setValue(rows[r][2]);
      sh.getRange('K'+row).setValue(rows[r][3]);
    }
  }

  fillSheet('内訳書1', data1);
  fillSheet('内訳書2', data2);
  fillSheet('内訳書3', data3);
  fillSheet('内訳書4', data4);
  fillSheet('内訳書5', data5);

  Logger.log(copy.getUrl());
  return copy.getUrl();
}
