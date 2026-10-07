/* ══════════════════════════════════════════════════════════
 * 📌 このファイルは【見積書請求書アプリ】用です（他プロジェクトと貼り間違え注意）
 * ------------------------------------------------------------
 * 反映先スプレッドシート: 「見積書テンプレート（母艦）」
 *   - スプレッドシートID: 18XrMR3ZdcuvpS-dmkvF_LArDsOWaUHrF2IK6P8SHb-w
 * 本番ポータルURL（Portal.htmlのWebアプリ）:
 *   https://script.google.com/macros/s/AKfycbwD1wDvA9kSehumZjIJ-N3jy3EnfCz9D9PRkjLPVFTG017TabDdr3LWOIlHv5wchJXB/exec
 * 反映方法(2026-08-30〜): clasp自動push環境を構築済み。
 *   スクリプトID: 1oNSHLRwW-1skgjj4OFu84SNPEf2vuEjV5FZMEDRK847uW1CKMpGFX4uP
 *   push用ローカルclone: 会社PC C:\Users\sugaw\repos\gas-mitsumori-master / 自宅PC C:\Users\user\repos\gas-mitsumori-master
 *   このファイル(コード.gs.txt)を編集したら、上記cloneの「コード.js」へ上書きコピー→
 *   `clasp push` → `clasp version "説明"` → `clasp redeploy <デプロイID> -V <バージョン番号>` で反映する。
 *   本番デプロイIDは `clasp list-deployments` で確認できる(上記ポータルURLに対応するもの)。
 * ══════════════════════════════════════════════════════════ */
/**
 * 見積書メーカー（母艦）用 Apps Script　※1案件1ファイル方式
 * このスプレッドシートは「見積一覧」の管理専用です。
 * 見積の中身（表紙・内訳書）は、案件ごとに別ファイル（見積ひな形.xlsx から複製）として作られます。
 */

const TEMPLATE_FILE_ID = '1JBKTU9zgI8c_329JDYtPcVSxUeHj8OOm8ge680_acYk';
const TARGET_FOLDER_ID = '1TYqtqHTMpmtAHRjC-CoJMCP5rkR0lvWU';
const INDEX_SHEET = '見積一覧';
const DELETE_LOG_SHEET = '削除ログ';
const TRASH_FOLDER_NAME = '削除済み';
const COVER_SHEET = '見積書';
const QUOTE_NO_CELL = 'L2';
const QUOTE_DATE_CELL = 'L3';
const INVOICE_FOLDER_ID = '1irzsO6gwQdZ9FMxXGTup4ke5-ruxUiIv'; const INVOICE_SHEET = '請求書'; const INVOICE_TEMPLATE_FILE_ID = '1frPDCeGGF7SUh1uV9IbxU-ES8QmMN4aZWXyPP8vfNDc';
const INVOICE_LIST_SS_ID = '1pQ8DCcWViX0Qc7SaAQYQeodaH1c5jcgGZl4rFL5TImE';
// 自動同期の間隔（分）。2026-09-16に5→10へ。変えたらメニュー「自動同期の間隔を直す」を1回押すこと
const AUTO_SYNC_MINUTES = 10;
// 実行時間を記録する「運用モニター」スプレッドシート（システム改善計画 Step 0）
const MONITOR_SS_ID = '11iJLKYnfFc4UYj0L5wR3v7G8E4GLEj6vmqrbIF5NOp8';
function onOpen(){var ui=SpreadsheetApp.getUi();ui.createMenu('見積書メーカー').addItem('新しい見積を作成','createNewQuote').addItem('選択した見積を複製','duplicateSelectedQuote').addItem('見積一覧を開く','openIndexSheet').addSeparator().addItem('見積一覧を作り直す（フォルダ内を再スキャン）','rebuildIndex').addSeparator().addItem('見積一覧を今すぐ同期','syncRecentQuotes').addItem('同期エラーのみ再試行','retryFailedQuotes').addItem('自動同期を有効にする（初回のみ）','ensureAutoSyncTrigger').addItem('自動同期の間隔を直す','resetAutoSyncTrigger').addSeparator().addItem('使い方シートを作り直す','rebuildUsageSheet').addToUi();ui.createMenu('請求書メーカー').addItem('新しい請求書を作成（見積なし）','createNewInvoice').addItem('選択した見積から請求書を作成','createInvoiceFromSelection').addToUi();}
function openIndexSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureIndexSheet_(ss);
  ss.setActiveSheet(sheet);
}

function ensureIndexSheet_(ss) {
  let sheet = ss.getSheetByName(INDEX_SHEET);
  if (sheet) return sheet;
  sheet = ss.insertSheet(INDEX_SHEET);
  const headers = ['見積NO','作成日','宛先','工事名称','担当部署','担当者','開く','合計金額'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function validateEntryName_(name) {
  const trimmed = (name || '').trim();
  if (!trimmed) return { ok: false, message: '名前が入力されていません。' };
  if (trimmed.length < 2) return { ok: false, message: '案件名は2文字以上で入力してください（「' + trimmed + '」は短すぎます）。' };
  return { ok: true, name: trimmed };
}

function logSyncError_(ss, fileName, fileId, message) {
  let sheet = ss.getSheetByName('同期エラー');
  if (!sheet) {
    sheet = ss.insertSheet('同期エラー');
    sheet.getRange(1, 1, 1, 4).setValues([['日時', 'ファイル名', 'ファイルID', 'エラー内容']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), fileName, fileId, message]);
}

// 現在の年の「<年>年」フォルダを取得（なければ作成）
function getYearFolder_(rootFolder) {
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  const folderName = year + '年';
  const existing = rootFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return rootFolder.createFolder(folderName);
}

// rootFolder直下と、その直下のサブフォルダ（年度フォルダ）内の見積ファイルをすべて集める
function getAllQuoteFiles_(rootFolder) {
  const result = [];
  const files = rootFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) result.push(files.next());
  const subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    // 削除したものを退避するフォルダは一覧に載せない
    if (sub.getName() === TRASH_FOLDER_NAME) continue;
    const subFiles = sub.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (subFiles.hasNext()) result.push(subFiles.next());
  }
  return result;
}

// rootFolder直下と、その直下のサブフォルダ（年度フォルダ等）内のファイルをすべて集める（種類を問わない）
function getAllInvoiceFiles_(rootFolder) {
  const result = [];
  const files = rootFolder.getFiles();
  while (files.hasNext()) result.push(files.next());
  const subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    const subFiles = subFolders.next().getFiles();
    while (subFiles.hasNext()) result.push(subFiles.next());
  }
  return result;
}

// 請求書一覧（宛名・金額などの集計台帳）は先頭シートを使う
function getInvoiceListSheet_() {
  return SpreadsheetApp.openById(INVOICE_LIST_SS_ID).getSheets()[0];
}

// 請求書一覧の全行を読み取り、ポータル表示・集計用のオブジェクト配列にする
// 列構成: A宛名 B日付 C工事名称 D御請求金額 ... M開くリンク(URL文字列) N見積NO(見積から作った請求書のみ。2026-09-11〜)
function getInvoiceListRows_() {
  const sheet = getInvoiceListSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const n = lastRow - 1;
  const values = sheet.getRange(2, 1, n, 13).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (!r[0] && !r[2]) continue;
    rows.push({
      client: r[0],
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(r[1] || ''),
      jobName: r[2],
      amount: r[3],
      url: r[12] || ''
    });
  }
  return rows;
}

// 請求書作成時に、請求書一覧に1行追記する
// N列「見積NO」は見積から作った請求書のときだけ入れる（2026-09-11）。
// ★施工実績表（別のスプレッドシート）が、この見積NOで請求書と見積を確実に結びつけて実績の行を自動で足していく。
// completionDate（I列「工事完了日」）は任意。まとめ請求書（2026-09-23〜）で工事ごとに入力したものを渡す。
// ★Dateオブジェクトではなく'yyyy/MM/dd'の文字列で渡すこと。このスプレッドシートのタイムゾーン設定が
//   America/Los_Angelesのままになっており(母艦と同じ、2026-09-23確認)、Dateで渡すと表示が最大1日ずれる。
function appendInvoiceListRow_(client, dateVal, jobName, amount, url, quoteNo, completionDate) {
  const sheet = getInvoiceListSheet_();
  ensureSheetSize_(sheet, 1, 14);
  if (String(sheet.getRange(1, 14).getValue() || '') === '') sheet.getRange(1, 14).setValue('見積NO');
  const row = new Array(14).fill('');
  row[0] = client || '';
  row[1] = dateVal || new Date();
  row[2] = jobName || '';
  row[3] = (amount === 0 || amount) ? amount : '';
  row[8] = completionDate || '';
  row[12] = url || '';
  row[13] = quoteNo || '';
  sheet.appendRow(row);
}

// 'yyyy-MM-dd'（HTMLのdate入力）を、タイムゾーンずれの無いDateに変換する。不正な形式ならnull。
function parseIsoDate_(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function nextQuoteNo_() {
  // ★見積NOの採番は必ず排他制御(ScriptLock)の中で行う。
  //   ロックが無いと、2人が同時に作成/複製したときに同じ番号を読んでしまい、
  //   見積NOが重複する(2026-09-10に追加)。
  //   ScriptLockはこのスクリプトプロジェクト全体で共有されるので、
  //   ポータル(Webアプリ)からの実行と母艦メニューからの実行の両方に効く。
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('見積NOの採番が混み合っています（ほかの人が同時に作成中の可能性があります）。少し待ってから、もう一度お試しください。');
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyMMdd');
    const savedDate = props.getProperty('QN_DATE');
    let seq = 1;
    if (savedDate === today) {
      seq = parseInt(props.getProperty('QN_SEQ') || '0', 10) + 1;
    }
    props.setProperty('QN_DATE', today);
    props.setProperty('QN_SEQ', String(seq));
    return today + '-' + ('000' + seq).slice(-3);
  } finally {
    lock.releaseLock();
  }
}

function createNewQuote() {
  const ui = SpreadsheetApp.getUi();
  if (!TEMPLATE_FILE_ID || TEMPLATE_FILE_ID.indexOf('ここに') === 0) {
    ui.alert('見積ひな形ファイルのIDが未設定です。スクリプトの TEMPLATE_FILE_ID を設定してください。');
    return;
  }
  const res = ui.prompt('新しい見積を作成', '案件名（顧客名・工事名など。ファイル名になります）を入力してください', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const validation = validateEntryName_(res.getResponseText());
  if (!validation.ok) { ui.alert(validation.message); return; }
  const name = validation.name;
  let quoteNo;
  try { quoteNo = nextQuoteNo_(); } catch (e) { ui.alert(e.message); return; }
  const fileName = quoteNo + '_' + name;

  let newFile;
  try {
    const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
    if (TARGET_FOLDER_ID) {
      const rootFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
      const folder = getYearFolder_(rootFolder);
      newFile = templateFile.makeCopy(fileName, folder);
    } else {
      newFile = templateFile.makeCopy(fileName);
    }
  } catch (e) {
    ui.alert('見積ファイルの作成に失敗しました：' + e.message);
    return;
  }

  const newSs = SpreadsheetApp.openById(newFile.getId());
  const newCover = newSs.getSheetByName(COVER_SHEET);
  if (!newCover) {
    ui.alert('作成はできましたが「' + COVER_SHEET + '」シートが見つかりません。ひな形ファイルの構成をご確認ください。');
    return;
  }
  const createdDate = new Date();
  writeQuoteNoAndDate_(newCover,quoteNo,createdDate);
  writeQuoteUsageSheet_(newSs);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idx = ensureIndexSheet_(ss);
  const row = idx.getLastRow() + 1;
  idx.getRange(row, 1).setValue(quoteNo);
  idx.getRange(row, 2).setValue(createdDate);
  idx.getRange(row, 7).setFormula('=HYPERLINK("' + newFile.getUrl() + '","開く")');
idx.getRange(row, 8).setValue(newCover.getRange('C7').getValue());
  ui.alert('新しい見積ファイル「' + fileName + '」を作成しました。\n' +
           'Googleドライブ、または「見積一覧」の「開く」リンクから開いてください。\n' +
           '入力後は、その見積ファイル側で「見積一覧に登録・更新」を押してください。');
}

function syncRecentQuotes(){
  var r=syncRecentQuotes_();
  var msg='見積一覧を同期しました（更新:'+r.updated+'件 / 追加:'+r.added+'件 / 対象:'+r.scanned+'件）。';
  if(r.failed>0){
    msg+='\n⚠️ '+r.failed+'件が同期に失敗しました。詳細は「同期エラー」シートをご確認ください。次回の同期で自動的に再試行されます。';
  }
  SpreadsheetApp.getUi().alert(msg);
}
function autoSyncQuotes(){
  var t0=Date.now();
  var r=syncRecentQuotes_();
  logRunToMonitor_('autoSyncQuotes','対象'+r.scanned+'件/更新'+r.updated+'/追加'+r.added+'/失敗'+r.failed,Date.now()-t0);
}

/**
 * 「運用モニター」スプレッドシートの「実行ログ」に1行足す（2026-09-16 追加）。
 * 定期処理が1日に何分トリガー枠を使っているかを実測するためのもの。
 * 記録に失敗しても同期本体は止めない。
 */
function logRunToMonitor_(funcName,note,elapsedMs){
  try{
    SpreadsheetApp.openById(MONITOR_SS_ID).getSheetByName('実行ログ')
      .appendRow([new Date(),'gas-mitsumori-master',funcName,note||'',elapsedMs]);
  }catch(e){
    Logger.log('実行ログの記録に失敗: '+e);
  }
}

/** 24時間に1回だけ全走査する（取りこぼしの保険）。判定したらその場で時刻を記録する。 */
function isFullScanDue_(props){
  var last=props.getProperty('FULL_SCAN_TS');
  if(last&&(Date.now()-new Date(last).getTime())<24*60*60*1000)return false;
  props.setProperty('FULL_SCAN_TS',new Date().toISOString());
  return true;
}

/**
 * 「見積保存」直下と年度フォルダ（「削除済み」は除く）のうち、
 * since より後に更新されたスプレッドシートだけを返す。
 * 年度フォルダは毎年増えるので、親フォルダのIDは毎回組み立てる。
 */
function searchUpdatedQuoteFiles_(since){
  try{
    var parentIds=[TARGET_FOLDER_ID];
    var subs=DriveApp.getFolderById(TARGET_FOLDER_ID).getFolders();
    while(subs.hasNext()){
      var sub=subs.next();
      if(sub.getName()===TRASH_FOLDER_NAME)continue;
      parentIds.push(sub.getId());
    }
    var parentClause=parentIds.map(function(id){return "'"+id+"' in parents";}).join(' or ');
    var q="mimeType = 'application/vnd.google-apps.spreadsheet'"
        +" and modifiedDate > '"+since.toISOString()+"'"
        +" and trashed = false and ("+parentClause+")";
    var it=DriveApp.searchFiles(q);
    var out=[];
    while(it.hasNext())out.push(it.next());
    return out;
  }catch(e){
    // 検索が使えないときは、従来どおりの全走査に戻して取りこぼさないようにする
    Logger.log('更新分の検索に失敗したため全走査に切り替えます: '+e);
    return getAllQuoteFiles_(DriveApp.getFolderById(TARGET_FOLDER_ID));
  }
}
function syncRecentQuotes_(){
  if(!TARGET_FOLDER_ID)return{scanned:0,updated:0,added:0,failed:0};
  var props=PropertiesService.getScriptProperties();
  var lastSyncStr=props.getProperty('LAST_SYNC_TS');
  var lastSync=lastSyncStr?new Date(lastSyncStr):new Date(Date.now()-7*24*60*60*1000);
  var syncStartedAt=new Date();
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var idx=ensureIndexSheet_(ss);

  var retryIds=JSON.parse(props.getProperty('SYNC_RETRY_IDS')||'[]');
  var seenIds={};
  var targets=[];

  // 2026-09-16: 毎回フォルダを全走査していたのをやめ、前回同期以降に更新されたファイルだけを
  // Drive検索で拾うようにした（見積が何年分たまっても1回あたりの時間が伸びないようにするため）。
  // ただし「中身は更新されずフォルダ間で移動だけされた」ファイルは検索に出てこないので、
  // 24時間に1回だけ従来どおりの全走査を行う（取りこぼしの保険）。
  var fullScan=isFullScanDue_(props);
  var candidates=fullScan
    ? getAllQuoteFiles_(DriveApp.getFolderById(TARGET_FOLDER_ID))
    : searchUpdatedQuoteFiles_(lastSync);
  candidates.forEach(function(f){
    if(f.getId()===TEMPLATE_FILE_ID)return;
    if(f.getLastUpdated()<=lastSync)return;
    if(seenIds[f.getId()])return;
    targets.push(f);
    seenIds[f.getId()]=true;
  });
  retryIds.forEach(function(id){
    if(seenIds[id])return;
    try{
      targets.push(DriveApp.getFileById(id));
      seenIds[id]=true;
    }catch(e){}
  });

  var data=idx.getDataRange().getValues();
  var rowByQuoteNo={};
  for(var i=1;i<data.length;i++){var qn=data[i][0];if(qn)rowByQuoteNo[String(qn)]=i+1;}

  var scanned=0,updated=0,added=0,failed=0;
  var stillFailedIds=[];

  targets.forEach(function(f){
    scanned++;
    try{
      var fss=SpreadsheetApp.openById(f.getId());
      var cover=fss.getSheetByName(COVER_SHEET);
      if(!cover)return;
      var cells=coverCells_(cover);
      var quoteNo=cover.getRange(cells.no).getValue();
      if(!quoteNo)return;
      var quoteDate=cover.getRange(cells.date).getValue();
      var client=cover.getRange('B5').getValue();
      var jobName=cover.getRange('C9').getValue();
      var department=cover.getRange('J10').getValue();
      var person=cover.getRange('L10').getValue();
      var row=rowByQuoteNo[String(quoteNo)];
      if(!row){row=idx.getLastRow()+1;rowByQuoteNo[String(quoteNo)]=row;added++;}else{updated++;}
      idx.getRange(row,1).setValue(quoteNo);
      idx.getRange(row,2).setValue(quoteDate);
      idx.getRange(row,3).setValue(client);
      idx.getRange(row,4).setValue(jobName);
      idx.getRange(row,5).setValue(department);
      idx.getRange(row,6).setValue(person);
      idx.getRange(row,7).setFormula('=HYPERLINK("'+f.getUrl()+'","開く")');
      idx.getRange(row,8).setValue(cover.getRange('C7').getValue());
    }catch(e){
      failed++;
      stillFailedIds.push(f.getId());
      var fileName='(取得失敗)';
      try{fileName=f.getName();}catch(e2){}
      logSyncError_(ss,fileName,f.getId(),String(e.message||e));
    }
  });

  props.setProperty('SYNC_RETRY_IDS',JSON.stringify(stillFailedIds));
  props.setProperty('LAST_SYNC_TS',syncStartedAt.toISOString());
  return{scanned:scanned,updated:updated,added:added,failed:failed};
}
function retryFailedQuotes(){
  var props=PropertiesService.getScriptProperties();
  var retryIds=JSON.parse(props.getProperty('SYNC_RETRY_IDS')||'[]');
  if(retryIds.length===0){
    SpreadsheetApp.getUi().alert('現在、再試行が必要な同期エラーはありません。');
    return;
  }
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var idx=ensureIndexSheet_(ss);
  var data=idx.getDataRange().getValues();
  var rowByQuoteNo={};
  for(var i=1;i<data.length;i++){var qn=data[i][0];if(qn)rowByQuoteNo[String(qn)]=i+1;}

  var fixed=0,stillFailed=0;
  var stillFailedIds=[];

  retryIds.forEach(function(id){
    var f;
    try{
      f=DriveApp.getFileById(id);
    }catch(e){
      stillFailed++;
      stillFailedIds.push(id);
      logSyncError_(ss,'(ID:'+id+')',id,'ファイルを開けません：'+String(e.message||e));
      return;
    }
    try{
      var fss=SpreadsheetApp.openById(f.getId());
      var cover=fss.getSheetByName(COVER_SHEET);
      if(!cover)throw new Error('「'+COVER_SHEET+'」シートが見つかりません');
      var cells=coverCells_(cover);
      var quoteNo=cover.getRange(cells.no).getValue();
      if(!quoteNo)throw new Error('見積NOが空です');
      var quoteDate=cover.getRange(cells.date).getValue();
      var client=cover.getRange('B5').getValue();
      var jobName=cover.getRange('C9').getValue();
      var department=cover.getRange('J10').getValue();
      var person=cover.getRange('L10').getValue();
      var row=rowByQuoteNo[String(quoteNo)];
      if(!row){row=idx.getLastRow()+1;rowByQuoteNo[String(quoteNo)]=row;}
      idx.getRange(row,1).setValue(quoteNo);
      idx.getRange(row,2).setValue(quoteDate);
      idx.getRange(row,3).setValue(client);
      idx.getRange(row,4).setValue(jobName);
      idx.getRange(row,5).setValue(department);
      idx.getRange(row,6).setValue(person);
      idx.getRange(row,7).setFormula('=HYPERLINK("'+f.getUrl()+'","開く")');
      idx.getRange(row,8).setValue(cover.getRange('C7').getValue());
      fixed++;
    }catch(e){
      stillFailed++;
      stillFailedIds.push(id);
      logSyncError_(ss,f.getName(),id,String(e.message||e));
    }
  });

  props.setProperty('SYNC_RETRY_IDS',JSON.stringify(stillFailedIds));
  var msg='同期エラーの再試行が完了しました（復旧:'+fixed+'件 / 対象:'+retryIds.length+'件）。';
  if(stillFailed>0)msg+='\n⚠️ '+stillFailed+'件はまだ失敗しています。「同期エラー」シートをご確認ください。';
  SpreadsheetApp.getUi().alert(msg);
}
function ensureAutoSyncTrigger(){var triggers=ScriptApp.getProjectTriggers();for(var i=0;i<triggers.length;i++){if(triggers[i].getHandlerFunction()==='autoSyncQuotes'){SpreadsheetApp.getUi().alert('自動同期はすでに有効です（'+AUTO_SYNC_MINUTES+'分ごと）。\n間隔を作り直したいときは「自動同期の間隔を直す」を押してください。');return;}}ScriptApp.newTrigger('autoSyncQuotes').timeBased().everyMinutes(AUTO_SYNC_MINUTES).create();SpreadsheetApp.getUi().alert('自動同期を有効にしました。今後'+AUTO_SYNC_MINUTES+'分ごとに見積一覧が自動更新されます。');}

/**
 * 自動同期のトリガーを作り直す（2026-09-16 追加）。
 * 間隔を変えてもトリガーは古いままなので、いったん消してから登録し直す。
 */
function resetAutoSyncTrigger(){
  var triggers=ScriptApp.getProjectTriggers();
  var removed=0;
  for(var i=0;i<triggers.length;i++){
    if(triggers[i].getHandlerFunction()==='autoSyncQuotes'){
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  ScriptApp.newTrigger('autoSyncQuotes').timeBased().everyMinutes(AUTO_SYNC_MINUTES).create();
  SpreadsheetApp.getUi().alert('自動同期のトリガーを作り直しました（'+AUTO_SYNC_MINUTES+'分ごと）。\n'
    +'古いトリガー：'+removed+'件を削除しました。\n\n'
    +'※見積ファイル側の「見積一覧に登録・更新」と、メニューの「見積一覧を今すぐ同期」は今までどおり使えます。');
}function rebuildIndex() {
  const ui = SpreadsheetApp.getUi();
  if (!TARGET_FOLDER_ID) {
    ui.alert('この機能を使うには、保存先フォルダ（TARGET_FOLDER_ID）の設定が必要です。');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idx = ensureIndexSheet_(ss);

  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const rows = [];
  let failed = 0;
  getAllQuoteFiles_(folder).forEach(function(f) {
    if (f.getId() === TEMPLATE_FILE_ID) return;
    try {
      const fss = SpreadsheetApp.openById(f.getId());
      const cover = fss.getSheetByName(COVER_SHEET);
      if (!cover) return;
      const cells = coverCells_(cover);
      rows.push([
        cover.getRange(cells.no).getValue(),
        cover.getRange(cells.date).getValue(),
        cover.getRange('B5').getValue(),
        cover.getRange('C9').getValue(),
        cover.getRange('J10').getValue(),
        cover.getRange('L10').getValue(),
        '=HYPERLINK("' + f.getUrl() + '","開く")',
        cover.getRange('C7').getValue()
      ]);
    } catch (e) {
      failed++;
      logSyncError_(ss, f.getName(), f.getId(), String(e.message || e));
    }
  });

  const lastRow = idx.getLastRow();
  if (lastRow > 1) idx.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (rows.length > 0) idx.getRange(2, 1, rows.length, 8).setValues(rows);

  let msg = '見積一覧を作り直しました（' + rows.length + '件）。';
  if (failed > 0) msg += '\n⚠️ ' + failed + '件の読み込みに失敗しました。「同期エラー」シートをご確認ください。';
  ui.alert(msg);
}function setupInvoiceTemplate_(){var ss2=SpreadsheetApp.create('請求書ひな形');var file2=DriveApp.getFileById(ss2.getId());var folder=DriveApp.getFolderById(INVOICE_FOLDER_ID);folder.addFile(file2);try{DriveApp.getRootFolder().removeFile(file2);}catch(e){}var templateSs=SpreadsheetApp.openById(TEMPLATE_FILE_ID);var srcCover=templateSs.getSheetByName(COVER_SHEET);var newSheet=srcCover.copyTo(ss2);newSheet.setName(INVOICE_SHEET);var defaultSheet=ss2.getSheetByName('シート1');if(defaultSheet)ss2.deleteSheet(defaultSheet);newSheet.getRange('C1').setValue('御 請 求 書');newSheet.getRange('K2').setValue('請求書NO.');newSheet.getRange('K3').setValue('発行日');newSheet.getRange('K4').setValue('');newSheet.getRange('B7').setValue('御請求金額');newSheet.getRange('C7').setValue(0);newSheet.deleteRows(11,20);SpreadsheetApp.getUi().alert('請求書ひな形を作成しました。このIDをINVOICE_TEMPLATE_FILE_IDに設定してください: '+ss2.getId());}function nextInvoiceNo_(){var props=PropertiesService.getScriptProperties();var tz=Session.getScriptTimeZone();var today=Utilities.formatDate(new Date(),tz,'yyMMdd');var savedDate=props.getProperty('IN_DATE');var seq=1;if(savedDate===today){seq=parseInt(props.getProperty('IN_SEQ')||'0',10)+1;}props.setProperty('IN_DATE',today);props.setProperty('IN_SEQ',String(seq));return 'INV-'+today+'-'+('000'+seq).slice(-3);}function createInvoiceFromSelection(){var ui=SpreadsheetApp.getUi();if(!INVOICE_TEMPLATE_FILE_ID){ui.alert('請求書ひな形が未設定です。INVOICE_TEMPLATE_FILE_ID を設定してください。');return;}var ss=SpreadsheetApp.getActiveSpreadsheet();var sheet=ss.getActiveSheet();if(sheet.getName()!==INDEX_SHEET){ui.alert('「'+INDEX_SHEET+'」シートで、請求書を作成したい見積の行を選択してから実行してください。');return;}var row=sheet.getActiveCell().getRow();if(row<2){ui.alert('見積の行を選択してから実行してください。');return;}var data=sheet.getRange(row,1,1,8).getValues()[0];var quoteNo=data[0],client=data[2],jobName=data[3],dept=data[4],person=data[5],total=data[7];if(!quoteNo){ui.alert('この行には見積データがありません。');return;}var res=ui.alert('確認','見積NO「'+quoteNo+'」（'+client+'）から請求書を作成します。よろしいですか？',ui.ButtonSet.OK_CANCEL);if(res!==ui.Button.OK)return;try{var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);var folder=getYearFolder_(rootFolder);var newFile=templateFile.makeCopy('請求書_'+String(client)+'_'+String(jobName),folder);var newSs=SpreadsheetApp.openById(newFile.getId());var cover=newSs.getSheetByName(INVOICE_SHEET);var invoiceNo=nextInvoiceNo_();cover.getRange('L2').setValue(invoiceNo);cover.getRange('L3').setValue(new Date());cover.getRange('B5').setValue(client);cover.getRange('C9').setValue(jobName);cover.getRange('C7').setValue(total);cover.getRange('J10').setValue(dept);cover.getRange('L10').setValue(person);appendInvoiceListRow_(client,new Date(),jobName,total,newFile.getUrl(),quoteNo);ui.alert('請求書を作成しました。\n'+newFile.getUrl());}catch(e){ui.alert('請求書の作成に失敗しました：'+e.message);}}
function createNewInvoice(){var ui=SpreadsheetApp.getUi();if(!INVOICE_TEMPLATE_FILE_ID){ui.alert('請求書ひな形が未設定です。INVOICE_TEMPLATE_FILE_ID を設定してください。');return;}var res=ui.prompt('新しい請求書を作成','案件名（顧客名・工事名など。ファイル名になります）を入力してください',ui.ButtonSet.OK_CANCEL);if(res.getSelectedButton()!==ui.Button.OK)return;var name=res.getResponseText();if(!name||!name.trim()){ui.alert('名前が入力されていません。');return;}var newFile;try{var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);var folder=getYearFolder_(rootFolder);newFile=templateFile.makeCopy('請求書_'+name.trim(),folder);}catch(e){ui.alert('請求書ファイルの作成に失敗しました：'+e.message);return;}var newSs=SpreadsheetApp.openById(newFile.getId());var cover=newSs.getSheetByName(INVOICE_SHEET);if(!cover){ui.alert('作成はできましたが「'+INVOICE_SHEET+'」シートが見つかりません。ひな形ファイルの構成をご確認ください。');return;}var invoiceNo=nextInvoiceNo_();cover.getRange('L2').setValue(invoiceNo);cover.getRange('L3').setValue(new Date());appendInvoiceListRow_('',new Date(),name.trim(),'',newFile.getUrl());ui.alert('新しい請求書ファイル「'+name.trim()+'」を作成しました。\n'+newFile.getUrl()+'\n\n宛先・工事名称・金額などは、開いたファイル内に直接入力してください。');}
const MOTHER_SS_ID='18XrMR3ZdcuvpS-dmkvF_LArDsOWaUHrF2IK6P8SHb-w';
function doGet(e){return HtmlService.createTemplateFromFile('Portal').evaluate().setTitle('見積・請求ポータル').addMetaTag('viewport','width=device-width, initial-scale=1');}

// マニュアルPDFの更新専用。管理者がマニュアルを差し替えたいときに、
// このURLへ{action:'uploadManual', data:'<base64のPDF>'}をPOSTする(ポータル画面からは使わない)。
function doPost(e){
  try{
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'uploadManual') {
      return ContentService.createTextOutput(JSON.stringify(uploadManualPdf_(data.data))).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({success:false,message:'不明なactionです。'})).setMimeType(ContentService.MimeType.JSON);
  }catch(e){
    return ContentService.createTextOutput(JSON.stringify({success:false,message:e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
function getPortalData(){var ss=SpreadsheetApp.openById(MOTHER_SS_ID);var idx=ss.getSheetByName(INDEX_SHEET);var quotes=[];if(idx){var lastRow=idx.getLastRow();if(lastRow>=2){var n=lastRow-1;var values=idx.getRange(2,1,n,8).getValues();var formulas=idx.getRange(2,7,n,1).getFormulas();for(var i=0;i<values.length;i++){var row=values[i];if(!row[0])continue;var url='';var f=formulas[i][0];var m=f&&f.match(/HYPERLINK\("([^"]+)"/);if(m)url=m[1];quotes.push({sheetRow:i+2,id:extractFileId_(url),no:row[0],date:row[1] instanceof Date?Utilities.formatDate(row[1],Session.getScriptTimeZone(),'yyyy/MM/dd'):String(row[1]||''),client:row[2],jobName:row[3],dept:row[4],person:row[5],total:row[7],url:url});}}}quotes.sort(function(a,b){return a.date<b.date?1:(a.date>b.date?-1:0);});var invoices=getInvoiceListRows_();invoices.sort(function(a,b){return a.date<b.date?1:(a.date>b.date?-1:0);});return{quotes:quotes,invoices:invoices};}
/**
 * 見積ファイルのURLからファイルIDを取り出す。
 */
/**
 * 「削除ログ」シート。見積一覧から削除したものの控え。
 * ★ここは記録だけで、一覧の中身には影響しない。
 *   削除したファイルは「ゴミ箱」か「見積保存 ＞ 削除済み」フォルダへ移すので、
 *   走査の対象から外れて一覧に戻らない（2026-09-10）。
 */
function ensureDeleteLogSheet_(ss){
  var sheet=ss.getSheetByName(DELETE_LOG_SHEET);
  if(sheet)return sheet;
  sheet=ss.insertSheet(DELETE_LOG_SHEET);
  sheet.getRange(1,1,1,8).setValues([['削除した日時','見積NO','宛先','工事名','ファイルの行き先','ファイル名','ファイルID','削除した人']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,150);
  sheet.setColumnWidth(3,180);
  sheet.setColumnWidth(4,240);
  sheet.setColumnWidth(5,260);
  sheet.setColumnWidth(6,240);
  sheet.setColumnWidth(7,290);
  return sheet;
}

/**
 * 「見積保存 ＞ 削除済み」フォルダを返す（無ければ作る）。
 * ★他の人が所有している見積ファイルは、こちらの権限ではゴミ箱に入れられない
 *   （Googleドライブは所有者しかゴミ箱に入れられない）ため、その受け皿。
 *   getAllQuoteFiles_ はこのフォルダを走査しないので、一覧には戻らない。
 */
function getDeletedFolder_(rootFolder){
  var it=rootFolder.getFoldersByName(TRASH_FOLDER_NAME);
  if(it.hasNext())return it.next();
  return rootFolder.createFolder(TRASH_FOLDER_NAME);
}

/**
 * 見積一覧から行を消し、ドライブの見積ファイルもゴミ箱へ移す。
 * items: [{sheetRow:見積一覧の行番号, no:見積NO, id:ファイルID}]
 * ・ゴミ箱に入れられないファイル（他の人が所有）は「見積保存 ＞ 削除済み」フォルダへ移す。
 * ・同じファイルIDの行が一覧に2行以上ある（重複行）ときは、行だけ消してファイルには触らない。
 */
function removeQuotesFromIndexWeb(items){
  if(!items||!items.length)return{success:false,message:'削除する行が選ばれていません。'};
  var lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
  }catch(e){
    return{success:false,message:'ほかの処理と重なっています。少し待ってから、もう一度お試しください。'};
  }
  try{
    var ss=SpreadsheetApp.openById(MOTHER_SS_ID);
    var idx=ensureIndexSheet_(ss);
    var logSheet=ensureDeleteLogSheet_(ss);
    var lastRow=idx.getLastRow();
    var idCount={};
    if(lastRow>=2){
      idx.getRange(2,7,lastRow-1,1).getFormulas().forEach(function(f){
        var m=f[0]&&f[0].match(/HYPERLINK\("([^"]+)"/);
        var id=m?extractFileId_(m[1]):'';
        if(id)idCount[id]=(idCount[id]||0)+1;
      });
    }
    var who='';
    try{who=Session.getActiveUser().getEmail()||'';}catch(e){}
    var now=new Date();
    var deletedFolder=null;
    var trashed=0,moved=0,dupOnly=0,noFile=0,gone=0,skipped=0,failed=0;
    var failedNames=[];
    // 下の行から消す（消すと行番号がずれるため）
    items.slice().sort(function(a,b){return Number(b.sheetRow)-Number(a.sheetRow);}).forEach(function(it){
      var r=Number(it.sheetRow);
      if(!(r>=2)||r>idx.getLastRow()){skipped++;return;}
      var vals=idx.getRange(r,1,1,8).getValues()[0];
      // 一覧が動いていたら触らない（別の見積を消してしまわないため）
      if(String(vals[0])!==String(it.no)){skipped++;return;}
      var id=String(it.id||'');
      var where='',fileName='';
      if(!id){
        where='ファイルへのリンクなし（行のみ削除）';
        noFile++;
      }else if(idCount[id]>1){
        where='重複行のため行のみ削除（ファイルはそのまま）';
        idCount[id]--;
        dupOnly++;
      }else{
        var file=null;
        try{
          file=DriveApp.getFileById(id);
          fileName=file.getName();
        }catch(e){
          // ファイルがもう無い（以前に手で消された等）。一覧に残骸の行だけが残っている状態なので、行は消す。
          file=null;
        }
        if(!file){
          where='ドライブにファイルが見つからない（行のみ削除）';
          gone++;
        }else{
          var trashErr='';
          try{
            file.setTrashed(true);
          }catch(e){
            trashErr=String(e&&e.message||e);
          }
          var isTrashed=false;
          try{isTrashed=file.isTrashed();}catch(e){}
          if(!trashErr&&isTrashed){
            where='ゴミ箱';
            trashed++;
          }else{
            // 他の人が所有しているとゴミ箱に入れられないので、「削除済み」フォルダへ退避する
            try{
              if(!deletedFolder)deletedFolder=getDeletedFolder_(DriveApp.getFolderById(TARGET_FOLDER_ID));
              file.moveTo(deletedFolder);
              where='「'+TRASH_FOLDER_NAME+'」フォルダ（他の人が所有しているためゴミ箱に入れられません）';
              moved++;
            }catch(e2){
              failed++;
              failedNames.push(String(vals[0])+'：'+(trashErr||String(e2&&e2.message||e2)));
              return;
            }
          }
        }
      }
      logSheet.appendRow([now,vals[0],vals[2],vals[3],where,fileName,id,who]);
      idx.deleteRow(r);
    });
    var removed=trashed+moved+dupOnly+noFile+gone;
    var parts=[];
    if(trashed>0)parts.push('ファイル'+trashed+'件をゴミ箱へ移動');
    if(moved>0)parts.push('ファイル'+moved+'件を「'+TRASH_FOLDER_NAME+'」フォルダへ移動（他の人が所有しているためゴミ箱に入れられません）');
    if(dupOnly>0)parts.push(dupOnly+'件は同じファイルの重複行のため行のみ削除');
    if(noFile>0)parts.push(noFile+'件はファイルへのリンクが無いため行のみ削除');
    if(gone>0)parts.push(gone+'件はドライブにファイルが無かったため行のみ削除');
    var msg='見積一覧から'+removed+'件を削除しました。'+(parts.length?'（'+parts.join(' / ')+'）':'');
    if(failed>0)msg+=' ※'+failed+'件はファイルを移動できなかったため行を残しました（'+failedNames.join(' / ')+'）。';
    if(skipped>0)msg+=' ※'+skipped+'件は一覧が更新されていたため消していません。画面を再読み込みしてからやり直してください。';
    if(removed===0)msg='削除できた行がありませんでした。'+(failed>0||skipped>0?msg.replace(/^見積一覧から0件を削除しました。/,''):'');
    return{success:(removed>0),removed:removed,failed:failed,skipped:skipped,message:msg};
  }finally{
    lock.releaseLock();
  }
}

/**
 * 見積NOらしい値か。新しい形式 yymmdd-nnn（例 260916-001）と、
 * 2026年7月の移行前に使っていた8桁（例 25090115）を許す。
 */
function looksLikeQuoteNo_(v){
  var s=String(v==null?'':v).trim();
  return /^\d{6}-\d{3}$/.test(s) || /^\d{8}$/.test(s);
}

/**
 * 表紙の「見積NO」「見積作成日」が入っているセルを見分ける。
 * ★2026年7月より前に作られた見積ファイルは、表紙の右上ブロックが1列ぶん左にずれていて、
 *   見積NOがK2・作成日がK3にある（新しいものはL2・L3で、K2は見出し「見積NO.」）。
 *   見分けは「見積NO.」の見出しがK2にあるか（新）J2にあるか（旧）で行う。
 *   ここを見ないと、古いレイアウトの見積はL2が空なので
 *   syncRecentQuotes_ が `if(!quoteNo)return;` で丸ごと読み飛ばし、
 *   中身を直しても見積一覧に永久に反映されない（2026-09-16に260916-001で判明）。
 */
function coverCells_(cover){
  var NEW={no:QUOTE_NO_CELL,date:QUOTE_DATE_CELL,old:false};
  var OLD={no:'K2',date:'K3',old:true};
  try{
    var k2=String(cover.getRange('K2').getValue()||'').replace(/[\s　]/g,'');
    if(k2.indexOf('見積NO')===0)return NEW;
    var j2=String(cover.getRange('J2').getValue()||'').replace(/[\s　]/g,'');
    if(j2.indexOf('見積NO')===0)return OLD;
    // 見出しで判別できないときは、値の形で決める
    if(looksLikeQuoteNo_(cover.getRange(QUOTE_NO_CELL).getValue()))return NEW;
    if(looksLikeQuoteNo_(cover.getRange('K2').getValue()))return OLD;
  }catch(e){}
  return NEW;
}

/**
 * 表紙に見積NOと作成日を書く（古いレイアウトにも対応）。
 * 古い見積を複製したとき、L2に書いても表紙に見えるのはK2なので、
 * 複製元の古い番号が表紙に残ったままになってしまう。
 * 使わない方のセルに見積NOらしい値が残っていたら消す（同期がそちらを読むのを防ぐ）。
 */
function writeQuoteNoAndDate_(cover,quoteNo,date){
  var c=coverCells_(cover);
  cover.getRange(c.no).setValue(quoteNo);
  cover.getRange(c.date).setValue(date);
  var other=c.old?QUOTE_NO_CELL:'K2';
  try{
    if(looksLikeQuoteNo_(cover.getRange(other).getValue()))cover.getRange(other).clearContent();
  }catch(e){}
  return c;
}

function extractFileId_(url){
  var m=String(url||'').match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  return m?m[1]:'';
}

/**
 * 既存の見積ファイルを複製して、新しい見積NOを振る。
 * ★複製の処理は「母艦」側に置いてある。見積ファイル1件ずつに埋め込まれたスクリプトは
 *   ファイル作成時のコピーで、直しても既存ファイルには伝播しないため（2026-09-10）。
 */
function duplicateQuoteWeb(fileId,name){
  if(!fileId)return{success:false,message:'複製元の見積ファイルが特定できません。見積一覧の「開く」リンクをご確認ください。'};
  var validation=validateEntryName_(name);
  if(!validation.ok)return{success:false,message:validation.message};
  if(!TARGET_FOLDER_ID)return{success:false,message:'見積保存フォルダのIDが未設定です。'};
  var quoteNo;
  try{quoteNo=nextQuoteNo_();}catch(e){return{success:false,message:e.message};}
  var fileName=quoteNo+'_'+validation.name;
  var newFile;
  try{
    var srcFile=DriveApp.getFileById(fileId);
    var rootFolder=DriveApp.getFolderById(TARGET_FOLDER_ID);
    var folder=getYearFolder_(rootFolder);
    newFile=srcFile.makeCopy(fileName,folder);
  }catch(e){
    return{success:false,message:'見積ファイルの複製に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var newCover=newSs.getSheetByName(COVER_SHEET);
  if(!newCover)return{success:false,message:'複製はできましたが「'+COVER_SHEET+'」シートが見つかりません。',url:newFile.getUrl()};
  var createdDate=new Date();
  writeQuoteNoAndDate_(newCover,quoteNo,createdDate);
  writeQuoteUsageSheet_(newSs);
  var ss=SpreadsheetApp.openById(MOTHER_SS_ID);
  var idx=ensureIndexSheet_(ss);
  var row=idx.getLastRow()+1;
  idx.getRange(row,1).setValue(quoteNo);
  idx.getRange(row,2).setValue(createdDate);
  idx.getRange(row,3).setValue(newCover.getRange('B5').getValue());
  idx.getRange(row,4).setValue(newCover.getRange('C9').getValue());
  idx.getRange(row,5).setValue(newCover.getRange('J10').getValue());
  idx.getRange(row,6).setValue(newCover.getRange('L10').getValue());
  idx.getRange(row,7).setFormula('=HYPERLINK("'+newFile.getUrl()+'","開く")');
  idx.getRange(row,8).setValue(newCover.getRange('C7').getValue());
  return{success:true,url:newFile.getUrl(),no:quoteNo,message:'見積NO「'+quoteNo+'」として複製しました（ファイル名：'+fileName+'）。'};
}

/**
 * 母艦の「見積一覧」で選択中の行の見積を複製する（メニュー用）。
 */
function duplicateSelectedQuote(){
  var ui=SpreadsheetApp.getUi();
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  if(sheet.getName()!==INDEX_SHEET){ui.alert('「'+INDEX_SHEET+'」シートで、複製したい見積の行を選択してから実行してください。');return;}
  var row=sheet.getActiveCell().getRow();
  if(row<2){ui.alert('見積の行を選択してから実行してください。');return;}
  var srcNo=sheet.getRange(row,1).getValue();
  var srcClient=sheet.getRange(row,3).getValue();
  var srcJobName=sheet.getRange(row,4).getValue();
  if(!srcNo){ui.alert('この行には見積データがありません。');return;}
  var formula=sheet.getRange(row,7).getFormula();
  var m=formula&&formula.match(/HYPERLINK\("([^"]+)"/);
  var fileId=m?extractFileId_(m[1]):'';
  if(!fileId){ui.alert('この行には見積ファイルへのリンクがありません。「見積一覧を作り直す（フォルダ内を再スキャン）」をお試しください。');return;}
  var res=ui.prompt('見積を複製',
    '複製元：'+srcNo+'（'+srcClient+' / '+srcJobName+'）\n\n新しい案件名（ファイル名になります）を入力してください',
    ui.ButtonSet.OK_CANCEL);
  if(res.getSelectedButton()!==ui.Button.OK)return;
  var r=duplicateQuoteWeb(fileId,res.getResponseText());
  if(!r.success){ui.alert(r.message);return;}
  ui.alert(r.message+'\n'+r.url+'\n\n中身は複製元のままです。金額などを直してからお使いください。');
}

/**
 * 「使い方」シートに書き出す文面。
 * ★アプリの機能・メニューを変えたら、ここも直して
 *   母艦メニュー「見積書メーカー」→「使い方シートを作り直す」を1回押すこと。
 *   （2026-09-10: 手書きのシートが機能変更から取り残されていたため、コード生成に切り替えた）
 */
/**
 * 見積ファイル側の「使い方」シート（ひな形から複製されて1件ずつ入っている）に書き出す文面。
 * ★見積ファイルの操作を変えたら、ここも直すこと。
 *   新規作成・複製のたびに書き直されるので、以後に作られるファイルには自動で行き渡る。
 */
/**
 * シートの行数・列数が足りなければ足す。★必ず clear() の前に呼ぶこと。
 * 行を詰めたシート（行数を減らしてあるシート）に setValues すると例外になり、
 * 先に clear() していると「消しただけ」で終わってしまうため。
 */
function ensureSheetSize_(sheet,rows,cols){
  var r=sheet.getMaxRows();
  if(r<rows)sheet.insertRowsAfter(r,rows-r);
  var c=sheet.getMaxColumns();
  if(c<cols)sheet.insertColumnsAfter(c,cols-c);
}

function quoteUsageSheetLines_(){
  return [
    'この見積ファイルの使い方',
    '',
    'このファイルは1件の見積専用です。「見積書」シートに宛先・工事名などを入力してください。',
    '',
    '① 入力した内容は、10分ごとに自動で母艦（見積書テンプレート）の「見積一覧」へ反映されます。',
    '　 すぐに反映したいときは、母艦を開いて「見積書メーカー」→「見積一覧を今すぐ同期」を実行してください。',
    '　 ★このファイルのメニュー「見積機能」→「見積一覧に登録・更新」は押す必要がありません。',
    '　 　古い見積ファイルでは「見積NO…が見積一覧に見つかりません」と出ることがありますが、',
    '　 　中身は自動で反映されているので、そのままで大丈夫です（2026-09-16）。',
    '',
    '② この見積を元に、似た内容の別の見積を作りたいとき（複製）',
    '　 ・「見積・請求ポータル」の見積一覧で、その見積の行の「複製」ボタンを押す',
    '　 ・または母艦の「見積一覧」で行を選び、「見積書メーカー」→「選択した見積を複製」',
    '　 どちらも、新しい見積NOが自動で振られ、年フォルダに保存され、「見積一覧」にも登録されます。',
    '　 ★このファイルのメニュー「見積機能」→「この見積を複製（新規保存）」は使わないでください。',
    '　 　古い見積ファイルではエラーになります（2026-09-10に、複製の処理を母艦側へ移しました）。',
    '',
    '③ 内訳書は10枚まで使えます。表紙の明細1〜10行目が、内訳書1〜10の小計とそれぞれ自動連動します。',
    '　 使わない内訳書は、シートごと削除して構いません（表紙の対応する行は空欄のままになります）。',
    '',
    '④「データ」シートに部署名・担当者名を追加すると、表紙の担当部署・担当者欄でリストから選べます。',
    '',
    '⑤ 法定福利費率・諸経費率の設定シートは、この見積だけに効きます（他の見積には影響しません）。',
    '',
    '⑥ 印刷・PDF化は「見積書」シートを開いて ファイル→印刷（またはダウンロード→PDF）から行います。',
    '',
    '※ 印刷設定（用紙A4・縦向き・余白狭い）は自動保存されないことがあります。印刷/PDF化の前に、',
    '　 各シートで ファイル→印刷 の用紙サイズA4・ページの向き縦向き・余白狭い を確認してください。'
  ];
}

/**
 * 見積ファイル（またはひな形）の「使い方」シートを最新の文面に書き直す。
 * 名前に「使い方」を含むシートだけを対象にする（他のシートは絶対に触らない）。
 * 途中で失敗しても見積の作成・複製そのものは止めない。
 */
function writeQuoteUsageSheet_(ss){
  try{
    var sheets=ss.getSheets();
    var sheet=null;
    for(var i=0;i<sheets.length;i++){
      if(sheets[i].getName().indexOf('使い方')!==-1){sheet=sheets[i];break;}
    }
    if(!sheet)sheet=ss.insertSheet('使い方',0);
    var values=quoteUsageSheetLines_().map(function(l){return [l];});
    ensureSheetSize_(sheet,values.length+1,2);
    try{sheet.getDataRange().breakApart();}catch(e){}
    sheet.clear();
    sheet.getRange(2,2,values.length,1).setValues(values);
    sheet.getRange(2,2,values.length,1).setWrap(false).setVerticalAlignment('middle');
    sheet.getRange(2,2).setFontSize(14).setFontWeight('bold');
    sheet.setColumnWidth(1,16);
    sheet.setColumnWidth(2,900);
    sheet.setHiddenGridlines(true);
    return true;
  }catch(e){
    return false;
  }
}

function usageSheetLines_(){
  var portalUrl='https://script.google.com/macros/s/AKfycbwD1wDvA9kSehumZjIJ-N3jy3EnfCz9D9PRkjLPVFTG017TabDdr3LWOIlHv5wchJXB/exec';
  return [
    '見積書メーカーの使い方（1案件1ファイル方式）',
    '',
    'このファイルは「見積一覧」の管理専用です。見積の中身（表紙・内訳書）は、案件ごとに別のファイルとして作られます。',
    '',
    '■ ふだんの入口は「見積・請求ポータル」です',
    '　見積の新規作成・検索・閲覧、見積の複製、請求書の作成は、ポータル（Webアプリ）からまとめてできます。',
    '　'+portalUrl,
    '　（Chromeのブックマークバーに「見積・請求ポータル」として登録済み。Googleアカウントでのログインが必要です）',
    '',
    '① 新しい見積を作る',
    '　ポータルの「＋ 新規作成」ボタン、またはこのファイルのメニュー「見積書メーカー」→「新しい見積を作成」で、',
    '　案件名（顧客名・工事名など）を入力すると、新しい見積ファイルがGoogleドライブ（見積保存 ＞ 年フォルダ）に',
    '　作られ、下の「見積一覧」に1行増えます。',
    '',
    '② 見積の中身を入力する',
    '　作られた見積ファイルを開いて、宛先・工事名・金額などを入力してください。',
    '　入力した内容は10分ごとに自動で「見積一覧」へ反映されます。',
    '　すぐに反映したいときは「見積書メーカー」→「見積一覧を今すぐ同期」を実行してください。',
    '　（見積ファイル側のメニュー「見積一覧に登録・更新」は押す必要がありません。古い見積ファイルでは',
    '　　「見つかりません」と出ることがありますが、中身は自動で反映されています）',
    '',
    '③ 過去の見積を探す',
    '　下の「見積一覧」シートで宛先や工事名を並べ替え・フィルタして探し、その行の「開く」リンクを押します。',
    '　ポータルの検索窓（宛先・工事名・見積NO）からでも探せます。',
    '',
    '④ 過去の見積に似た内容で新しく作る（複製）',
    '　「見積一覧」で複製したい行を選び、「見積書メーカー」→「選択した見積を複製」を実行します。',
    '　ポータルの見積一覧なら、各行の「複製」ボタンからでも複製できます。',
    '　どちらも、新しい見積NOが自動で振られ、年フォルダに保存され、「見積一覧」にも登録されます。',
    '　中身は複製元のままなので、開いて金額などを直してからお使いください。',
    '　★見積ファイル側のメニュー「この見積を複製（新規保存）」は使わないでください。',
    '　　古い見積ファイルではエラーになります（2026-09-10。複製の処理はこの母艦側に移しました）。',
    '',
    '⑤ 請求書を作る',
    '　メニュー「請求書メーカー」→「選択した見積から請求書を作成」（「見積一覧」の行を選んでから実行）。',
    '　見積が無い請求書は「新しい請求書を作成（見積なし）」。ポータルの「請求書」タブからも作れます。',
    '　★請求書ひな形には「法定福利費」「諸経費」の行はありません（2026-09-23〜。請求書には明示しない方針）。内訳・小計・端数調整・消費税・合計だけです。',
    '',
    '⑤-2 複数の工事をまとめて1枚の請求書にする（取引会社によっては月に数件まとめて請求するため。2026-09-23〜）',
    '　ポータルの見積書タブで、まとめたい見積（同じ宛先のものだけ）にチェックを入れ、',
    '　ツールバーの「🧾 まとめて請求書を作成」を押します。工事ごとに工事完了日を入力し（必須）、',
    '　件名（工事名称欄に入る文言）を確認・編集して「まとめて作成する」。',
    '　できた請求書は、内訳（12〜21行目）に工事ごとの1行（名称＝工事名称／品質・規格の欄に完了日／数量1／呼称「式」／金額＝その見積の合計金額）が入ります。',
    '　★内訳の行数の都合で最大10件まで。各工事の金額には法定福利費・諸経費・消費税がすでに含まれているため、',
    '　　端数調整・消費税は請求書側で二重に計算しないよう0円にしてあります（合計は内訳の金額の単純合計になります）。',
    '　★入力した工事完了日は、請求書一覧のI列（工事完了日）にも工事ごとに記録されます。',
    '',
    '⑥ いらない見積を一覧から消す（テストで作ったもの・書きかけのものなど）',
    '　ポータルの見積一覧で、その行の「削除」ボタン。まとめて消すときはツールバーの「🧹 未完成を整理」。',
    '　★一覧から消すと、ドライブの見積書ファイルもゴミ箱へ移動します（完全削除ではないので、ゴミ箱から元に戻せます）。',
    '　　ほかの人が作った見積は、所有者しかゴミ箱に入れられないため「見積保存 ＞ 削除済み」フォルダへ移動します。',
    '　★この「見積一覧」シートの行を手で消しても、同期や「見積一覧を作り直す」で復活します。必ずポータルから消してください。',
    '　　削除したものの控えは「削除ログ」シートに残ります。',
    '',
    '※ 一覧が乱れたときは「見積書メーカー」→「見積一覧を作り直す（フォルダ内を再スキャン）」で再構築できます',
    '　（対象フォルダ内の見積ファイルを全部スキャンし直すので、件数が多いと時間がかかります）。',
    '※ 同期に失敗したものは「同期エラー」シートに記録され、次の同期で自動的に再試行されます。',
    '　 今すぐ試すときは「同期エラーのみ再試行」。自動同期が止まっているときは「自動同期を有効にする（初回のみ）」。',
    '',
    '（このシートは「見積書メーカー」→「使い方シートを作り直す」で、いつでも最新の説明に書き直せます）'
  ];
}

/**
 * 「使い方」シートを探す。名前に「使い方」を含むシート → 見積一覧・同期エラー以外の先頭シート → 無ければ新規作成。
 */
function findUsageSheet_(ss){
  var sheets=ss.getSheets();
  for(var i=0;i<sheets.length;i++){
    if(sheets[i].getName().indexOf('使い方')!==-1)return sheets[i];
  }
  for(var j=0;j<sheets.length;j++){
    var n=sheets[j].getName();
    if(n!==INDEX_SHEET&&n!=='同期エラー'&&n!==DELETE_LOG_SHEET)return sheets[j];
  }
  return ss.insertSheet('使い方',0);
}

/**
 * 「使い方」シートを最新の文面で書き直す（メニュー用）。
 */
function rebuildUsageSheet(){
  var ui=SpreadsheetApp.getUi();
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=findUsageSheet_(ss);
  var res=ui.alert('使い方シートを作り直す',
    'シート「'+sheet.getName()+'」の中身を消して、最新の説明を書き直します。よろしいですか？\n\n'+
    '（元に戻したいときは、ファイル ＞ 版の履歴 ＞ 版の履歴を表示 から戻せます）',
    ui.ButtonSet.OK_CANCEL);
  if(res!==ui.Button.OK)return;
  writeUsageSheet_(sheet);
  var tplOk=writeQuoteUsageSheet_(SpreadsheetApp.openById(TEMPLATE_FILE_ID));
  ss.setActiveSheet(sheet);
  ui.alert('シート「'+sheet.getName()+'」を最新の説明に書き直しました。\n'+
    (tplOk?'「見積ひな形」の使い方シートも書き直しました（以後に作る見積に反映されます）。':'※「見積ひな形」の使い方シートは書き直せませんでした。'));
}

function writeUsageSheet_(sheet){
  var values=usageSheetLines_().map(function(l){return [l];});
  ensureSheetSize_(sheet,values.length+1,2);
  try{sheet.getDataRange().breakApart();}catch(e){}
  sheet.clear();
  sheet.getRange(2,2,values.length,1).setValues(values);
  sheet.getRange(2,2,values.length,1).setWrap(false).setVerticalAlignment('middle');
  sheet.getRange(2,2).setFontSize(14).setFontWeight('bold');
  sheet.setColumnWidth(1,16);
  sheet.setColumnWidth(2,900);
  sheet.setHiddenGridlines(true);
}

function createQuoteWeb(name){
  var validation=validateEntryName_(name);
  if(!validation.ok)return{success:false,message:validation.message};
  if(!TEMPLATE_FILE_ID)return{success:false,message:'見積ひな形ファイルのIDが未設定です。'};
  var quoteNo;
  try{quoteNo=nextQuoteNo_();}catch(e){return{success:false,message:e.message};}
  var fileName=quoteNo+'_'+validation.name;
  var newFile;
  try{
    var templateFile=DriveApp.getFileById(TEMPLATE_FILE_ID);
    if(TARGET_FOLDER_ID){
      var rootFolder=DriveApp.getFolderById(TARGET_FOLDER_ID);
      var folder=getYearFolder_(rootFolder);
      newFile=templateFile.makeCopy(fileName,folder);
    }else{
      newFile=templateFile.makeCopy(fileName);
    }
  }catch(e){
    return{success:false,message:'見積ファイルの作成に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var newCover=newSs.getSheetByName(COVER_SHEET);
  if(!newCover)return{success:false,message:'作成はできましたが「'+COVER_SHEET+'」シートが見つかりません。'};
  var createdDate=new Date();
  writeQuoteNoAndDate_(newCover,quoteNo,createdDate);
  writeQuoteUsageSheet_(newSs);
  var ss=SpreadsheetApp.openById(MOTHER_SS_ID);
  var idx=ensureIndexSheet_(ss);
  var row=idx.getLastRow()+1;
  idx.getRange(row,1).setValue(quoteNo);
  idx.getRange(row,2).setValue(createdDate);
  idx.getRange(row,7).setFormula('=HYPERLINK("'+newFile.getUrl()+'","開く")');
  idx.getRange(row,8).setValue(newCover.getRange('C7').getValue());
  return{success:true,url:newFile.getUrl(),message:'新しい見積ファイル「'+fileName+'」を作成しました。'};
}
function createInvoiceWeb(name){if(!name||!name.trim())return{success:false,message:'名前が入力されていません。'};if(!INVOICE_TEMPLATE_FILE_ID)return{success:false,message:'請求書ひな形が未設定です。'};var newFile;try{var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);var folder=getYearFolder_(rootFolder);newFile=templateFile.makeCopy('請求書_'+name.trim(),folder);}catch(e){return{success:false,message:'請求書ファイルの作成に失敗しました：'+e.message};}var newSs=SpreadsheetApp.openById(newFile.getId());var cover=newSs.getSheetByName(INVOICE_SHEET);if(!cover)return{success:false,message:'作成はできましたが「'+INVOICE_SHEET+'」シートが見つかりません。'};var invoiceNo=nextInvoiceNo_();cover.getRange('L2').setValue(invoiceNo);cover.getRange('L3').setValue(new Date());appendInvoiceListRow_('',new Date(),name.trim(),'',newFile.getUrl());return{success:true,url:newFile.getUrl(),message:'新しい請求書ファイル「'+name.trim()+'」を作成しました。'};}

function createInvoiceFromQuoteWeb(q){
  if(!q||!q.client||!q.jobName){
    return{success:false,message:'見積データが正しく渡されませんでした。'};
  }
  if(!INVOICE_TEMPLATE_FILE_ID){
    return{success:false,message:'請求書ひな形が未設定です。'};
  }
  var newFile;
  try{
    var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);
    var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);
    var folder=getYearFolder_(rootFolder);
    newFile=templateFile.makeCopy('請求書_'+String(q.client)+'_'+String(q.jobName),folder);
  }catch(e){
    return{success:false,message:'請求書ファイルの作成に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var cover=newSs.getSheetByName(INVOICE_SHEET);
  if(!cover){
    return{success:false,message:'作成はできましたが「'+INVOICE_SHEET+'」シートが見つかりません。'};
  }
  var invoiceNo=nextInvoiceNo_();
  cover.getRange('L2').setValue(invoiceNo);
  cover.getRange('L3').setValue(new Date());
  cover.getRange('B5').setValue(q.client);
  cover.getRange('C9').setValue(q.jobName);
  cover.getRange('C7').setValue(q.total);
  cover.getRange('J10').setValue(q.dept);
  cover.getRange('L10').setValue(q.person);
  appendInvoiceListRow_(q.client,new Date(),q.jobName,q.total,newFile.getUrl(),q.no);
  return{success:true,url:newFile.getUrl(),message:'見積NO「'+q.no+'」('+q.client+')から請求書を作成しました。'};
}

// 同じ宛先の複数の見積をまとめて1枚の請求書にする（取引会社によっては工事案件を月にまとめて請求するため。2026-09-23）。
// items: [{no,client,jobName,dept,person,total}, ...]（Portal.htmlの見積一覧の行そのもの）。label: 請求書の「工事名称」欄に入れる件名。
// 請求書ひな形の内訳（12〜21行目）に1件＝1行で書く：B列=工事名称／H列=数量1／I列=呼称「式」／K列=金額（見積の合計金額そのまま）。
// ★2026-09-23: 請求書ひな形から「法定福利費」「諸経費」の行を削除した(菅原氏の判断。請求書には明示不要のため)。
//   これに伴い行がずれ、22行目=小計(=SUM(K12:K21))・23行目=端数調整・24行目=消費税・25行目=合計(税込、C7はこの式のまま)。
// ★各工事の金額はすでに法定福利費・諸経費・消費税込みの最終金額（見積のC7＝合計と同じ計算式）なので、
//   請求書側のK23(端数調整)・K24(消費税)は0にして二重計上を防ぐ。これでC7・K22(小計)・K25(合計)は内訳の合計と自動で一致する。
function createBundledInvoiceWeb(items,label){
  if(!items||!items.length){
    return{success:false,message:'まとめる見積が選ばれていません。'};
  }
  if(items.length>10){
    return{success:false,message:'まとめられる工事は10件までです（内訳の行数の都合）。'};
  }
  var client=items[0].client;
  for(var i=1;i<items.length;i++){
    if(items[i].client!==client){
      return{success:false,message:'宛先（取引会社）が異なる見積が含まれています。同じ宛先の見積だけを選んでください。'};
    }
  }
  var completionDates=[];
  for(var d=0;d<items.length;d++){
    var parsed=parseIsoDate_(items[d].completionDate);
    if(!parsed){
      return{success:false,message:'「'+(items[d].jobName||items[d].no)+'」の工事完了日が入力されていません。すべての工事に工事完了日を入力してください。'};
    }
    completionDates.push(parsed);
  }
  if(!INVOICE_TEMPLATE_FILE_ID){
    return{success:false,message:'請求書ひな形が未設定です。'};
  }
  var fileLabel=(label&&String(label).trim())?String(label).trim():(String(client)+'_'+items.length+'件まとめ');
  var newFile;
  try{
    var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);
    var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);
    var folder=getYearFolder_(rootFolder);
    newFile=templateFile.makeCopy('請求書_'+fileLabel,folder);
  }catch(e){
    return{success:false,message:'請求書ファイルの作成に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var cover=newSs.getSheetByName(INVOICE_SHEET);
  if(!cover){
    return{success:false,message:'作成はできましたが「'+INVOICE_SHEET+'」シートが見つかりません。'};
  }
  var invoiceNo=nextInvoiceNo_();
  var total=0;
  for(var j=0;j<items.length;j++)total+=Number(items[j].total)||0;
  cover.getRange('L2').setValue(invoiceNo);
  cover.getRange('L3').setValue(new Date());
  cover.getRange('B5').setValue(client);
  cover.getRange('C9').setValue(fileLabel);
  cover.getRange('J10').setValue(items[0].dept||'');
  cover.getRange('L10').setValue(items[0].person||'');
  var tz=Session.getScriptTimeZone();
  for(var k=0;k<items.length;k++){
    var row=12+k;
    cover.getRange(row,2).setValue(items[k].jobName||'');
    cover.getRange(row,4).setValue('完了日：'+Utilities.formatDate(completionDates[k],tz,'yyyy/MM/dd'));
    cover.getRange(row,8).setValue(1);
    cover.getRange(row,9).setValue('式');
    cover.getRange(row,11).setValue(Number(items[k].total)||0);
  }
  cover.getRange('K23').setValue(0);
  cover.getRange('K24').setValue(0);
  var createdDate=new Date();
  for(var m=0;m<items.length;m++){
    appendInvoiceListRow_(client,createdDate,items[m].jobName,items[m].total,newFile.getUrl(),items[m].no,Utilities.formatDate(completionDates[m],tz,'yyyy/MM/dd'));
  }
  return{success:true,url:newFile.getUrl(),message:items.length+'件の工事をまとめて請求書「'+fileLabel+'」を作成しました（合計'+total.toLocaleString('ja-JP')+'円）。'};
}


function createBlankInvoiceWeb(){
  if(!INVOICE_TEMPLATE_FILE_ID){
    return{success:false,message:'請求書ひな形が未設定です。'};
  }
  var newFile;
  var name='下書き_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  try{
    var templateFile=DriveApp.getFileById(INVOICE_TEMPLATE_FILE_ID);
    var rootFolder=DriveApp.getFolderById(INVOICE_FOLDER_ID);
    var folder=getYearFolder_(rootFolder);
    newFile=templateFile.makeCopy('請求書_'+name,folder);
  }catch(e){
    return{success:false,message:'請求書ファイルの作成に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var cover=newSs.getSheetByName(INVOICE_SHEET);
  if(!cover){
    return{success:false,message:'作成はできましたが「'+INVOICE_SHEET+'」シートが見つかりません。'};
  }
  var invoiceNo=nextInvoiceNo_();
  cover.getRange('L2').setValue(invoiceNo);
  cover.getRange('L3').setValue(new Date());
  appendInvoiceListRow_('',new Date(),name,'',newFile.getUrl());
  return{success:true,url:newFile.getUrl(),message:'新しい請求書ファイルを作成しました。'};
}

function createBlankQuoteWeb(){
  if(!TEMPLATE_FILE_ID){
    return{success:false,message:'見積ひな形ファイルのIDが未設定です。'};
  }
  var newFile;
  var name='下書き_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  try{
    var templateFile=DriveApp.getFileById(TEMPLATE_FILE_ID);
    if(TARGET_FOLDER_ID){
      var rootFolder=DriveApp.getFolderById(TARGET_FOLDER_ID);
      var folder=getYearFolder_(rootFolder);
      newFile=templateFile.makeCopy(name,folder);
    }else{
      newFile=templateFile.makeCopy(name);
    }
  }catch(e){
    return{success:false,message:'見積ファイルの作成に失敗しました：'+e.message};
  }
  var newSs=SpreadsheetApp.openById(newFile.getId());
  var newCover=newSs.getSheetByName(COVER_SHEET);
  if(!newCover)return{success:false,message:'作成はできましたが「'+COVER_SHEET+'」シートが見つかりません。'};
  var quoteNo;
  try{quoteNo=nextQuoteNo_();}catch(e){return{success:false,url:newFile.getUrl(),message:'ファイルは作成されましたが、見積NOを振れませんでした。'+e.message};}
  var createdDate=new Date();
  writeQuoteNoAndDate_(newCover,quoteNo,createdDate);
  writeQuoteUsageSheet_(newSs);
  var ss=SpreadsheetApp.openById(MOTHER_SS_ID);
  var idx=ensureIndexSheet_(ss);
  var row=idx.getLastRow()+1;
  idx.getRange(row,1).setValue(quoteNo);
  idx.getRange(row,2).setValue(createdDate);
  idx.getRange(row,7).setFormula('=HYPERLINK("'+newFile.getUrl()+'","開く")');
  idx.getRange(row,8).setValue(newCover.getRange('C7').getValue());
  return{success:true,url:newFile.getUrl(),message:'新しい見積ファイルを作成しました。'};
}

// ══════════════════════════════════════════
//  マニュアル(PDF)のポータル内表示
// ══════════════════════════════════════════
const MANUAL_FOLDER_NAME = 'マニュアル';
const MANUAL_FILE_NAME = '見積・請求ポータル利用マニュアル.pdf';
// 「見積・請求管理」フォルダ(TARGET_FOLDER_IDの一つ上の階層。診断エンドポイントで確認済み)
const MANUAL_PARENT_FOLDER_ID = '1R_8hfNavAOVMXO1duJvyLswtNoxny1ci';

// マニュアルPDFの保存先フォルダを取得(なければ作成)。
// 「見積・請求管理」フォルダ(MANUAL_PARENT_FOLDER_ID)の中の「マニュアル」サブフォルダに置く
// (ユーザーが普段から使っているフォルダにまとめるため。TARGET_FOLDER_IDは見積保存という
// 一段階下のフォルダを指しており、当初はそちらを見に行ってしまっていたため修正)。
function getManualFolder_(){
  var parent = DriveApp.getFolderById(MANUAL_PARENT_FOLDER_ID);
  var it = parent.getFoldersByName(MANUAL_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return parent.createFolder(MANUAL_FOLDER_NAME);
}

// 現在保存されているマニュアルPDFのURLを取得する。Portal.htmlの「マニュアル」ボタンから呼ばれる。
// ファイル自体を都度アップロードし直す設計のため、URLは固定ではなくその都度検索して返す。
function getManualUrl(){
  var folder = getManualFolder_();
  var it = folder.getFilesByName(MANUAL_FILE_NAME);
  if (!it.hasNext()) return {success:false, message:'マニュアルがまだアップロードされていません。'};
  var file = it.next();
  return {success:true, url:file.getUrl()};
}

// マニュアルPDFをアップロード(既存があれば置き換え)する。ポータルからは呼ばれず、
// マニュアルを更新したときに管理者側から1回だけ実行する用(GAS_URLへPOST、action=uploadManual)。
function uploadManualPdf_(base64Data){
  var folder = getManualFolder_();
  var existing = folder.getFilesByName(MANUAL_FILE_NAME);
  while (existing.hasNext()) { existing.next().setTrashed(true); } // 既存版はゴミ箱へ(URLは変わるが、ポータル側は毎回検索するため影響なし)
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', MANUAL_FILE_NAME);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {success:true, url:file.getUrl()};
}
