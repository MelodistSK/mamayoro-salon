/**
 * ままよろサロン 申込フォーム → スプレッドシート集約（Apps Script ウェブアプリ）
 * スプレッドシートに「拡張機能 → Apps Script」で貼り付けて使う（このシートに紐づくスクリプト）。
 * 列 = 将来プラットフォーム（開催回×参加者×回答）の最小スキーマ。
 */
var SHEET_NAME = '回答';
var HEADERS = ['回答ID','受付日時','開催ID','開催年月','シリーズ','種別',
               'お名前','会社名','メール','電話','業種','会社規模','興味の方向','困りごと','source','同意'];

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(30000);                                  // 同時送信でも行が壊れないように
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if(sh.getLastRow() === 0) sh.appendRow(HEADERS);       // 初回だけ見出し行
    var p = (e && e.parameter) ? e.parameter : {};
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([
      Utilities.getUuid(), now,
      p.event_id || '', p.event_ym || '', p.series || '', p.kind || '無料相談',
      p.name || '', p.company || '', p.email || '', p.phone || '',
      p.industry || '', p.size || '', p.direction || '', p.concern || '',
      p.source || '', p.consent || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
                         .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
                         .setMimeType(ContentService.MimeType.JSON);
  }finally{
    lock.releaseLock();
  }
}

// 動作確認用（ブラウザでウェブアプリURLを開くと表示される）
function doGet(){
  return ContentService.createTextOutput('ままよろサロン フォーム受付エンドポイント：正常稼働中');
}
