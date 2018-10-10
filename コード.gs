function pcode() {

  //UIを取得する
  var ui = SpreadsheetApp.getUi();
  var msg = "";
  
  //ドキュメントロックを使用する
  var lock = LockService.getDocumentLock();
  //30秒間のロックを取得
try {
  //ロックを実施する
  lock.waitLock(30000);

  //スプレッドシート情報取得の定義
  var SPREADSHEET_ID = '17TcG7lRyRCoyc7XCDz8zWJA1GUvfgaTYjzSlSz-1n_4';
  var SHEET_NAME = '郵便番号×距離検索';  
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var slast_row = sheet.getLastRow();
  //var klast_row = sheet.getRange("F:F").getLastRow();
  //var tlast_row = sheet.getRange("B:B").getLastRow();

  sheet.getRange(3,6,slast_row - 2,14).clearContent();
  
  var now = new Date();
    
  sheet.getRange(1,  3).setValue('最終実行日時：'+ Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  
 var last_row = sheet.getLastRow();
    
 var ash = SpreadsheetApp.getActiveSpreadsheet();
 var cnt = ash.getNumSheets();

　ash.moveActiveSheet(1);

　for(var z = cnt; z >= 2; z--){
　　var sh = ash.getSheets()[z-1];
　　ash.deleteSheet(sh);
　}
  
  var CHIIKI={};
  var PCODE={};
  var KYORI={};
  var sname={};
  var kekka={};
  
  for (i = 0; 3 + i <= last_row; i++) {
  
  CHIIKI[i] = sheet.getRange(3 + i, 2).getValue(); 
  PCODE[i] = sheet.getRange(3 + i, 3).getValue();   
  KYORI[i] = sheet.getRange(3 + i, 4).getValue();  
    
  kekka[i] = ash.insertSheet(CHIIKI[i]);
  
  var json_url={};  
    
  json_url[i] = 'https://everyday-growth.com/zipdistance/api/?pcode='+ PCODE[i] +'&dis='+ KYORI[i];
  
  var json={};   
  var jsonData={};  
    
  json[i] = UrlFetchApp.fetch(json_url[i]).getContentText();
　jsonData[i] = JSON.parse(json[i]);
 
 var SPOSTCODE={}; 
 var SADD1={}; 
 var SADD2={}; 
 var SADD3={};  
 var SYOMI1={}; 
 var SYOMI2={}; 
 var SYOMI3={}; 
 var LONGITUDE={}; 
 var LATITUDE={};  
 var DISTANCE={};  
 var ResponseCode={}; 
 var ERResponseCode={};    
 var ResponseStatus={}; 
 var ResponseDescription={};     
 
 ResponseCode[i] = jsonData[i].Response.ResponseCode;
 ResponseStatus[i] = jsonData[i].Response.ResponseStatus;
 ResponseDescription[i] = jsonData[i].Response.ResponseDescription;    

 if (ResponseStatus[i] == 'OK'){      
    
 SPOSTCODE[i] = jsonData[i].Response.Search.POSTCODE;
 SADD1[i] = jsonData[i].Response.Search.ADD001; 
 SADD2[i] = jsonData[i].Response.Search.ADD002; 
 SADD3[i] = jsonData[i].Response.Search.ADD003;  
 SYOMI1[i] = jsonData[i].Response.Search.ADDYOMI001; 
 SYOMI2[i] = jsonData[i].Response.Search.ADDYOMI002; 
 SYOMI3[i] = jsonData[i].Response.Search.ADDYOMI003; 
 LONGITUDE[i] = jsonData[i].Response.Search.LONGITUDE; 
 LATITUDE[i] = jsonData[i].Response.Search.LATITUDE;  
 DISTANCE[i] = jsonData[i].Response.Search.DISTANCE;  
     
 var DATA = jsonData[i].Response.Results; 
  
 sheet.getRange(3 + i, 6).setValue(SPOSTCODE[i]);  
 sheet.getRange(3 + i, 7).setValue(SADD1[i]);
 sheet.getRange(3 + i, 8).setValue(SADD2[i]);
 sheet.getRange(3 + i, 9).setValue(SADD3[i]);
 sheet.getRange(3 + i, 10).setValue(SYOMI1[i]);
 sheet.getRange(3 + i, 11).setValue(SYOMI2[i]);
 sheet.getRange(3 + i, 12).setValue(SYOMI3[i]); 
 sheet.getRange(3 + i, 13).setValue(LONGITUDE[i]);  
 sheet.getRange(3 + i, 14).setValue(LATITUDE[i]);
 sheet.getRange(3 + i, 15).setValue(DISTANCE[i]);  
 sheet.getRange(3 + i, 16).setValue(ResponseCode[i]);   
 sheet.getRange(3 + i, 17).setValue(ResponseStatus[i]);
 sheet.getRange(3 + i, 18).setValue(ResponseDescription[i]);
 //sheet.getRange(3, 19).setValue(ResponseCount);  
 
  kekka[i].getRange(1, 1).setValue('地域名');
  kekka[i].getRange(1, 2).setValue('郵便番号');
  kekka[i].getRange(1, 3).setValue('住所');
  kekka[i].getRange(1, 4).setValue('距離（km）');
  kekka[i].getRange(1, 5).setValue('hue貼付用');
  
  //var ary={};  
  
  var ary = [[]];  
  for (j = 0; j < DATA.length; j++) {
    
  var a = DATA[j].POSTCODE;
  var b = DATA[j].ADD001+' '+DATA[j].ADD002+' '+DATA[j].ADD003+' '; 
  var c = DATA[j].DISTANCE;
  var d = "''"+DATA[j].POSTCODE+"'"+",";
  
   
  ary[j] = [CHIIKI[i],a,b,c,d];       
     
}

  
  var rows = ary.length;
  var cols = ary[0].length;
  kekka[i].getRange(2,1,rows,cols).setValues(ary);
   
 }else{   
   
 sheet.getRange(3 + i, 16).setValue(ResponseCode[i]);  
 sheet.getRange(3 + i, 17).setValue(ResponseStatus[i]);
 sheet.getRange(3 + i, 18).setValue(ResponseDescription[i]);  
 
 }  
 }    
    
   //メッセージを格納
   msg = "処理が完了しました";
 
} catch (e) {
   //ロック取得できなかった時の処理等を記述する
   var checkword = "ロックのタイムアウト: 別のプロセスがロックを保持している時間が長すぎました。";
  
   //通常のエラーとロックエラーを区別する
   if(e.message == checkword){
      //ロックエラーの場合
      msg = "誰かが使用中です";
   }else{
      //それ以外のエラーの場合
      msg = e.message;
   }  　 
 
}　finally　{

  var objSheet = ash.getSheetByName('郵便番号×距離検索'); 
　ash.setActiveSheet(objSheet); 
  
   //ロックを開放する
   lock.releaseLock();
  
   //メッセージを表示する
   ui.alert(msg);
}
  Logger.log(jsonData); 
  Logger.log(DATA); 
  Logger.log(ary); 
  Logger.log(last_row); 

}
