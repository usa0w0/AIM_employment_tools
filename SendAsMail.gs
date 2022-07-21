function SendAsMail() {
  // 確認メッセージ
  Browser.msgBox("一斉送信を開始します", Browser.Buttons.OK);

  //CONFIG 習得
  //ここでメール文の型を拾ってくるよ
  var config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG_for_MAIL"); //メールの型が書いてあるシートの名前にしてね
  var config_data = config_sheet.getDataRange().getValues();
  var config_lastrow = config_sheet.getLastRow();
  var config = {};
  for(var i=1; i<config_lastrow; i++){
    config[config_data[i][0]] = config_data[i][1];
  }
  
  //スプレッドシート 習得
  //ここで学生の個人情報を拾ってくるよ
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("スクリプト処理用コピー"); //学生の情報がまとまっているシートの名前にしてね
  var data = sheet.getDataRange().getValues();
  var lastrow = sheet.getLastRow();
  
  //スプレッドシート1行目 習得
  //デバッグ用
  var headerKeys = data[0];
  //var headerKeys = data.getRange(1,1,1, sheet.getLastColumn()).getValues()[0];
  //Logger.log("headerKeys:"+headerKeys);　//log で調節できるように入れているよ
  /* logはスクリプトエディタの画面の　表示→ログ　で見れるよ */
  
  //送信メールの文章を確定
  JPweekTbl = new Array("0","月","火","水","木","金","土","日");
  
  var mail_body = config["MAIL_HEADER"];
  for(var row=1; row<lastrow; row++){
    //CONFIG からデータの取り込み
    var line = {};
    data[row].forEach(function(value, i){
      line[headerKeys[i]] = value;
    });
    
    var mail_to = line["メールアドレス (AOYAMA-mail)"] //シートから拾ってきた送信先のメアド
    var subject = ""; //件名
    var message = ""; //本文
    var flag = 0; //分岐フラグ、本文に対する処理を行う挙動を分岐させる
//    Logger.log("最初:"+line["スクリプトステータス"]);
    //応募者の条件に応じた本文作成
    if(line["スクリプトステータス"] == "書類不合格_未送信"){ 
        subject = config["書類不合格_タイトル"];
        message = config["書類不合格_本文"];
        sheet.getRange(row+1, data[0].indexOf('スクリプトステータス')+1).setValue('不合格_送信済み');
        
      } else if (line["スクリプトステータス"] == "アサイン済み"){
        subject = config["書類合格_タイトル"];
        message = config["書類合格_本文"];
        sheet.getRange(row+1, data[0].indexOf('スクリプトステータス')+1).setValue('面接案内_送信済み');

        // var week_num = Utilities.formatDate(new Date(line["面接日時"]), "Asia/Tokyo", "u");
        // line["面接日時"] = Utilities.formatDate(new Date(line["面接日時"]), "Asia/Tokyo", "yyyy/MM/dd(%曜日%) HH:mm");
        // line["面接日時"] = line["面接日時"].replace("%曜日%", JPweekTbl[week_num]);
        
        flag = 1;
                
      } else if (line["スクリプトステータス"] == "面接合格_未送信"){
        subject = config["面接合格_タイトル"];
        message = config["面接合格_本文"];
        sheet.getRange(row+1, data[0].indexOf('スクリプトステータス')+1).setValue('面接合格_送信済み');
        
        flag = 1;

      } else if (line["スクリプトステータス"] == "面接不合格_未送信"){
        subject = config["面接不合格_タイトル"];
        message = config["面接不合格_本文"];
        sheet.getRange(row+1, data[0].indexOf('スクリプトステータス')+1).setValue('不合格_送信済み');
        
        flag = 1;

      } else {
        continue;
      }
    headerKeys.forEach(function(value, i){
      message = message.replace("%" + value + "%", line[value]);
    });
    if(flag == 1) message = message.replace("%氏名%",line["氏名"]); //本文中に氏名が２回出てくるのでもう一度変換  
        
    //テスト用、mail_toに管理スタッフの誰かのメアドを記入してください
    // mail_to = "c5622040@ima.aim.aoyama.ac.jp";
    // GmailApp.sendEmail(mail_to, subject, message,{from:config["MAIL_CC"],name:"情報メディアセンター　学生スタッフ担当"});
    GmailApp.sendEmail(mail_to, subject, message,{cc:config["MAIL_CC"],from:config["MAIL_CC"],name:"情報メディアセンター　学生スタッフ担当"}); //ccにsagamipro-contact を入れている、本番はこっちでやって
  }
}