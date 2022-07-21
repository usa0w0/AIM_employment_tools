// 面接日時取得
function getInterviewDate(){
  // フォームの集計用SSを取得
  const form_ss = SpreadsheetApp.getActiveSpreadsheet();
  // アサインカレンダーとスクリプト処理用コピーのシートを取得
  const schedule_asign = form_ss.getSheetByName("日程調整シート");
  const form_copy_sheet = form_ss.getSheetByName("スクリプト処理用コピー");

  // [スクリプト処理用コピー]のデータと件数を取得
  var data = form_copy_sheet.getDataRange().getValues();
  var lastrow = form_copy_sheet.getLastRow();

  // [スクリプト処理用コピー]のヘッダーラベルを取得すると共に、書き込み用に"面接日時"のインデックスを取得
  var headerKeys = data[0];
  var interview_date_index = headerKeys.indexOf("面接日時");

  // アサインカレンダー（記入用）を取得
  var schedule = schedule_asign.getRange(24, 2, 21, schedule_asign.getLastColumn()-1).getValues();
  // 面接日のヘッダー
  var dateheader = schedule[0];
  // 面接時間のヘッダー
  var timeheader = schedule.map(item => item[0]);

  // 応募者それぞれについて面接日時を検索
  for(var row=1; row<lastrow; row++){
    var line = {};
    data[row].forEach(function(value, i){
      line[headerKeys[i]] = value;
    });
    
    var name = line["氏名"] //シートから拾ってきた名前

    // 行数（時間）インデックス取得
    for (let i = 0; i < schedule.length; i++){
      if (schedule[i].indexOf(name) >= 0){
        var time_index = i
        break;
      }
    }
    if (!time_index){
      break;
    }
    // 列数（日付）インデックス
    var date_index = schedule[time_index].indexOf(name);

    // m月d日(曜日) H:MM〜H:MM 形式に結合
    var date = dateheader[date_index] + " " + timeheader[time_index].replace("\n|\n", "〜");

    // 面接日時と担当者を[スクリプト処理用コピー]に転記
    form_copy_sheet.getRange(row+1, interview_date_index+1).setValue(date);
    form_copy_sheet.getRange(row+1, interview_date_index+2).setValue(schedule[time_index+1][date_index]);
  }
}