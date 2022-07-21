// tool_ss_property："採用ツール"SSのスクリプトプロパティ
var property = PropertiesService.getScriptProperties();

function setNewProperty(){
  // "採用ツール"SS（スプレッドシート）をSpreadsheetオブジェクトとして取得
  const tool_ss = SpreadsheetApp.getActiveSpreadsheet();
  // tool_ssのID
  const tool_ss_ID =tool_ss.getId();

  // tool_ss_IDをスクリプトプロパティに追加
  property.setProperty("Tool_SS_ID", tool_ss_ID);

  // tool_ssから親フォルダ（今回の採用活動フォルダ）を特定
  const parentsFolder_ID = DriveApp.getFileById(tool_ss_ID).getParents().next().getId();

  // parentsFolder_IDをスクリプトプロパティに追加
  property.setProperty("ParentsFolder_ID", parentsFolder_ID);

  // データベースのID
  const db_ID = "1wPDkdte2AH78NtwRqu9LNMMSwDBy4lZ9wM5oP69h0kY";
  property.setProperty("DB_ID", db_ID);

  // parameters
  // 今日の日付から、採用年度と学期を取得
  const now_date = new Date();
  var year = now_date.getFullYear();
  var month = now_date.getMonth();
  if (8 < month){
    year += 1;
    semester = "前期";
    season = "冬";
  } else {
    semester = "後期";
    season = "夏";
  }

  // スクリプトステータスに追加
  property.setProperty("Year", year);
  property.setProperty("Semester", semester);

  // SS: "採用ツール" → "20XX年度前・後期AIM新規学生スタッフ 採用スケジュール"
  const ss_title = year + "年度" + semester + "AIM新規学生スタッフ 採用スケジュール";
  // tool_ss.rename(ss_title);

  // 日程の取得
  // "採用スケジュール"SSの"工程表"シート
  const schedule_sheet = tool_ss.getSheetByName("工程表");
  // [工程表]シートのデータを取得
  const schedule_data = schedule_sheet.getDataRange().getValues();
  // calender：内容ごとの辞書オブジェクト化したschedule
  var calender = {};
  // sliceで3列目（日付の始まるとこ）から末までを指定
  calender["日付"] = schedule_data[5].slice(3, schedule_data[5].length);
  // "日付"以外は、for文で読み込む
  for(let i = 7; i < schedule_data.length; i++){
    var prosess = schedule_data[i][1]
    calender[prosess] = schedule_data[i].slice(3, schedule_data[i].length);
  };

  // 面接期間の取得
  // 期間の"開始"日と"締切"日のインデックスを取得
  for (let i = 0; i < calender["日付"].length; i++){
    if (calender["面接期間"][i] == '開始'){
      startdate_index = i;
    } else if(calender["面接期間"][i] == "締め切り"){
      enddate_index = i;
    }
  }
  // 期間日付の配列に
  var interview_period = calender["日付"].slice(startdate_index, enddate_index + 1);
  // 締切から日を遡る
  for(let i = interview_period.length -1; i >= 0; i--){
    // .getDayで曜日を取得
    // 期間内の日付が、日曜（=0）か土曜（=6）ならば削除
    if (interview_period[i].getDay() == 0 || interview_period[i].getDay() == 6){
      interview_period.splice(i, 1);
      continue;
    } else {
      interview_period[i] = getJapaneseDate(interview_period[i]);
    }
  }

  // 面接期間をスクリプトプロパティに追加
  property.setProperty("Interview_Period", JSON.stringify(interview_period))

  // 募集期間の取得
  for (let i = 0; i < calender["日付"].length; i++){
    if (calender["募集期間(書類足切り)"][i] == '開始'){
      opendate_index = i;
    } else if(calender["募集期間(書類足切り)"][i] == "締め切り"){
      closedate_index = i;
    };
  };
  const open_date = new Date(calender["日付"][opendate_index].setHours(0, 0, 0));
  const close_date = new Date(calender["日付"][closedate_index].setHours(12, 0, 0));

  // 募集期間の開始終了日をスクリプトプロパティに追加
  property.setProperty("募集開始日", open_date);
  property.setProperty("募集終了日", close_date);

  // トリガーを全削除
  var allTriggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}