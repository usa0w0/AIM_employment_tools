// フォームの集計用のSSを生成する関数
function createSS() {
  // "採用スケジュール"SS
  const tool_ss = SpreadsheetApp.openById(property.getProperty("Tool_SS_ID"))
  // それぞれのシートを宣言
  var interview_sheet = tool_ss.getSheetByName("SS用・面接項目")
  var config_for_sheet_data = tool_ss.getSheetByName("SS用・CONFIG_for_SHEET")
  var config_for_mail_data = tool_ss.getSheetByName("SS用・CONFIG_for_MAIL")
  var question_sheet = tool_ss.getSheetByName("フォーム用・質問");

  // parameters
  const parentsFolder_ID = property.getProperty("ParentsFolder_ID");
  const year = property.getProperty("Year");
  const semester = property.getProperty("Semester");
  const interview_period = JSON.parse(property.getProperty("Interview_Period"));
  const db_ID = property.getProperty("DB_ID");

  // form_ss：応募フォームの集計や面接アサインをするためのSS
  // form_ss = Spreadsheetオブジェクト
  const form_ss_title = year + "年度" + semester + "AIM新規学生スタッフ 応募フォーム 集計";
  const form_ss = SpreadsheetApp.create(form_ss_title);
  // form_ss_ID：form_ssのID
  const form_ss_ID = form_ss.getId();
  // スクリプトプロパティに追加
  property.setProperty("Form_SS_ID", form_ss_ID)

  // マイドライブに生成されたSSを採用フォルダに移動
  const form_ss_File = DriveApp.getFileById(form_ss_ID);
  DriveApp.getFolderById(parentsFolder_ID).addFile(form_ss_File);
  DriveApp.getRootFolder().removeFile(form_ss_File);

  /* ------------------------------------------- */
  // シート生成
  // 生成されたSSの"シート1"を"スクリプト処理用コピー"へ
  const form_copy_sheet = form_ss.getSheetByName("シート1").setName("スクリプト処理用コピー");
  // 日程調整シート：面接日アサイン用シート を追加
  const schedule_asign = form_ss.insertSheet("日程調整シート", 1);
  // 面接シート：面接記録のシート を追加
  const interview_report = form_ss.insertSheet("面接シート", 2);
  // CONFIG_for_SHEET を追加
  const config_for_sheet = form_ss.insertSheet("CONFIG_for_SHEET", 3);
  // CONFIG_for_MAIL を追加
  const config_for_mail = form_ss.insertSheet("CONFIG_for_MAIL", 4);
  // 過去応募者（past_applicants）を追加
  const past_applicants = form_ss.insertSheet("過去応募者（2016年度冬〜）", 5);

  /* ------------------------------------------- */
  // [フォーム用・質問]シートから、募集フォームの質問項目を取得
  // Array.map(item => item[n])…二次元配列のn番目列を取り出す定型文
  var question_label = question_sheet.getDataRange().getValues().map(item => item[0]);

  // 先頭を[質問タイトル]→[タイムスタンプ]に
  question_label[0] = "タイムスタンプ";

  // [面接可能時間]の部分は、実際の日程リストに置き換える
    // [面接可能時間]のインデックスを、質問形式[面接日程]から取得
    const target = question_sheet.getDataRange().getValues().map(item => item[2]).indexOf("面接日程");
    // [面接可能時間]を日程に置換
    question_label = question_label.slice(0, target).concat(interview_period).concat(question_label.slice(target+1));
    // 末尾に[スクリプトステータス]の項目を追加
    question_label.push('スクリプトステータス', '面接日時', '担当者', '過去応募');
  // [スクリプト処理用コピー]のラベルに
  form_copy_sheet.getRange(1,1,1,question_label.length).setValues([question_label]);

  // 過去応募回数数式をセット
  form_copy_sheet.getRange(2, question_label.length).setFormulaR1C1("=COUNTIF('過去応募者（2016年度冬〜）'!R2C4:R1000C4,RC4)")

  // 余分範囲を削除
  remakeSheetRange(form_copy_sheet, 2, false);

  // ラベルを表示固定（1行5列）
  form_copy_sheet.setFrozenRows(1);
  form_copy_sheet.setFrozenColumns(3);

  /* ------------------------------------------- */
  // 面接者
  schedule_asign.getRange('A2').setValue("面接者");
  // 応募者を入力するプルダウンリスト
  var asign_range = form_copy_sheet.getRange("B2:B");
  // 面接担当
  var interview_range = config_for_sheet.getRange("C2:C");
  // 入力規則を作成
  var asign_rule = SpreadsheetApp.newDataValidation().requireValueInRange(asign_range).build();
  var interview_rule = SpreadsheetApp.newDataValidation().requireValueInRange(interview_range).build();
  // リストをセットするセル範囲を取得（）
  var cell = schedule_asign.getRange('B2');
  //セルに入力規則をセット
  cell.setDataValidation(asign_rule);
  // アサイン済か表示
  schedule_asign.getRange('C2').setFormula("IF(COUNTIF(C24:M43, B2)>0, \"Already\", \"Yet\")");

  // 時間割を表示固定
  schedule_asign.setFrozenColumns(2);

  // 確認用
  schedule_asign.getRange('B4').setValue("確認用");
  // 記入用
  schedule_asign.getRange('B23').setValue("記入用")

  // 日付追加
  // 確認用側：細線罫線
  schedule_asign.getRange(5, 3, 1, interview_period.length).setValues([interview_period])
  .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  // 記入用側：細線罫線
  schedule_asign.getRange(24, 3, 1, interview_period.length).setValues([interview_period])
  .setHorizontalAlignment('center').setBorder(false, true, false, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);

  // 時間割
  for (let i = 1; i <= 5; i++){
    // 確認用側結合（A 6:7, 9:10, 12:13, 15:16, 18:19）
    schedule_asign.getRange(3 * i + 3, 1, 2, 1).merge().setValue(i)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 記入用側結合
    schedule_asign.getRange(4 * i + 21, 1, 4, 1).merge().setValue(i)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // 面接時間とExcel関数のセッティング
  const time_table = ["9:10\n|\n9:40", "9:50\n|\n10:20", "11:10\n|\n11:40", "11:50\n|\n12:20", "13:30\n|\n14:00", "14:10\n|\n14:40", "15:15\n|\n15:45", "15:55\n|\n16:25", "17:00\n|\n17:30", "17:40\n|\n18:10"];
  var base_line = 6;
  for (let i = 0; i < time_table.length; i++){
    // 確認用側面接時間
    schedule_asign.getRange(base_line, 2).setValue(time_table[i])
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(null, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID);
    // [スクリプト処理用コピー]と記入用を参照するExcel関数
    for (let j = 3; j < interview_period.length + 3; j++){
      schedule_asign.getRange(base_line, j).setFormulaR1C1("=IF(COUNTIF(R21C,\"*" + Math.trunc(i/2+1) + "限*\"),R2C2,\"\")&char(10)&R" + (2 * i + 25) + "C")
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(null, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
    }
    if (i % 2 == 0){
      base_line += 1;
    } else {
      // 面接可能管理スタッフメモ欄に罫線
      schedule_asign.getRange(base_line+1, 2, 1, interview_period.length + 1)
      .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID); // 細線罫線
      schedule_asign.getRange(base_line+1, 2, 1, interview_period.length + 1)
      .setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE); // 上下に二重線
      base_line += 2
    }

    // 記入用側面接時間
    schedule_asign.getRange(2 * i + 25, 2).setValue(time_table[i]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    schedule_asign.getRange(2 * i + 25, 2, 2, 1)
    .setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID); // 細線罫線
    // 面接者の入力規則を設定
    schedule_asign.getRange(2 * i + 25, 3, 1, interview_period.length).setDataValidation(asign_rule)
    .setBorder(null, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
    // 担当者の入力規則を設定
    schedule_asign.getRange(2 * i + 26, 3, 1, interview_period.length).setDataValidation(interview_rule)
    .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID); // 細線で縦線
    schedule_asign.getRange(2 * i + 26, 3, 1, interview_period.length)
    .setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE); // 二重線で上下線
  }

  // [スクリプト処理用コピー]からの一時参照
  schedule_asign.getRange(base_line, 2).setValue("データ欄").setHorizontalAlignment('center').setVerticalAlignment('middle');
  for (let j = 1; j <= interview_period.length; j++){
    schedule_asign.getRange(base_line, j+2)
    .setFormulaR1C1("=VLOOKUP(R2C2,'スクリプト処理用コピー'!C2:C" + question_label.length + "," + (target + j - 1) + ",FALSE)");
  }

  // 追加罫線
  schedule_asign.getRange(5, 2, 1, interview_period.length+1).setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID); // 細線罫線
  schedule_asign.getRange(5, 2, 1, interview_period.length+1).setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE); // 上下に二重線
  schedule_asign.getRange(24, 2, 1, interview_period.length+1).setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID); // 細線罫線
  schedule_asign.getRange(24, 2, 1, interview_period.length+1).setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE); // 上下に二重線

  // 余分範囲を削除
  remakeSheetRange(schedule_asign, 44, interview_period.length+2);

  /* ------------------------------------------- */
  // [面接シート]
  // [SS用・面接項目]から面接の質問項目を取得
  const interview_data = interview_sheet.getDataRange().getValues();

  // データの１つ１つを読み込み、面接シートに追加、必要に応じてフォームからの転記をセット
  for (let i = 1; i < interview_data.length; i++){
    // [面接シート]での質問項目名
    var interview_title = interview_data[i][0];

    // [スクリプト処理用コピー]から転記を行うか
    var transpose_title = interview_data[i][1];

    if (transpose_title == ""){
      // 転記でなくメモ欄
      var color = "gray";
      interview_report.setRowHeight(i, 100);
    } else {
      // 転記
      var color = "black";
      // 転記元のインデックスを取得
      var transpose_index = question_label.indexOf(transpose_title);
      // 転記関数をセット
      interview_report.getRange(i, 2).setFormulaR1C1("=TRANSPOSE('スクリプト処理用コピー'!R2C" + (transpose_index+1) + ":C" + (transpose_index+1) + ")")
      // シートの保護警告
      interview_report.getRange(i, 2).protect().setWarningOnly(true);
    };

    // [面接シート]のラベルとしてセット
      interview_report.getRange(i, 1).setValue(interview_title).setBackground(color).setFontColor("white")
      .setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID); // 細線罫線;
  }

  // [SS用・面接項目]での表示固定に従って固定
  interview_report.setFrozenColumns(1);
  interview_report.setFrozenRows(interview_sheet.getFrozenRows() - 1);

  // 余分範囲を削除
  remakeSheetRange(interview_report, interview_data.length-1, 2);

  // ラベル列の幅を設定
  interview_report.setColumnWidth(1, 260);
  // 記入列の幅を設定
  interview_report.setColumnWidth(2, 300);

  // シート全体に折り返しと上辺合わせを設定
  interview_report.getRange(1, 1, interview_data.length-1, 2)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('top')
  .setHorizontalAlignment('left');

  /* ------------------------------------------- */
  // CONFIG_for_SHEET
  const sheet_data = config_for_sheet_data.getDataRange().getValues();
  config_for_sheet.getRange(1, 1, sheet_data.length, sheet_data[0].length).setValues(sheet_data);

  /* ------------------------------------------- */
  // CONFIG_for_MAIL
  const mail_data = config_for_mail_data.getDataRange().getValues();
  config_for_mail.getRange(1, 1, mail_data.length, mail_data[0].length).setValues(mail_data);

  /* ------------------------------------------- */
  // 過去応募者
  past_applicants.getRange(1, 1)
  .setFormulaR1C1("=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/" + property.setProperty("DB_ID", db_ID) + "/\", \"過去応募者（2016年度冬〜）!A:D\")");

  /* ------------------------------------------- */
  // 機能メニューの追加
  // 開いた時にメニュー追加の関数が実行されるトリガーをセット
  ScriptApp.newTrigger("form_ss_addMenu")
  .forSpreadsheet(form_ss)
  .onOpen().create();
}

// sheetをrows行columns列に整形
function remakeSheetRange(sheet, rows, columns) {
  // 新しいシートの現在の行数と列数を取得する
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();

  // 余分な行と列を削除する
  // 引数がfalseだったら、削除不要
  if (rows){
    sheet.deleteRows(rows + 1, maxRows - rows);
  }
  if (columns){
    sheet.deleteColumns(columns + 1, maxCols - columns);
  }
}