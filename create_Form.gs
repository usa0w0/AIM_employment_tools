// フォーム生成
function createForm() {
  // "採用スケジュール"SS
  const tool_ss = SpreadsheetApp.openById(property.getProperty("Tool_SS_ID"));
  // それぞれのシートを宣言
  var description_sheet = tool_ss.getSheetByName("フォーム用・概要");
  var question_sheet = tool_ss.getSheetByName("フォーム用・質問");

  // parameters
  const parentsFolder_ID = property.getProperty("ParentsFolder_ID");
  const year = property.getProperty("Year");
  const semester = property.getProperty("Semester");
  const interview_period = property.getProperty("Interview_Period");
  const open_date = property.getProperty("募集開始日");
  const close_date = property.getProperty("募集終了日");

  const start_date = interview_period[0];
  const end_date = interview_period[1];

  // 応募フォームのタイトル
  const form_title = year + "年度" + semester + "AIM新規学生スタッフ募集";

  // form = Formオブジェクト
  const form = FormApp.create(form_title);
  // form_ID：formのID
  const form_ID = form.getId()
  // スクリプトのプロパティにformのIDを追加
  PropertiesService.getScriptProperties().setProperty("Form_ID", form_ID);

  // マイドライブに生成されるフォームを削除し、採用フォルダに移動
  const form_File = DriveApp.getFileById(form_ID);
  DriveApp.getFolderById(parentsFolder_ID).addFile(form_File);
  DriveApp.getRootFolder().removeFile(form_File);

  // アクセス権限の設定
  // ログインを必要にしない（AIMに制限しない）
  form.setRequireLogin(false);

  // フォーム作成時には、事故防止のために回答受付を停止
  form.setAcceptingResponses(false);
  // 受付終了メッセージの設定
  var custom_close_message = year + "年度" + semester + "のAIM新規学生スタッフ募集はまだ開始していません。別途、質問等ある方は、sagamipro-contact@aim.aoyama.ac.jpまでご連絡ください。"
  form.setCustomClosedFormMessage(custom_close_message);

  // [フォーム用・概要]シートから、応募概要（年度・期間変更済）を取得
  const description_data = description_sheet.getDataRange().getValues();
  var form_description = ""
  for (let i = 1; i < description_data.length; i++){
    form_description += description_data[i][0] + "\n";
  }
  for (let i = 0; i < description_data.length; i++){
    // 置換対象の出現回数をカウント
    var counta = (form_description.match(new RegExp(description_data[i][2], "g") ) || []).length ;
    for (let j = 0; j < counta; j++){
      form_description = form_description.replace("{"+description_data[i][2]+"}", description_data[i][3])
    }
  }
  // 応募概要をフォームの説明に
  form.setDescription(form_description);

  // フォームにセクション追加
  // セクションのタイトル
  const section_title = year + "年度" + semester + "AIM新規学生スタッフ 応募フォーム";
  // 追加
  form.addPageBreakItem().setTitle(section_title);

  // [フォーム用・質問]シートから、質問とそのプロパティを取得
  const question_data = question_sheet.getDataRange().getValues()
  // for文でそれぞれの質問を追加していく
  for (let i = 1; i < question_data.length; i++){
    // それぞれの要素を変数に格納
    // console.log(question_data[i])
    // 質問タイトル
    var question_title = question_data[i][0]
    // 補助テキスト
    var question_helptext = question_data[i][1]
    // 質問の形式
    var question_type = question_data[i][2]
    // 必須回答
    var question_required = question_data[i][3]
    // 回答の検証：正規表現
    var question_validation = question_data[i][4]
    // チェックボックス選択肢（カンマ区切り）
    var choice_list = question_data[i][5].split(",")

    // 質問の形式に応じてGASで質問を追加
    if (question_type == "記述式" && question_validation == ''){
      form.addTextItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setRequired(question_required);
    } 
    else if (question_type == "記述式"){
      form.addTextItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setRequired(question_required)
      .setValidation(FormApp.createTextValidation().requireTextMatchesPattern(question_validation).build());
    } 
    else if (question_type == "段落"){
      form.addParagraphTextItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setRequired(question_required);
    } 
    else if (question_type == "ラジオボタン"){
      form.addMultipleChoiceItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setChoiceValues(choice_list)
      .setRequired(question_required);
    } 
    else if (question_type == "チェックボックス"){
      form.addCheckboxItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setChoiceValues(choice_list)
      .setRequired(question_required);
    } 
    else if (question_type == "プルダウン"){
      form.addListItem()
      .setTitle(question_title)
      .setHelpText(question_helptext)
      .setChoiceValues(choice_list)
      .setRequired(question_required);
    } 
    else if (question_type == "面接日程"){
      // 質問と概要を記載
      question_helptext = question_helptext.replace("{面接開始日}", start_date).replace("{面接締切日}", end_date);
      form.addSectionHeaderItem().setTitle(question_title).setHelpText(question_helptext);
      // グリッド式チェックボックス
      form.addCheckboxGridItem()
      .setRows(interview_period)
      .setColumns(choice_list);
    }
    else if (question_type == "タイトルと説明"){
      form.addSectionHeaderItem().setTitle(question_title).setHelpText(question_helptext);
    }
  }

  // 回答後の案内
  form.setConfirmationMessage("エントリーを受け付けました。 後日、面接の日時につきまして、ご記入いただいたメールアドレス宛にご連絡させていただきます。今しばらくお待ちください。");

  // 自動で回答の受付・締切を制御するトリガーをセット
  // 指定の日付と時刻でトリガーをセットする
  ScriptApp.newTrigger("openform")
  .timeBased()
  .at(open_date)
  .create();
  ScriptApp.newTrigger("closeform")
  .timeBased()
  .at(close_date)
  .create();

  // QRコード生成
  createQRCode(form.getPublishedUrl(), form_title, parentsFolder);
}