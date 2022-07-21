// "採用ツール"にメニューの追加
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("採用ツール")
    .addItem("応募フォーム・集計SS 生成&連携", "createTools")
    .addSubMenu(
      ui.createMenu("フォームの開閉")
        .addItem("募集開始", "openform")
        .addItem("募集終了", "closeform")
        .addItem("開始・締切トリガーの更新（延長）", "changeManageSche")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("個別生成")
        .addItem("応募フォームのみ生成", "createForm")
        .addItem("集計SSのみ生成", "createSS")
        .addItem("応募フォーム・集計SS 再連携", "linkFormToSS")
        .addItem("トリガーの再設定", "resetTriger")
    )
    .addToUi();
}