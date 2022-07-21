// Dateオブジェクトから和暦表示に
function getJapaneseDate(date) {
  var m = date.getMonth() + 1;
  var dd = date.getDate();
  var week = date.getDay();
  const weekly = ["日","月","火","水","木","金","土"];
  week = weekly[week]
  var str_date = m + "月" + dd + "日" + "(" + week + ")";
  return str_date;
}