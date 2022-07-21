// QRコード生成
function createQRCode(link, name, folder_ID) {
  let url = 'https://chart.googleapis.com/chart?chs=100x100&cht=qr&chl=' + link;
  let option = {
    method: "get",
    muteHttpExceptions: true
  };
  let ajax = UrlFetchApp.fetch(url, option);

  let fileBlob = ajax.getBlob()
  let folder = DriveApp.getFolderById(folder_ID);
  folder.createFile(fileBlob).setName(name + '.png');
}