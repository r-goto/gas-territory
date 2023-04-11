//指定されたフォルダ内のすべてワンルーム区域のシート名を一括で変更できます。

function changeSheetNames() {
  var allArea = { 
    '高洲': '1_VWlV4zxj6Tt6RLA1cBBhL_2wVE9LVnC',
    '東野': '1SC0LTvqYQBvyFDMKU5Fwbt5kSwM8gSTj',
    '冨岡': '1Rjhq6VTvDfFemim5fxn9zwbtx_6X8ayb',
    '入船': '17-IwLekMA7ay1VlqbZOEHB8XyKfQ9iEn',
    '今川': '13fJoPCoIYDh-ncY0E8NtUJ9Mr4ZHrfJs'
  };

  for (var area in allArea) {
    Logger.log(`フォルダ内のすべてワンルーム区域のシート名を一括で変しています。これには時間がかかります。`)
    var folder = DriveApp.getFolderById(allArea[area]);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      var file = files.next();
      var sheet = SpreadsheetApp.openById(file.getId()).getActiveSheet();
      sheet.setName("部屋番号");
    }
  }
}
