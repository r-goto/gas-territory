/* 
----------------------------------------------------------------------------------------------------------------------------------------
指定されたフォルダ内の全ワンルーム区域に保護設定と編集警告設定を一括で設定できます。
区域の数が多すぎるためタイムアウト(6分)してしまいます。何度かリトライさせる必要があります。
リトライ時は、既に処理が完了している(保護定義が6つ設定済み)シートは、処理をスキップするようにしています。
----------------------------------------------------------------------------------------------------------------------------------------
*/

function setAllProtections() {
  var allArea = { 
    '高洲': '1jogpzUlhJtwHhkO2GaGwjJY1syy6rc0F',
    '東野': '1-AgJ7VlwdqAHzaI2qP4AZinl4vQQXW9S',
    '冨岡': '15uj1soY7cJhV3hpbyGGRVlCYhMqDAqzN',
    '入船': '1Tc3p2LI3BZXfLc5BVAkIoZMNtfmr4lFx',
    '今川': '1K-r5w6jDF0UNzNHujwCcuk692RfD6Wsi'
  };

  for (var area in allArea) {
    var folder = DriveApp.getFolderById(allArea[area]);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    Logger.log(`${DriveApp.getFolderById(allArea[area]).getName()}のワンルーム区域にデータ保護の設定をしています。これには時間がかかります。\nタイムアウトが発生する場合、すべての区域への処理が完了するまでリトライしてください。`)

    while (files.hasNext()) {
      var file = files.next();
      var sheet = SpreadsheetApp.openById(file.getId()).getActiveSheet();
      // 保護ルールの数が足りている場合は、そのシートはループをスキップさせる。
      var count = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length;
      if ( count == 6){
        Logger.log(`${SpreadsheetApp.openById(file.getId()).getName()}には、既に${count}のルールが既に存在するためループ処理をスキップ`);
        continue
      }

      // 必要な保護設定を追加していきます。

      var target1 = sheet.getRange(1, 1, 2, sheet.getLastColumn()).protect(); // 1行目と2行目を取得。
      var target2 = sheet.getRange(1,1, 100).protect(); //一(A)列目の100行目までを取得。
      var target3 = sheet.getRange(3,3, 100).protect().setDescription('編集前に警告を表示');; //3(C)列目の3行目から100行目までを取得。
      var target4 = sheet.getRange(3,5, 100).protect().setDescription('編集前に警告を表示');; //5(E)列目の3行目から100行目までを取得。
      var target5 = sheet.getRange(3,7, 100).protect().setDescription('編集前に警告を表示');; //7(G)列目の3行目から100行目までを取得。
      var target6 = sheet.getRange(3,8, 100).protect().setDescription('編集前に警告を表示');; //8(H)列目の3行目から100行目までを取得。
      
      // 1行目と2行目、及び一列目の100行目までを編集不可に設定します。
      target1.removeEditors(target1.getEditors());
      target2.removeEditors(target2.getEditors());

      // 「日時」と「拒否」を編集する前に警告文を表示させるように設定します。
      target3.setWarningOnly(true)
      target4.setWarningOnly(true)
      target5.setWarningOnly(true)
      target6.setWarningOnly(true)
    }
    Logger.log(`${DriveApp.getFolderById(allArea[area]).getName()} のすべての保護設定が完了しました`)
  }
}

/* 
----------------------------------------------------------------------------------------------------------------------------------------
指定されたフォルダ内の全ワンルーム区域の保護設定と編集警告設定を一括で解除できます。
----------------------------------------------------------------------------------------------------------------------------------------
*/

function removeAllProtections() {
  var allArea = { 
    '高洲': '1jogpzUlhJtwHhkO2GaGwjJY1syy6rc0F',
    '東野': '1-AgJ7VlwdqAHzaI2qP4AZinl4vQQXW9S',
    '冨岡': '15uj1soY7cJhV3hpbyGGRVlCYhMqDAqzN',
    '入船': '1Tc3p2LI3BZXfLc5BVAkIoZMNtfmr4lFx',
    '今川': '1K-r5w6jDF0UNzNHujwCcuk692RfD6Wsi'
  };

  for (var area in allArea) {
    var folder = DriveApp.getFolderById(allArea[area]);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      var file = files.next();
      var sheet = SpreadsheetApp.openById(file.getId()).getActiveSheet();
      var count = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length;
      if ( count == 0){
        Logger.log(`${SpreadsheetApp.openById(file.getId()).getName()}にはルールが存在しないため、処理をスキップ`);
        continue
      }
      var all_protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < all_protections.length; i++) {
        var protection = all_protections[i];
        if (protection.canEdit()) {
          protection.remove();
        }
      }
    }
  }
}
