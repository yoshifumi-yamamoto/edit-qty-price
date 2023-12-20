var SHEET_NAME = 'シート1' // フォームを出力するシート名
var SETTING_SHEET_NAME = '設定' // 設定シート名
var URL_GET_SHEET_ID = '1SZ2lXSoSunmNWmiJqPN3vlvPVCOdUgDmWLkP-gyj9uo' // 仕入れ先を参照するシートのID
var RC_ROW = 2;     // 作成フォームのレコード開始行
var RC_COL = 1;      // 作成フォームのレコード開始列

// ドライブ内にあるcsvデータ全取得し、特定の条件で列を並べ替える
function extractDataFromCSVFiles() {
  var folderId = "1K92RsdR2OT3wh6nGt_EXKv0zs-OjsX4U";  // 抽出したいCSVファイルが含まれるフォルダのIDを指定します
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.CSV);
  var data = [];  // 抽出したデータを格納するための配列

  while (files.hasNext()) {
    var file = files.next();

    // ファイル名に「メルカリ」が含まれているかチェック
    var isMercariFile = file.getName().indexOf('メルカリ') !== -1;
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ',');

    // CSVデータのヘッダー行から各列のインデックスを取得
    var headers = csvData[0];
    var productIndex = headers.indexOf('商品名');
    var stockIndex = headers.indexOf('在庫');
    var keywordIndex = headers.indexOf('店铺URL');
    var priceIndex = headers.indexOf('価格');

    // ヘッダーを除去してCSVファイルのデータを並べ替えて配列に追加
    for (var i = 1; i < csvData.length; i++) {
      var row = csvData[i];
      var reorderedRow = isMercariFile ? [
        row[productIndex], // 商品名
        row[stockIndex],   // 在庫
        row[keywordIndex], // Keyword
        row[priceIndex]    // 価格
      ] : row;

      data.push(reorderedRow);
    }
  }

  // 抽出したデータを返す場合
  return data;
}


// アップロードボタン
function sendForm() {

  const values = extractDataFromCSVFiles()

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('シート1');

  // 仕入れ先を取得するシートを取得
  const urlGetSheet = SpreadsheetApp.openById(URL_GET_SHEET_ID)


  // 仕入れ先列全取得
  const suppliers = urlGetSheet.getSheetByName("出品 年月").getRange(2,5,30000,1).getValues(); 

  // ebayURL列全取得
  const ebayURLs = urlGetSheet.getSheetByName("出品 年月").getRange(2,12,30000,1).getValues();

  // 仕入価格列全取得
  var supplierPricies = urlGetSheet.getSheetByName("出品 年月").getRange(2,6,30000,1).getValues();


  // 二次元配列を一次元配列に変換

  const formattedSuppliers = suppliers.reduce(function (acc, cur, i) {
    return acc.concat(cur);
  });

  const formattedEbayURLs = ebayURLs.reduce(function (acc, cur, i) {
    return acc.concat(cur);
  });

  //在庫切れ
  var soldOuts = []

  // 2次元配列に整形
  var addValues = []
  // 売り切れ判定
  const soldOutsMsgs = ['', '売り切れました', 'ただいま売り切れ中です']
  // 設定シート
  const Settings = ss.getSheetByName(SETTING_SHEET_NAME)
  // URLを削除するかどうか
  const deleteFlg = Settings.getRange('A2').getValue() === "ON"

  values.map(function (value) {
    // 在庫状況が空だったら売り切れ判定
    const buyNowBtnMsg = value[1]
    const isSoldOut =  soldOutsMsgs.indexOf(buyNowBtnMsg) !== -1
    // 仕入れ先URL
    const supplierURL = value[2]
    // console.log("仕入れ先URL", supplierURL)
    // 仕入れ先と同じ行のitemNumberを取得する
    const itemRow = formattedSuppliers.indexOf(supplierURL)
    const itemNumber =  formattedEbayURLs[itemRow] ? formattedEbayURLs[itemRow].replace('https://www.ebay.com/itm/', '') : formattedSuppliers[itemRow]
    const latestPrice = value[3]
    if(isSoldOut){
      // 売り切れだったらqtyを0
      // Action(Revise = 変更), itemNumber, qty
      addValues.push(['Revise', itemNumber, 0])
      // 仕入れ先を配列から削除
      if(itemRow > -1){
        suppliers[itemRow][0] = ''
      }
    }else{
      // 最新の仕入価格を配列に上書き
      // 商品が見つからない場合のエラー回避
      if(supplierPricies[itemRow]){
        if(supplierPricies[itemRow][0] !== 0.0){
          supplierPricies[itemRow][0] = latestPrice
        }
      }
      addValues.push(['Revise', itemNumber, 1])
    }
  })
  // 仕入れ価格の更新
  urlGetSheet.getSheetByName("出品 年月").getRange(2,6,30000,1).setValues(supplierPricies)
  // flagがONなら仕入れ先URLを商品管理表から削除
  if(deleteFlg){
    // 仕入れ先url更新
    urlGetSheet.getSheetByName("出品 年月").getRange(2,5,30000,1).setValues(suppliers);
  }
  // 書き込みを最終行以降から開始
  const startRow = sheet.getLastRow()+1
  // 既存レコードをクリアし、CSVのレコードを貼り付け
  // clearRecords(RC_ROW, RC_COL, sheet);
  sheet.getRange(startRow, RC_COL, addValues.length, addValues[0].length).setValues(addValues);
}

function clearSheet1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
  range.clearContent();
}
