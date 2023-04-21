var SHEET_NAME = 'シート1' // フォームを出力するシート名
var SETTING_SHEET_NAME = '設定' // 設定シート名
var URL_GET_SHEET_ID = '1wI4ZkfSsmcHkINtEP3x2iNRbr8pnsvetVbmedECkjOg' // リサーチ者を参照するシートのID
var RC_ROW = 2;     // 作成フォームのレコード開始行
var RC_COL = 1;      // 作成フォームのレコード開始列

// モーダルを開く
function showModal() {

  // 開いているスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // HTMLファイルを取得
  const output = HtmlService.createTemplateFromFile('form');
  const data = spreadsheet.getSheetByName(SHEET_NAME);

  const projectsLastRow = data.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  output.projects = data.getRange(2, 1, projectsLastRow - 1).getValues();

  const html = output.evaluate();
  spreadsheet.show(html);
}

// アップロードボタン
function sendForm(formObject) {
  
  // フォームから受け取ったcsvデータ
  const blob = formObject.myFile;
  const csvText = blob.getDataAsString();
  const values = Utilities.parseCsv(csvText);

  // アップロードするファイル名を取得
  const fileName = blob.getName()

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('シート1');

  // 仕入れ先を取得するシートを取得
  const urlGetSheet = SpreadsheetApp.openById(URL_GET_SHEET_ID)


  // 仕入れ先列全取得
  const suppliers = urlGetSheet.getSheetByName("出品 年月").getRange(2,5,20000,1).getValues(); 

  // ebayURL列全取得
  const ebayURLs = urlGetSheet.getSheetByName("出品 年月").getRange(2,12,20000,1).getValues();

  // 仕入価格列全取得
  var supplierPricies = urlGetSheet.getSheetByName("出品 年月").getRange(2,6,20000,1).getValues();


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
  const soldOutsMsgs = ['', '売り切れました']
  // 設定シート
  const Settings = ss.getSheetByName(SETTING_SHEET_NAME)
  // URLを削除するかどうか
  const deleteFlg = Settings.getRange('A2').getValue() === "ON"

  // １行目は項目名なのでsliceで排除
  values.slice(1).map(function (value) {
    // 在庫状況が空だったら売り切れ判定
    const buyNowBtnMsg = value[1]
    const isSoldOut =  soldOutsMsgs.indexOf(buyNowBtnMsg) !== -1
    // 仕入れ先URL
    const supplierURL = value[2]

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
  urlGetSheet.getSheetByName("出品 年月").getRange(2,6,20000,1).setValues(supplierPricies)
  // flagがONなら仕入れ先URLを商品管理表から削除
  if(deleteFlg){
    // 仕入れ先url更新
    urlGetSheet.getSheetByName("出品 年月").getRange(2,5,20000,1).setValues(suppliers);
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
