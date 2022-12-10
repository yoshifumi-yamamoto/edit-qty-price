var SHEET_NAME = 'シート2' // 出力するシート名
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

  const projectsLastRow = data.getRange(1, 1).getNextDataCell(spreadsheetApp.Direction.DOWN).getRow();
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

  //データがある最終列を取得（上手くいってない）
  const lastCol = urlGetSheet.getLastColumn();
  console.log('lastCol')
  console.log(lastCol)

  // 仕入れ先列全取得
  const suppliers = urlGetSheet.getSheetByName("出品 年月").getRange(2,5,5999,1).getValues(); 

  // ebayURL列全取得
  const ebayURLs = urlGetSheet.getSheetByName("出品 年月").getRange(2,12,5999,1).getValues();



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
  
  // １行目は項目名なのでsliceで排除
  values.slice(1).map(function (value) {
    // 在庫状況が空だったら売り切れ判定
    const isSoldOut = value[1] === ''
    if(isSoldOut){
      // 売り切れだったら配列にurlを追加
      soldOuts.push(value[2])
    }
  })

  soldOuts.map(function(soldOut) {
    // 仕入れ先と同じ行のitemNumberを取得する
    const urlRow = formattedSuppliers.indexOf(soldOut)
    const itemNumber =  formattedEbayURLs[urlRow].replace('https://www.ebay.com/itm/', '')

    // Action(Revise = 変更), itemNumber, qty
    addValues.push(['Revise', itemNumber, 0])
  })
  
  // 既存レコードをクリアし、CSVのレコードを貼り付け
  // clearRecords(RC_ROW, RC_COL, sheet);
  sheet.getRange(RC_ROW, RC_COL, addValues.length, addValues[0].length).setValues(addValues);
}
