var INPUT_SHEET_NAME = '出品 年月' // 取得するシート名
var OUTPUT_SHEET_NAME = '在庫管理' // 出力するシート名
var URL_GET_SHEET_ID = '1wI4ZkfSsmcHkINtEP3x2iNRbr8pnsvetVbmedECkjOg' // リサーチ者を参照するシートのID
var ss = SpreadsheetApp.getActive(); // 現在のシートを取得
var RC_ROW = 2;     // 作成フォームのレコード開始行
var RC_COL = 1;      // 作成フォームのレコード開始列



function output() {
  // 出力先のシートを取得
  const sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  const data = input()

  // 1次元配列を2次元配列化
  var addValues = []
  data.map(function (d, i) {
    addValues[i] = [i + 1, d]
  })
  // 現在日時を取得
  var today = new Date();
  // Date型データをフォーマット
  var todayStr = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd HH:mm:ss');
  // 最終更新日を出力
  sheet.getRange('D1').setValue(todayStr);

  sheet.getRange(RC_ROW , RC_COL, addValues.length, addValues[0].length).setValues(addValues);
}

// 仕入れ先の取得
function input () {
  // 仕入れ先を取得するシートを取得
  const urlGetSheet = SpreadsheetApp.openById(URL_GET_SHEET_ID).getSheetByName(INPUT_SHEET_NAME)
  // 仕入れ先を取得
  const suppliers = urlGetSheet.getRange(4,5,5999,1).getValues();
  const formattedSuppliers = suppliers.reduce(function (acc, cur, i) {
    return acc.concat(cur);
  });

  // 空を排除しソート
  const exclusions = formattedSuppliers.filter(function(sup){return !!sup}).sort()
  return exclusions

}