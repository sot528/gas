// 今選んでいるセルからフォームに規則性を持って選択する項目を追加してゆく
function createForms() {
  // フォーム作成
  //const TITLE = 'DX推進指標:システム部分'
  //var form = FormApp.create(TITLE);
  //form.setDescription('経産省:DX推進指標(http://bit.ly/2ON1bT0)よりITシステム部のみ抜粋。各項目で、現状に即した選択肢を選んでください。(所要時間10分以下)')
  //return;

  var form = FormApp.openById('xxx');

  const SHEET = SpreadsheetApp.getActiveSheet()
  const ACTIVE_CELL = SHEET.getActiveCell();
  const LIST_ITEM_TITLE =  SHEET.getRange(ACTIVE_CELL.getRow(), ACTIVE_CELL.getColumn()).getValue().trim();
  const LIST_CHOICES_COLMUN = ACTIVE_CELL.getColumn() + 5
  const LIST_CHOICES_ROW = ACTIVE_CELL.getRow() + 3
  
  var listItem = form.addListItem();
  listItem.setTitle(LIST_ITEM_TITLE)
  .setRequired(true)

  listItem = getListItemChoice(listItem, SHEET, LIST_CHOICES_COLMUN, LIST_CHOICES_ROW);
  
  //次の質問を選択
  SHEET.getRange(ACTIVE_CELL.getRow() + 17, ACTIVE_CELL.getColumn()).activate()
  
}

// 選択項目を取得
function getListItemChoice(listItem, sheet, colmun, row){

  var choices = []
  for (var i=0; i<5; i++){
    var value = sheet.getRange(row + i, colmun).getValue().trim() + sheet.getRange(row + i, colmun + 1).getValue().trim();
    choices.push(listItem.createChoice(value))
  }
  return listItem.setChoices(choices);
}
