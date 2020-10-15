function createStyleForMainFromCell(){
  let mainCellsArray = [makerCell, modelCell,nomerCell,vinCell,yearCell,probegCell,workVidCell,workTypeCell,ownerCell, phoneCell, summCell]
  for (let i = 0; i < mainCellsArray.length; i++){
    mainCellsArray[i].setFontSize(17)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setFontColor('#000000')
  .setBackground(null)
  .setFontFamily(null);
  }
  let addictionalCellArray = [nomerCell, nomerCell,vinCell,workVidCell,workTypeCell,ownerCell,yearCell,probegCell]
  for (let i = 0; i < addictionalCellArray.length; i++){
    if (i === 6 || i === 7){
      addictionalCellArray[i].setNumberFormat('0')
    } else {
      addictionalCellArray[i].setNumberFormat('@');
    }
  }
  summCell.setNumberFormat('#,##0');
  phoneCell.setNumberFormat('"0"0');
}

function firstLetterUpper(val){
  let ret = '';
  if (val != null && val != ''){
    let str = String(val);
  let arr = str.split(' ');
  if (arr.length === 1){
    if (str.match(/[A-Za-zА-Яа-я]/g) !== null){
      if (str.match(/[A-Za-zА-Яа-я]/g).length > 3){
        ret = str[0].toUpperCase() + str.toLowerCase().slice(1);
      } else {
        ret = str;
      }
    } else {
      ret = str;
    }
  } else {
    for (let i = 0; i < arr.length; i++){
      if (ret === ''){
      ret = firstLetterUpper(arr[i]);
      } else {
        ret = ret + ' ' + firstLetterUpper(arr[i]);
      }
    }
  }
  } else {
    ret = '';
  }
  return ret;
}

function claerMainCells(){
  let mainCellArray = [makerCell,modelCell,nomerCell,vinCell,probegCell,yearCell,workVidCell,workTypeCell,ownerCell,phoneCell,summCell]
  for (i = 0; i < mainCellArray.length; i++){
    mainCellArray[i].clearContent();
  }
}

function replaceEngSymb(str){
  let engSymbArray =  [/E/g,/e/g,/T/g,/t/g,/Y/g,/y/g,/I/g,/i/g,/O/g,/o/g,/P/g,/p/g,/A/g,/a/g,/H/g,/h/g,/K/g,/k/g,/X/g,/x/g,/C/g,/c/g,/B/g,/b/g,/M/g,/m/g];
  let rusSymbArray = ["Е","Е","Т","Т","У","У","І","І","О","О","Р","Р","А","А","Н","Н","К","К","Х","Х","С","С","В","В","М","М"];
  let st = String(str);
  for (let i = 0; i < engSymbArray.length; i++){
    st = st.replace(engSymbArray[i], String(rusSymbArray[i]));
  }
  return st
}

function correctVinAndNumColumn(){
  for (i = 2; i < uchetSheet.getLastRow() + 2; i++){
    if (!uchetSheet.getRange(i, 4).isBlank()){
      let str = String(uchetSheet.getRange(i, 4).getValue()).toUpperCase();
      if (str.match(/[A-Z]/) !== null) {
          uchetSheet.getRange(i, 4).setValue(replaceEngSymb(uchetSheet.getRange(i, 4).getValue()))
        }
    }
    if (!uchetSheet.getRange(i, 5).isBlank()){
    uchetSheet.getRange(i, 5).setValue(String(uchetSheet.getRange(i, 5).getValue()).toUpperCase());
    }
  }
}

function clearSubTotalFormulas(){
  let columnToCheck = findColumnByName(/Дата*/);
  let formulasArray = uchetSheet.getRange(2,columnToCheck, uchetSheet.getLastRow()).getFormulas();
  for (i = 0; i < formulasArray.length; i++){
    if (String(formulasArray[i]).match(/SUBTOTAL/) != null){
      uchetSheet.getRange(i + 2, columnToCheck).clearContent();
    }
  }
}

function autoFitAllColumns(sheetName){
  sheetName.autoResizeColumns(2, sheetName.getLastColumn());
}

function addSumToVitag(){
  let count = countNonEmpty(vipiskaSheet.getRange(5,1, vipiskaSheet.getLastRow()).getValues());
  let columnWithSumma = findColumnByName('Сумма');
  let summa = sum(vipiskaSheet.getRange(6, columnWithSumma, vipiskaSheet.getLastRow()).getValues());
  createStyleForCell(vipiskaSheet.getRange(vipiskaSheet.getLastRow() + 2, 1), 14, true, 'left', 'middle',false);
  vipiskaSheet.getRange(vipiskaSheet.getLastRow() + 2, 1).setValue('Количество записей: ' + count + '. Сумма: ' + summa);
}

function sum(array){
  let x = 0
  for (let i = 0; i < array.length; i++){
    x += Number(array[i])
  }
  return x
}

function countNonEmpty(array){
  let x = 0
  for (let i = 0; i < array.length; i++){
    if (array[i] !== ''){
      ++x
    }
  }
  return x
}

function count(what, where){
  let x = 0;
  for (let i = 0; i < where.length; i++){
    if (where[i] !== ''){
      if (where[i] === what){++x}
    }
  }
  return x
}

function clearFilter(){
  if (uchetSheet.getFilter() !== null){
    uchetSheet.getFilter().remove();
  }
}

function isShHave(SheetName){
  for (let i = 0; i < ss.getNumSheets(); i++){
    if (ss.getSheets()[i].getSheetName() === SheetName){
      return true
    }
  }
  return false
}

function find(what, where){
  for (let i = 0; i<where.length; i++){
    if (String(where[i]).match(what)){
      return i;
    }
  }
  return -1
}

function addModelToList(){
  let x =  modelCell.getDataValidation().getCriteriaValues()[0].getValues()[0].indexOf(modelCell.getValue());
  if (x === -1){
    let findCell = avtoSheet.getRange(1, 1, 1, avtoSheet.getLastColumn()).getValues()[0].indexOf(makerCell.getValue());
    if (findCell !== - 1){
      findCell++;
      avtoSheet.getRange(2, findCell)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
        .offset(1, 0)
        .setValue(modelCell.getValue());
    }
  }
}

function addWorkTypeToList(){
  let x =  workTypeCell.getDataValidation().getCriteriaValues()[0].getValues().toString().split(',').indexOf(workTypeCell.getValue());
  if (x === -1){
    let findCell = workSheet.getRange(1, 1, 1, workSheet.getLastColumn()).getValues()[0].indexOf(workVidCell.getValue());
    if (findCell !== - 1){
      findCell++;
      if (!workSheet.getRange(2, findCell).isBlank()){
        workSheet.getRange(2, findCell).setValue(workTypeCell.getValue());
      } else {
      workSheet.getRange(1, findCell).getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(1, 0).setValue(workTypeCell.getValue());
    }
  }
}
}

function createValidationForModel(){
  let findCell = avtoSheet.getRange(1, 1, 1, avtoSheet.getLastColumn()).getValues()[0].indexOf(makerCell.getValue());
    if (findCell !== -1){
      findCell++;
      modelCell.clearDataValidations();
      modelCell.setDataValidation(SpreadsheetApp.newDataValidation()
       .requireValueInRange(avtoSheet.getRange(2, findCell, avtoSheet.getRange(1, findCell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1, 1), true)
       .build());
    }
}

function createValidationForWorkType(findValue,cellRange){
    let findCell = workSheet.getRange(1, 1, 1, workSheet.getLastColumn()).getValues()[0].indexOf(findValue);
    if (findCell !== -1){
      findCell++;
      cellRange.clearDataValidations();
      cellRange.setDataValidation(SpreadsheetApp.newDataValidation()
       .requireValueInRange(workSheet.getRange(2, findCell, avtoSheet.getRange(1, findCell).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() + 1, 1), true)
       .build());
    }
}

function analizeCarNumber(){
  let num = replaceEngSymb(String(nomerCell.getValue()).toUpperCase());
  let i = count(num, uchetSheet.getRange(2, 4, uchetSheet.getLastRow()).getValues());
  if (i !== 0){
    nomerCell.offset(0, 1).setValue('Количество записей с таким номером: ' + i);
  } else {
    nomerCell.offset(0, 1).clearContent();
  }
}

function analizeVin(){
  if (!vinCell.isBlank()){
    let vin = String(vinCell.getValue()).toUpperCase();
    let i = count(vin, uchetSheet.getRange(2, 5, uchetSheet.getLastRow()).getValues());
    if (i !== 0){
    vinCell.offset(0, 1).setValue('Количество записей с таким VIN: ' + i);
  } else {
    vinCell.offset(0, 1).clearContent();
  }
  }
}

function findColumnByName(searchedColumnHeader){
  let x = find(searchedColumnHeader, uchetSheet.getRange(1,1,1,uchetSheet.getLastColumn()).getValues()[0]) + 1;
  return x;
}

function createStyleForCell(cell, fontSize, fontBold, horizontalAllignment, vertivalAllignment, wrap){
  cell.setFontSize(fontSize);
  if (fontBold) {
    cell.setFontWeight('bold')
  } else {
    cell.setFontWeight(null);
  }
  cell.setHorizontalAlignment(String(horizontalAllignment));
  cell.setVerticalAlignment(String(vertivalAllignment));
  if (wrap) {
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  } else {
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  }
}
