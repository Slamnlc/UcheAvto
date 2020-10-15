function createCustomMenu(){
 let ui = SpreadsheetApp.getUi();
 ui.createMenu('Макросы')
 .addSubMenu(ui.createMenu('Фильтрация')
             .addItem('Текущий месяц', 'filterOneMonth')
             .addItem('3 месяца', 'filterThreeMonth')
             .addItem('Пол года', 'filterHalfYear')
             .addItem('Текущий год', 'filterYear')
             .addItem('Сбросить фильт', 'delFilter'))
 .addSubMenu(ui.createMenu('Нарисовать график')
             .addItem('Текущий месяц', 'createDiargrammCurrentMonth')
             .addItem('3 месяца', 'createDiargrammThreeMonth')
             .addItem('6 месяцев', 'createDiargrammHalfYear')
             .addItem('Год', 'createDiargrammYear')
             .addItem('Все время', 'createDiargrammAllTime'))
 .addSeparator()
 .addItem('Причесать таблицу','makeTableGreatAgain')
 .addItem('Экспорт выписки в PDF', 'pdfExport')
 .addToUi();
}

function filterOneMonth(){
  filterCells('oneMonth');
}

function filterThreeMonth(){
  filterCells('threeMonth');
}

function filterHalfYear(){
  filterCells('halfYear');
}

function filterYear(){
  filterCells('year');
}

function delFilter(){
  clearFilter();
  clearSubTotalFormulas();
}

function createDiargrammCurrentMonth(){
  createDiargramm('oneMonth');
}

function createDiargrammThreeMonth(){
  createDiargramm('threeMonths');
}

function createDiargrammHalfYear(){
  createDiargramm('halfYear');
}

function createDiargrammYear(){
  createDiargramm('year');
}

function createDiargrammAllTime(){
  createDiargramm('allTime');
}

function filterCells(when){
  let date = new Date();
  let columnWithDate = find(/Дата*/, uchetSheet.getRange(1,1,1,uchetSheet.getLastColumn()).getValues()[0]) + 1;
  if (when === 'oneMonth'){
    date.setDate(1);
  } else if (when === 'threeMonth'){
    date.setDate(1);
    date.setMonth(date.getMonth() - 3);
  } else if (when === 'halfYear'){
    date.setDate(1);
    date.setMonth(date.getMonth() - 6);
  } else if (when === 'year'){
    date.setFullYear(date.getFullYear() - 1);
  }
  let filterCriteria = SpreadsheetApp.newFilterCriteria()
  .whenDateAfter(date)
  .build();
  clearSubTotalFormulas();
  clearFilter();
  uchetSheet.getRange(1,1, uchetSheet.getLastRow(), uchetSheet.getLastColumn())
       .createFilter();
  uchetSheet.getFilter()
       .setColumnFilterCriteria(columnWithDate, filterCriteria);
  uchetSheet.getRange(uchetSheet.getLastRow() + 1, columnWithDate)
  .setFormula('=SUBTOTAL(9;' + uchetSheet.getRange(2,columnWithDate).getA1Notation() + ':' + uchetSheet.getRange(uchetSheet.getLastRow(),columnWithDate).getA1Notation() + ')');
}

function makeTableGreatAgain(){
  let err = 'Дата не указана в ячейках:\n';
  let x = 0;
  let errCount = 0;
  let searchArr = [/*записи*/, "Марка", "Модель", "Гос номер", "VIN", /Пробег*/, "Год", "Дата ТО", "Сумма", "Вид работы", "Тип работы", "Владелец", "Телефон"];
  let columnsHeadersArray = uchetSheet.getRange(1, 1, 1, uchetSheet.getLastColumn()).getValues()[0];
  let dateColumn = findColumnByName(/Дата*/);
  clearFilter();
  clearSubTotalFormulas();
  for (let i = 2; i < uchetSheet.getLastRow(); i++){
    if (uchetSheet.getRange(i,1,1,8).getValues().toString().replace(/,/g, '') === ''){
      uchetSheet.deleteRows(i);
      --i;
    }
  }
  for (let i = 2; i < uchetSheet.getLastRow(); i++){
    if (uchetSheet.getRange(i,dateColumn).isBlank()){
      err = err + '\n - ' + uchetSheet.getRange(i,dateColumn).getA1Notation();
      ++ errCount;
      uchetSheet.getRange(i,dateColumn).setBackground('#ff0000');
    } else if (uchetSheet.getRange(i,dateColumn).getBackground() === '#ff0000'){
      uchetSheet.getRange(i,dateColumn).setBackground(null);
    }
  }
  correctVinAndNumColumn();
  x = uchetSheet.getLastRow();
  for (i = 0; i < searchArr.length; i++){
        let y = find(searchArr[i], columnsHeadersArray);
        y++;
        if (i === 0 || i === 5 || i === 6){
          uchetSheet.getRange(1, y, x).setNumberFormat('0');
        } else if (i === 1 || i === 2 || i === 3 || i === 4 || i === 9 || i === 10  || i === 11 || i === 13){
          uchetSheet.getRange(1, y, x).setNumberFormat('@');
        } else if (i === 7){
          uchetSheet.getRange(1, y, x).setNumberFormat('dd mmmm yyyy');
        } else if (i === 8){
          uchetSheet.getRange(1, y, x).setNumberFormat('#,##0');
        } else if (i === 12){
          uchetSheet.getRange(1, y, x).setNumberFormat('"0"0');
        }
  }
  uchetSheet.getRange(2,1, uchetSheet.getLastRow()-1, 14).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  if (errCount !== 0){SpreadsheetApp.getUi().alert(err + '\n\n Общее количество ячеек: ' + errCount)}
}

function pdfExport(){
  myExportPDF();
}
