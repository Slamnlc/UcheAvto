function createDiargramm(when){
  let summFromReport;
  let err = '';
  let curDate = new Date();
  let titleText;
  if (isShHave('График')){
    ss.deleteSheet(ss.getSheetByName('График'))
  }
  let arr = [[]];
  let arr1 = [[]];
  let sumColumn =      find('Сумма', uchetSheet.getRange(1,1,1,uchetSheet.getLastColumn()).getValues()[0]) + 1;
  let workTypeColumn = find('Вид работы', uchetSheet.getRange(1,1,1,uchetSheet.getLastColumn()).getValues()[0]) + 1;
  let dateColumn =     find(/Дата*/, uchetSheet.getRange(1,1,1,uchetSheet.getLastColumn()).getValues()[0]) + 1;
  if (when === 'oneMonth'){
    curDate.setDate(1);
    titleText = 'текущий месяц';
  } else if (when === 'threeMonths'){
    curDate.setDate(1);
    curDate.setMonth(curDate.getMonth() - 3);
    titleText = '3 месяца';
  } else if (when === 'halfYear'){
    curDate.setDate(1);
    curDate.setMonth(curDate.getMonth() - 6);
    titleText = 'пол года';
  } else if (when === 'year'){
    curDate.setFullYear(curDate.getFullYear() - 1);
    titleText = 'год';
  } else if (when === 'allTime'){
    curDate.setFullYear(curDate.getFullYear() - 50);
    titleText = 'все время';
  }
  ss.insertSheet(ss.getNumSheets());
  let diagrammSheet = ss.getActiveSheet();
  ss.setActiveSheet(uchetSheet);
  diagrammSheet.setName('График');
  diagrammSheet.getRange(1, 1).setValue('Сумма');
  diagrammSheet.getRange(2, 1).setValue('Вид работы');
  for (let i = 2; i < uchetSheet.getLastRow(); i++){
    if (!uchetSheet.getRange(i, dateColumn).isBlank()){
      if (curDate < uchetSheet.getRange(i, dateColumn).getValue()){
        if (uchetSheet.getRange(i, sumColumn).isBlank()){
          summFromReport = 0;
        } else {
          summFromReport = uchetSheet.getRange(i, sumColumn).getValue();
        }
        arr[0].push(summFromReport);
        arr1[0].push(uchetSheet.getRange(i, workTypeColumn).getValue());
      }
    } else {
      err = err + '\n В ячейке H' + i + ' нет даты';
    }
  }
  diagrammSheet.getRange(1,2,1, arr[0].length).setValues(arr);
  diagrammSheet.getRange(2,2,1, arr[0].length).setValues(arr1);
  if (err != ''){Browser.msgBox(err)};
  drawDiagramm(titleText);
}

function drawDiagramm(textToTitle){
  let diagrammSheet = ss.getSheetByName('График');
  diagrammSheet.hideRow(diagrammSheet.getRange(1,1, diagrammSheet.getLastRow()));
 let diagramm = diagrammSheet.newChart()
 .asColumnChart()
 .addRange(diagrammSheet.getRange(2, 1, 1, diagrammSheet.getLastColumn()))
 .addRange(diagrammSheet.getRange(1, 1,1, diagrammSheet.getLastColumn()))
 .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
 .setTransposeRowsAndColumns(true)
 .setNumHeaders(1)
 .setOption('applyAggregateData', 0)
 .setOption('title', 'Показатели за ' + textToTitle)
 .setOption('useFirstColumnAsDomain', true)
 .setOption('titleTextStyle.color', '#000000')
 .setOption('titleTextStyle.alignment', 'center')
 .setOption('series.0.dataLabel', 'value')
 .setXAxisTitle('Вид работы')
 .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
 .setOption('height', 350)
 .setOption('width', 600)
 .setPosition(1, 1, 0, 0)
 .build();
  diagrammSheet.insertChart(diagramm);
  ss.setActiveSheet(diagrammSheet, true);
}
