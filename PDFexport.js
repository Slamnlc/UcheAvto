function encodeDate(yy,mm,dd,hh,ii,ss){
  var days=[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if(((yy % 4) == 0) && (((yy % 100) != 0) || ((yy % 400) == 0)))days[1]=29;
  for(var i=0; i<mm; i++)dd+=days[i];
  yy--;
  return ((((yy * 365 + ((yy-(yy % 4)) / 4) - ((yy-(yy % 100)) / 100) + ((yy-(yy % 400)) / 400) + dd - 693594) * 24 + hh) * 60 + ii) * 60 + ss)/86400.0;
}

function exportPDF(ssID,source,options,format){
  var dt=new Date();
  var d=encodeDate(dt.getFullYear(),dt.getMonth(),dt.getDate(),dt.getHours(),dt.getMinutes(),dt.getSeconds());
  var pc=[null,null,null,null,null,null,null,null,null,0,
          source,
          10000000,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
          d,
          null,null,
          options,
          format,
          null,0,null,0];

 var js = " \
    <script> \
      window.open('https://docs.google.com/spreadsheets/d/"+ssID+"/pdf?id="+ssID+"&a=true&pc="+JSON.stringify(pc)+"&gf=[]'); \
      google.script.host.close(); \
    </script> \
  ";
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(10)
    .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, "Save To PDF");
}


function myExportPDF(){

  let newNm;
  if (!vipiskaSheet.getRange('A2').isBlank()){
    newNm = 'Выписка для ' + vipiskaSheet.getRange('A2').getValue();
  } else if (!vipiskaSheet.getRange('B2').isBlank()){
    newNm = 'Выписка для ' + vipiskaSheet.getRange('B2').getValue();
  } else if (!vipiskaSheet.getRange('C2').isBlank()){
    newNm = 'Выписка для ' + vipiskaSheet.getRange('C2').getValue();
  }
  vipiskaSheet.setName(newNm);
  exportPDF(ss.getId(), // Идентификатор таблицы
    [
      [vipiskaSheet.getSheetId().toString(), // ID листа в виде строки
       3,                             // начальная граница по вертикали (с первой строки)
       vipiskaSheet.getLastRow(),                           // конечная граница по вертикали (по 160-ую строку включительно)
       0,                             // начальная граница по горизонтали (с ячейки A)
       14                              // конечная граница по горизонтали (по ячейку H включительно)
      ]
    ],
    [
      0,         // Не показывать заметки
      null,
      1,         // Показывать линии сетки
      0,         // Не показывать номера страниц
      0,         // Не показывать название книги
      0,         // Не показывать название листа
      0,         // Не показывать текущую дату
      0,         // Не показывать текущее время
      1,         // Повторять закрепленные строки
      1,         // Повторять закрепленные столбцы
      1,         // Порядок страниц вниз, затем вверх
      1,
      null,
      null,
      1,         // Горизонтальное выравниване по левому краю
      1          // Вертикальное выравнивание по верхнему краю
    ],
    [
      "A4",      // Фрмат листа A4
      1,         // Ориентация страницы вертикальная
      2,         // Выровнять по высоте
      1,
      [
        0.75,    // Отступ сверху 0.75 дюймов
        0.75,    // Отступ снизу 0.75 дюймов
        0.7,     // Отступ слева 0.7 дюйма
        0.7      // Отступ справа 0.7 дюйма
      ]
    ]
  );
  //Utilities.sleep(4000);
  vipiskaSheet.setName('Выписка');
}
