let ss = SpreadsheetApp.getActiveSpreadsheet();
//Getting sheets//
let formSheet    = ss.getSheetByName('Форма');
let avtoSheet    = ss.getSheetByName('Авто');
let uchetSheet   = ss.getSheetByName("Учет Авто");
let workSheet    = ss.getSheetByName('Виды работ');
let vipiskaSheet = ss.getSheetByName('Выписка');
//Getting cells//
let nomerCell        = formSheet.getRange('A2');
let vinCell          = formSheet.getRange('A4');
let makerCell        = formSheet.getRange('A6');
let modelCell        = formSheet.getRange('A8');
let probegCell       = formSheet.getRange('A10');
let yearCell         = formSheet.getRange('A12');
let workVidCell      = formSheet.getRange('A14');
let workTypeCell     = formSheet.getRange('A16');
let ownerCell        = formSheet.getRange('A18');
let phoneCell        = formSheet.getRange('A20');
let summCell         = formSheet.getRange('A22');
let addToListCell    = formSheet.getRange('A24');
let sendResponseCell = formSheet.getRange('B1');

function startFunction(e){
  let cell = e.range;
  let isForma = cell.getSheet().getName() === formSheet.getSheetName();
  let isVipiska = cell.getSheet().getName() === vipiskaSheet.getSheetName();
  let isUchet = cell.getSheet().getName() === uchetSheet.getSheetName();
  let ui = SpreadsheetApp.getUi();
  if (isForma){
    if (cell.getA1Notation() === makerCell.getA1Notation()){ // валидация для модели
      workWithMakerCell();
    } else if (cell.getA1Notation() === nomerCell.getA1Notation()){ // отослать запрос
      sendRequest(cell);
    } else if (cell.getA1Notation() === addToListCell.getA1Notation()){ // перенести в список учета автом
      transferFromFormaToMainSheet(cell);
    } else if (cell.getA1Notation() === workVidCell.getA1Notation()){ // валидация для типов работ
      addValidationToWorkType();
    } else if (cell.getA1Notation() === vinCell.getA1Notation()){ // обработка VIN
      vinCellAnalize();
    }
  } else if (isVipiska){
    if (cell.getA1Notation() == 'B2'){ // выписка по номеру авто
      makeVipiska('nomer',cell);
    } else if (cell.getA1Notation() === 'A2'){ // выписка по владельцу
      makeVipiska('owner',cell);
    } else if (cell.getA1Notation() === 'C2'){ // выписка по VIN
      makeVipiska('vin',cell);
    }
  } else if(isUchet){
    if (cell.getRow() > 1){
      createValidationOnMainSheet(cell);
    }
  }
}

function workWithMakerCell(){
  createValidationForModel();
  let arr = ["Lamborghini","Koenigsegg","Bentley","Bugatti","Ferrari","Maserati"];
  if (arr.indexOf(makerCell.getValue()) !== -1){
    modelCell.setValue("Ой, не пизди");
  } else {
    modelCell.clearContent();
  }
  modelCell.activateAsCurrentCell();
}

function sendRequest(cellRange){
  let StandartAvtoNumberLength = 8;
  if (String(cellRange.getValue()).length === StandartAvtoNumberLength){
    sendResponseCell.check();
    analizeCarNumber();
    analizeVin();
    GetRespone();
  }
}

function transferFromFormaToMainSheet(unceckCell){
  if (unceckCell.isChecked()===true){
    clearFilter();
    clearSubTotalFormulas();
    addWorkTypeToList();
    addToList();
    unceckCell.uncheck();
  }
}

function addValidationToWorkType(){
  createValidationForWorkType(workVidCell.getValue(), workTypeCell);
  workTypeCell.clearContent();
  workTypeCell.activate();
}

function vinCellAnalize(){
  if (String(vinCell.getValue()).length === 17){
    analizeVin();
  }
}

function makeVipiska(forWho,cellRange){
  if (!cellRange.isBlank()){
    if (cellRange.getDataValidation().getCriteriaValues()[0].getValues().toString().split(',').indexOf(cellRange.getValue()) != -1){
      if (forWho === 'owner'){
        createVipiska('owner');
      } else if (forWho === 'nomer'){
        createVipiska('nomer');
      } else if (forWho === 'vin'){
        createVipiska('vin');
      }
    }
  }
}

function addToList(){
  let ui = SpreadsheetApp.getUi();
  let phone = String(phoneCell.getValue());
  let areMainCellHaveData;
  let i;
  clearFilter();
  clearSubTotalFormulas();
  if (phone.length === 0 || phone.length === 9){
    let checkArr = [makerCell,modelCell,nomerCell,vinCell];
    for (i = 0; i < checkArr.length; i++){
      if (!checkArr[i].isBlank()){
        areMainCellHaveData = true
      }
    }
    if (areMainCellHaveData){
      let columnsHeadersArray = uchetSheet.getRange(1, 1, 1, uchetSheet.getLastColumn()).getValues()[0];
      let id = uchetSheet.getRange(2,1).getValue() + 1;
      let dat = new Date();
      dat = dat.getDate() + '.' + (dat.getMonth()+1) + '.' + dat.getFullYear();
      let num = replaceEngSymb(String(nomerCell.getValue()).toUpperCase());
      let vin = String(vinCell.getValue()).toUpperCase();
      let searchArr = [/*записи*/, "Марка", "Модель", "Гос номер", "VIN", /Пробег*/, "Год", "Дата ТО", "Сумма", "Вид работы", "Тип работы", "Владелец", "Телефон"]
      let addArr = [id, makerCell.getValue(), modelCell.getValue(), num, vin, probegCell.getValue(), yearCell.getValue(), dat, summCell.getValue(),
                    workVidCell.getValue(), workTypeCell.getValue(), ownerCell.getValue(), phoneCell.getValue()]
      let whatAddArr = [[]];
      for (i = 0; i < searchArr.length; i++){
        let z = find(searchArr[i],columnsHeadersArray);
        whatAddArr[0][z] = addArr[i];
      }
      let x = uchetSheet.getLastRow() + 1;
      uchetSheet.getRange(x, 1, 1, searchArr.length).setValues(whatAddArr);
      for (i = 0; i < searchArr.length; i++){
        let y = find(searchArr[i], columnsHeadersArray);
        y++;
        if (i === 0 || i === 5 || i === 6){
          uchetSheet.getRange(x, y).setNumberFormat('0');
        } else if (i === 1 || i === 2 || i === 3 || i === 4 || i === 10  || i === 11 || i === 13){
          uchetSheet.getRange(x, y).setNumberFormat('@');
        } else if (i === 7){
          uchetSheet.getRange(x, y).setNumberFormat('dd mmmm yyyy');
        } else if (i === 8){
          uchetSheet.getRange(x, y).setNumberFormat('#,##0');
        } else if (i === 12){
          uchetSheet.getRange(x, y).setNumberFormat('"0"0');
        } else if (i === 9){
          let workVidCellOnMainTab = uchetSheet.getRange(x, y);
          workVidCellOnMainTab.setNumberFormat('@');
          workVidCellOnMainTab.clearDataValidations();
          workVidCellOnMainTab.setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .requireValueInRange(workSheet.getRange(1, 1, 1, workSheet.getLastColumn() ), true)
                                                 .build());
          createValidationForWorkType(uchetSheet.getRange(x, y).getValue(), uchetSheet.getRange(x, findColumnByName(searchArr[i + 1])));
        }
      }
      uchetSheet.getRange(x, 1, 1, 14).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      uchetSheet.getRange(2, 1, x, uchetSheet.getLastColumn()).sort({column: 1, ascending: false});
      phoneCell.setBackground(null);
      claerMainCells();
      nomerCell.offset(0, 1).clearContent();
      nomerCell.offset(2, 1).clearContent()

    } else {
      ui.alert("Указано слишком мало данных");
    }
  } else {
    ui.alert("Неправильный формат номера телефона");
    phoneCell.setBackground('#ff0000');
  }
}

function GetRespone(){
  let addictionalFieldsArray = [workVidCell,workTypeCell,ownerCell,phoneCell,summCell]
  if (sendResponseCell.isChecked() === true){
    if (String(nomerCell.getValue()).length === 8){
      clearFilter();
      clearSubTotalFormulas();
      let url = "https://baza-gai.com.ua/nomer/" + replaceEngSymb(String(nomerCell.getValue()).toUpperCase());
      let response = UrlFetchApp.fetch(url, {headers: {"Accept": "application/json"}, muteHttpExceptions: true});
      if (response.getResponseCode() === 200){
        let json = JSON.parse(response);
        makerCell.setBackground(null);
        makerCell.setValue(json.vendor);
        modelCell.setValue(firstLetterUpper(json.model))
        yearCell.setValue(json.year);
        createValidationForModel();
        addModelToList();
        analizeCarNumber();
      } else {
        makerCell.setValue("Не найдено в базе");
        makerCell.setBackground('#ff0000');
        modelCell.clearContent();
        yearCell.clearContent();
      }
    }
    for (let i = 0; i < addictionalFieldsArray.length; i++){
      addictionalFieldsArray[i].clearContent();
    }
    sendResponseCell.uncheck();
    createStyleForMainFromCell();
  }
}

function createValidationOnMainSheet(cellRange){
  let workVidColumn = findColumnByName('Вид работы');
  let workTypeColumn = findColumnByName('Тип работы');
  if (cellRange.getColumn() === workVidColumn){ // валидация на общем списке машин
    createValidationForWorkType(cellRange.getValue(), uchetSheet.getRange(cellRange.getRow(), workTypeColumn));
    uchetSheet.getRange(cellRange.getRow(), workTypeColumn).clearContent();
    uchetSheet.getRange(cellRange.getRow(), workTypeColumn).activate();
  }
}

function createVipiska(forWho){
  let whatFind;
  let arrayWithData = [];
  clearFilter();
  clearSubTotalFormulas();
  if (forWho === 'owner'){
    whatFind = vipiskaSheet.getRange(2, 1).getValue();
    arrayWithData = uchetSheet.getRange(2, findColumnByName('Владелец'),uchetSheet.getLastRow()).getValues();
    vipiskaSheet.getRange(2, 2, 1, 2).clearContent();
  } else if (forWho === 'nomer'){
    whatFind = vipiskaSheet.getRange(2, 2).getValue();
    arrayWithData = uchetSheet.getRange(2, findColumnByName('Гос номер'),uchetSheet.getLastRow()).getValues();
    vipiskaSheet.getRange(2, 1).clearContent();
    vipiskaSheet.getRange(2, 3).clearContent();
  } else if (forWho === 'vin'){
    whatFind = vipiskaSheet.getRange(2, 3).getValue();
    arrayWithData = uchetSheet.getRange(2, findColumnByName('VIN'),uchetSheet.getLastRow()).getValues();
    vipiskaSheet.getRange(2, 1, 1, 2).clearContent();
  }
  vipiskaSheet.getRange(6, 1, vipiskaSheet.getLastRow(), vipiskaSheet.getLastColumn()).clear();
  for (let i = 0 ; i < arrayWithData.length; i++){
    if (String(arrayWithData[i]) === whatFind){
      uchetSheet.getRange(i + 2, 1, 1, 14).copyTo(vipiskaSheet.getRange(vipiskaSheet.getLastRow() + 1, 1));
    }
  }
  createStyleForCell(vipiskaSheet.getRange(4, 1), 16, true, 'center', 'middle', false);
  vipiskaSheet.getRange(4, 1).setValue('Выписка для ' + whatFind);
  vipiskaSheet.getRange(5, 1, vipiskaSheet.getLastRow(), vipiskaSheet.getLastColumn()).clearDataValidations();
  addSumToVitag();
  autoFitAllColumns(vipiskaSheet);
}

function searchOilChange(){
  let phone;
  let owner;
  let maker;
  let model;
  let columnWithDate =     findColumnByName(/Дата*/);
  let columnForChangeOil = findColumnByName('Тип работы');
  let columnForWork =      findColumnByName('Вид работы');
  let columnOwner =        findColumnByName('Владелец');
  let columnMaker =        findColumnByName('Марка');
  let columnModel =        findColumnByName('Модель');
  let columnNomerAvto =    findColumnByName('Гос номер');
  let columnPhone =        findColumnByName('Телефон');
  let columnForCheckIfMailWasSent = 18;
  let arrayForCheckOil = uchetSheet.getRange(2, columnForChangeOil,uchetSheet.getLastRow()).getValues();
  let date = new Date();
  date.setFullYear(date.getFullYear() - 1);
  let str = 'Более года наза было заменено масло:<br>';
  for ( i = 0; i < arrayForCheckOil.length; i++){
    if (String(arrayForCheckOil[i]) === 'Замена масла' && uchetSheet.getRange(i + 2, columnForWork).getValue() === 'Двигатель'){
      if (new Date(uchetSheet.getRange(i + 2, columnWithDate).getValue()) < date && uchetSheet.getRange(i + 2, columnForCheckIfMailWasSent).isBlank()){
        if (uchetSheet.getRange(i + 2, columnOwner).isBlank()){
          owner = 'Владелец не указан.'
        } else {
          owner = uchetSheet.getRange(i + 2, columnOwner).getValue()
        };
        if (uchetSheet.getRange(i + 2, columnMaker).isBlank()){
          maker = 'Марка не указана.'
        } else {
          maker = uchetSheet.getRange(i + 2, columnMaker).getValue()
        };
        if (uchetSheet.getRange(i + 2, columnModel).isBlank()){
          model = 'Модель не указана.'
        } else {
          model = uchetSheet.getRange(i + 2, columnModel).getValue();
        };
        if (uchetSheet.getRange(i + 2, columnPhone).isBlank()){
          phone = '. Номер телефона не указан.'
        } else {
          phone = '. Номер телефона: <a href="tel:0' + uchetSheet.getRange(i + 2, columnPhone).getValue() + '">0'
                  + uchetSheet.getRange(i + 2, columnPhone).getValue() + '</a>';
        };
        str = str + '<br>' + owner + ' ' + maker + ' ' + model + phone;
        uchetSheet.hideColumns(columnForCheckIfMailWasSent);
        uchetSheet.getRange(i + 2, columnForCheckIfMailWasSent).setValue('sent');
        }
      }
    }
  Browser.msgBox(str);
  if (str !== 'Более года наза было заменено масло:<br>'){
    sendMail(str);
  }

}

function sendMail(message){
  let dt = new Date();
  dt = dt.getDate() + '.' + (dt.getMonth() + 1) + '.' + dt.getFullYear();
  MailApp.sendEmail({
    to: 'maximbirukov77@gmail.com',
    subject: 'Оповещение о замене масла на ' + dt,
    htmlBody: 'Добрый день! <br><br>' + message + '<br><br>' +
    'Набери им!'
  })
}
