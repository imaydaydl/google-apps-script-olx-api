var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advert');

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Управління')
    .addItem('Оновити товари', 'updateProducts')
    .addItem('Відправити в OLX обрані', 'preSendToOlx')
    .addItem('Оновити в OLX обрані', 'preUpdateOnOlx')
    .addItem('Видалити обрані (модалка)', 'deleteProductsModal')
    .addItem('Дублювати обрані', 'doubleProducts')
    .addItem('Оновити статистику', 'preUpdateStats')
    .addItem('Оновити повідомлення', 'preUpdateMessages')
    .addItem('Відповісти на обраний (модалка)', 'preSendChosenMessage')
    .addItem('Перекласти обрані', 'translate')
    .addItem('Конвертація валюти в обраних (модалка)', 'currencyModal')
    .addItem('Оновити список категорій', 'preAskCategories')
    .addItem('Оновити список міст', 'preGetCityList')
    .addToUi();

  createSheets();
  insertCatSelector();
  insertCitySelector();
}

function onEdit(e) {
  let range = e.range;
  
  if(range.getSheet().getName() == 'Advert') {
    if(range.getColumn() == 32) {
      let cid = 0;
      let clo = 0;
      let cla = 0;
      let citySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cities');
      let cities = citySheet.getDataRange().getValues();
      for(let ci = 1; ci < cities.length; ci++) {
        if(cities[ci][2] == range.getValue()) {
          cid = cities[ci][1];
          clo = cities[ci][4];
          cla = cities[ci][3];
        }
      }
      let id = range.offset(0, 1);
      let latitude = range.offset(0, 2);
      let longitude = range.offset(0, 3);
      id.setValue(cid);
      latitude.setValue(cla);
      longitude.setValue(clo);
    }
  }
}

function preSendToOlx() {
  initToken('sendToOlx');
}

function preUpdateOnOlx() {
  initToken('updateOnOlx');
}

function preUpdateStats() {
  initToken('updateStats');
}

function preUpdateMessages() {
  initToken('updateMessages');
}

function preSendChosenMessage() {
  initToken('sendChosenMessage');
}

function preAskCategories() {
  initToken('askCategories');
}

function preGetCityList() {
  initToken('getCityList');
}

function initToken(method) {
  let accountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
  let accounts = accountSheet.getDataRange().getValues();

  for(let ac = 1; ac < accounts.length; ac++) {
    if(accounts[ac][3] != '') {
      if(['sendToOlx', 'updateOnOlx', 'deleteProductsOlx', 'updateStats', 'updateMessages', 'sendChosenMessage'].includes(method)) {
        if(olx_token[accounts[ac][0].toLowerCase()] == '') {
          getAuthorizationCode(accounts[ac][0].toLowerCase());
        }
      } else {
        if(client_token[accounts[ac][0].toLowerCase()] == '') {
          clientCredentials(accounts[ac][0].toLowerCase());
        }
      }

      switch(method) {
        case 'sendToOlx':
          sendToOlx(accounts[ac][0].toLowerCase());
          break;
        case 'updateOnOlx':
          updateOnOlx(accounts[ac][0].toLowerCase());
          break;
        case 'deleteProductsOlx':
          deleteProductsOlx(accounts[ac][0].toLowerCase());
          break;
        case 'updateStats':
          updateStats(accounts[ac][0].toLowerCase());
          break;
        case 'updateMessages':
          updateMessages(accounts[ac][0].toLowerCase());
          break;
        case 'sendChosenMessage':
          sendChosenMessage(accounts[ac][0].toLowerCase());
          break;
        case 'askCategories':
          askCategories(accounts[ac][0].toLowerCase());
          break;
        case 'getCityList':
          getCityList(accounts[ac][0].toLowerCase());
          break;
      }
    }
  }
}

function getAuthorizationCode(c = 'pt') {
  let tokenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tokens');
  let tokenRows = tokenSheet.getDataRange().getValues();

  if(tokenRows.length == 1) {
    callModal(c);
  } else if(getTokenDate() && new Date(getTokenDate().date) <= new Date()) {
    setExpired();
    callModal(c);
  } else if(getTokenDate() && new Date(getTokenDate().date) > new Date()) {
    authorizationCode(getTokenDate().token, c);
  }
}

function getTokenDate() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("Tokens");
  let lastRow = sheet.getLastRow();
  
  for (let i = lastRow; i > 0; i--) {
    let value = sheet.getRange(i, 4).getValue();
    if (value === 'active') {
      let lastRowData = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues();
      return {'token': lastRowData[0][2], 'date': lastRowData[0][4]};
    }
  }
  
  return null;
}

function setExpired() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("Tokens");
  let lastRow = sheet.getLastRow();
  
  for (let i = lastRow; i > 0; i--) {
    let value = sheet.getRange(i, 4).getValue();
    if (value === 'active') {
      sheet.getRange(i, 4).setValue('expired');
    }
  }
}

function callModal(c) {
  let template = HtmlService.createTemplateFromFile('authorization_code');
  template.c = c;

  let html = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Отримання коду');
}

function getAuthorizeToken(code, c) {
  let tokenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tokens');
  let tokensCount = tokenSheet.getDataRange().getValues().length;
  let nextRow = tokensCount + 1;
  tokenSheet.getRange(nextRow, 1).setValue(tokensCount);
  tokenSheet.getRange(nextRow, 2).setValue(c.toUpperCase());
  tokenSheet.getRange(nextRow, 3).setValue(code);
  tokenSheet.getRange(nextRow, 4).setValue('active');

  let ac = authorizationCode(code, c);

  let expire = ac.expires_in;
  let date = new Date();
  date.setSeconds(date.getSeconds() + expire);
  tokenSheet.getRange(nextRow, 5).setValue(date.getFullYear() + '-' + date.getMonth() + '-' + date.getDate() + ' ' + date.getHours() + ':' + date.getMinutes() + ':' + date.getSeconds());
}

function updateProducts() {
  // Тут можна вписати код для автоматичного збирання об'єкту із даними та використати fillTable() для заповнення таблиці Advert
  // Приклад даних:
  // {
  //    category: '',
  //    source: '', 
  //    account: '',
  //    products: 
  //        [ {
  //            id: '', - внутрішній айді товара (SKU)
  //            title: '',- назва товара
  //            price: '', - ціна за товар
  //            description: '', - опис товара
  //            image1: '', - url на зображення
  //            image2: '', - url на зображення
  //            image3: '', - url на зображення
  //            image4: '', - url на зображення
  //            image5: '', - url на зображення
  //            image6: '', - url на зображення
  //            image7: '', - url на зображення
  //            image8: '', - url на зображення
  //            geo: 'PT', - ISO code країни, необхідний, щоб код зрозумів на який OLX відправляти (PT - Португалія)
  //            language: 'pt', - мова, зазвичай такий самий як і geo
  //            url: '', - внутрішній url на товар
  //            product_status: '', - new або used
  //            currency: 'EUR', - валюта по товару
  //            acc_name: '', - ім'я користувача, який буде зазначений як контактна особа по товару в OLX
  //            acc_phone: '' - номер телефону цього користувача (необов'язково вказувати)
  //        } ]
  //}

  // if(adverts) {
    // fillTable(adverts);
  // }
}

function fillTable(list) {
  let productCount = sheet.getDataRange().getValues().length;
  for(let i = 0; i < list.products.length; i++) {
    let p = list.products[i];
    productCount = productCount + 1;
    sheet.getRange(productCount, 1).setValue(p.geo);
    sheet.getRange(productCount, 2).setValue(list.account);
    sheet.getRange(productCount, 3).setValue(p.language);
    sheet.getRange(productCount, 4).setValue(p.currency);
    sheet.getRange(productCount, 5).setValue(p.id);
    sheet.getRange(productCount, 6).setValue(p.url);
    sheet.getRange(productCount, 7).setValue(p.title);
    sheet.getRange(productCount, 8).setValue(list.category);
    if(p.image1) sheet.getRange(productCount, 10).setValue(p.image1);
    if(p.image2) sheet.getRange(productCount, 11).setValue(p.image2);
    if(p.image3) sheet.getRange(productCount, 12).setValue(p.image3);
    if(p.image4) sheet.getRange(productCount, 13).setValue(p.image4);
    if(p.image5) sheet.getRange(productCount, 14).setValue(p.image5);
    if(p.image6) sheet.getRange(productCount, 15).setValue(p.image6);
    if(p.image7) sheet.getRange(productCount, 16).setValue(p.image7);
    if(p.image8) sheet.getRange(productCount, 17).setValue(p.image8);
    sheet.getRange(productCount, 18).setValue(p.description);
    sheet.getRange(productCount, 19).insertCheckboxes().setValue('true');
    sheet.getRange(productCount, 20).setValue(p.price);
    sheet.getRange(productCount, 21).insertCheckboxes();
    sheet.getRange(productCount, 22).insertCheckboxes();
    sheet.getRange(productCount, 23).insertCheckboxes();
    sheet.getRange(productCount, 27).insertCheckboxes().setValue('true');
    sheet.getRange(productCount, 28).insertCheckboxes();
    sheet.getRange(productCount, 29).insertCheckboxes().setValue('true');
    sheet.getRange(productCount, 30).insertCheckboxes();
    sheet.getRange(productCount, 31).insertCheckboxes();
    sheet.getRange(productCount, 37).setValue(list.account);
    if(p.acc_phone) sheet.getRange(productCount, 38).setValue(p.acc_phone);
  }
}

function doubleProducts() {
  let selectedRanges = sheet.getActiveRangeList();
  let selectedRows = selectedRanges.getRanges();

  if(selectedRows.length == 0) {
    alert('Оберіть хоча б один рядок із товаром');
  } else {
    let advert = [];
    for(let m = 0; m < selectedRows.length; m++) {
      let ad = selectedRows[m].getValues();
      if(ad[0] === undefined) {
        alert('Оберіть хоча б один рядок із товаром');
      } else if(selectedRows[m].getRowIndex() == 1) {
        alert('Перший рядок не можна дублювати');
      } else if(ad[0].length > 0) {
        ad[0][23] = '';
        ad[0][24] = '';
        ad[0][25] = '';
        ad[0][39] = '';
        ad[0][40] = '';
        ad[0][41] = '';
        ad[0][42] = '';
        ad[0][43] = '';

        advert.push(ad[0]);
      }
    }

    let advertSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advert');
    let newRow = advertSheet.getLastRow() + 1;
    let firstColumn = advertSheet.getRange(newRow, 1).getA1Notation();
    let allNewRow = advertSheet.getLastRow() + advert.length;
    let lastColumn = advertSheet.getRange(allNewRow, advertSheet.getLastColumn()).getA1Notation();

    let range = advertSheet.getRange(firstColumn + ":" + lastColumn);
    range.setValues(advert);

    checkboxCols = [19, 21, 22, 23, 27, 28, 29, 30, 31];
    for(let l = 0; l < checkboxCols.length; l++) {
      let columnToInsertCheckboxes = checkboxCols[l];

      let dataRange = sheet.getRange(2, columnToInsertCheckboxes, sheet.getLastRow() - 1, 1);
      dataRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }

    insertCatSelector();
    insertCitySelector();
  }
}

function getCityList(c = 'pt') {
  let citySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cities');
  let countCities = citySheet.getDataRange().getValues().length;

  if(countCities > 1) {
    let lastInRange = citySheet.getRange(citySheet.getLastRow(), citySheet.getLastColumn()).getA1Notation();
    let rangeToClear = citySheet.getRange("A2:" + lastInRange);
    rangeToClear.clear({contentsOnly: true});
  }

  while(true) {
    let off = {'offset': countCities > 1 ? countCities : 0};
    let cities = getCities(off, c).data;

    for(let d = 0; d < cities.length; d++) {
      countCities++;
      citySheet.getRange(countCities, 1).setValue(c.toUpperCase());
      citySheet.getRange(countCities, 2).setValue(cities[d]['id']);
      citySheet.getRange(countCities, 3).setValue(cities[d]['name']);
      citySheet.getRange(countCities, 4).setValue(cities[d]['latitude']);
      citySheet.getRange(countCities, 5).setValue(cities[d]['longitude']);
      citySheet.getRange(countCities, 6).setValue(cities[d]['municipality']);
    }

    if(cities.length < 1000) {
      break;
    }
  }

  toast('Успішно!', 'Список міст успішно оновлено');
}

function getRangeList(searchValue) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cities');

  let textFinder = spreadsheet.createTextFinder(searchValue.toUpperCase());
  let ranges = textFinder.findAll();

  return [ranges[0].getA1Notation().replace('A', 'C'), ranges[ranges.length - 1].getA1Notation().replace('A', 'C')];
}

function sendToOlx(c = 'pt') {
  if(sheet.getSheetName() == 'Advert') {
    let selectedRanges = sheet.getActiveRangeList();
    let selectedRows = selectedRanges.getRanges();

    selectedRows.forEach(function(selectedRange) {
      let advert = selectedRange.getValues();
      if(advert[0] !== undefined && advert[0].length > 0 && advert[0][0] != '' && advert[0][0] == c.toUpperCase()) {
        let json = {};
        if(advert[0][23] ==  '' && advert[0][4] != '') {
          json.title = advert[0][6];
          json.description = advert[0][17];
          json.advertiser_type = advert[0][26] ? 'private': 'business';
          json.external_id = advert[0][4];
          json.category_id = advert[0][8] ? getCategoryId(advert[0][8]) : 1407;
          
          json.contact = {};
          let contactData = getAccountData(advert[0][1]);
          json.contact.name = contactData.name;
          if(contactData.phone) json.contact.phone = contactData.phone;

          json.location = {"city_id": advert[0][32], "latitude": advert[0][33], "longitude": advert[0][34]};

          json.price = {"value": advert[0][19], "currency": advert[0][3], "negotiable": advert[0][20], "trade": advert[0][22]};

          let state = advert[0][28] ? 'new' : 'used';
          json.attributes = [{"code": "state", "value": state}, {"code": "model", "value": "cts"}];

          let images = [];
          if(advert[0][9] != '') images.push({"url": advert[0][9]});
          if(advert[0][10] != '') images.push({"url": advert[0][10]});
          if(advert[0][11] != '') images.push({"url": advert[0][11]});
          if(advert[0][12] != '') images.push({"url": advert[0][12]});
          if(advert[0][13] != '') images.push({"url": advert[0][13]});
          if(advert[0][14] != '') images.push({"url": advert[0][14]});
          if(advert[0][15] != '') images.push({"url": advert[0][15]});
          if(advert[0][16] != '') images.push({"url": advert[0][16]});
          if(images.length > 0) json.images = images;

          let new_advert = createAdvert(json, c);
          if(new_advert.data) {
            advert[0][23] = new_advert.data.id;
            advert[0][24] = new_advert.data.url;
            advert[0][25] = new_advert.data.status;
            advert[0][40] = new_advert.data.created_at;
            advert[0][41] = new_advert.data.valid_to;

            selectedRange.setValues(advert);

            toast('Успішно!', 'Рядок ' + selectedRange.getRow() + ' успішно передано в OLX. ID: ' + new_advert.data.id);
          }
        }
      }
    });
  }
}

function updateOnOlx() {
  if(sheet.getSheetName() == 'Advert') {
    let selectedRanges = sheet.getActiveRangeList();
    let selectedRows = selectedRanges.getRanges();

    selectedRows.forEach(function(selectedRange) {
      let advert = selectedRange.getValues();
      if(advert[0] !== undefined && advert[0].length > 0 && advert[0][0] != '' && advert[0][23] != '') {
        let json = {};
        if(advert[0][23] ==  '' && advert[0][4] != '') {
          json.title = advert[0][6];
          json.description = advert[0][17];
          json.advertiser_type = advert[0][26] ? 'private': 'business';
          json.external_id = advert[0][4];
          json.category_id = advert[0][8] ? getCategoryId(advert[0][8]) : 1407;
          
          json.contact = {};
          let contactData = getAccountData(advert[0][1]);
          json.contact.name = contactData.name;
          if(contactData.phone) json.contact.phone = contactData.phone;

          json.location = {"city_id": advert[0][32], "latitude": advert[0][33], "longitude": advert[0][34]};

          json.price = {"value": advert[0][19], "currency": advert[0][3], "negotiable": advert[0][20], "trade": advert[0][22]};

          let state = advert[0][28] ? 'new' : 'used';
          json.attributes = [{"code": "state", "value": state}, {"code": "model", "value": "cts"}];

          let images = [];
          if(advert[0][9] != '') images.push({"url": advert[0][9]});
          if(advert[0][10] != '') images.push({"url": advert[0][10]});
          if(advert[0][11] != '') images.push({"url": advert[0][11]});
          if(advert[0][12] != '') images.push({"url": advert[0][12]});
          if(advert[0][13] != '') images.push({"url": advert[0][13]});
          if(advert[0][14] != '') images.push({"url": advert[0][14]});
          if(advert[0][15] != '') images.push({"url": advert[0][15]});
          if(advert[0][16] != '') images.push({"url": advert[0][16]});
          if(images.length > 0) json.images = images;

          updateAdvert(advert[0][23], json, advert[0][0].toLowerCase());

          toast('Успішно!', 'Рядок ' + selectedRange.getRow() + ' успішно оновлено в OLX.');
        }
      }
    });
  }
}

function getProductIds() {
  let rows = sheet.getDataRange().getValues();
  let ids = [];
  let productCount = rows.length;
  if(productCount > 1) {
    for (let i = 1; i < productCount; i++) {
      ids.push(rows[i][4]);
    }
  }

  return ids;
}

function askCategories(c = 'pt') {
  let sheetCategories = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');

  if(sheetCategories.getDataRange().getValues().length > 1) {
    let lastInRange = sheetCategories.getRange(sheetCategories.getLastRow(), sheetCategories.getLastColumn()).getA1Notation();
    let rangeToClear = sheetCategories.getRange("A2:" + lastInRange);
    rangeToClear.clear({contentsOnly: true});
  }

  let cats = [];

  while(true) {
    cats = getCategories(c);

    if(cats.data.length > 0){
      let r = sheetCategories.getDataRange().getValues().length + 1;
      for(let i = 0; i < cats.data.length; i++) {
        sheetCategories.getRange(r, 1).setValue(cats.data[i]['id']);
        sheetCategories.getRange(r, 2).setValue(cats.data[i]['name']);
        sheetCategories.getRange(r, 3).setValue(cats.data[i]['parent_id']);
        sheetCategories.getRange(r, 4).setValue(cats.data[i]['photos_limit']);
        sheetCategories.getRange(r, 5).setValue(cats.data[i]['is_leaf']);
        sheetCategories.getRange(r, 6).setValue(c.toUpperCase());
        r++;
      }
    }

    if(cats.data.length < 1000) {
      break;
    }
  }
}

function deleteProductsModal() {
  if(sheet.getSheetName() == 'Advert') {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('delete_modal');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Видалення обраних товарів');
  }
}

function deleteProducts(selectedValue) {
  switch(selectedValue) {
    case 'document':
      deleteProductsDocument();
      break;
    case 'olx':
      initToken('deleteProductsOlx');
      break;
    case 'olx_plus_document':
      initToken('deleteProductsOlx');
      deleteProductsDocument();
      break;
  }
}

function deleteProductsDocument() {
  let selectedRanges = sheet.getActiveRangeList();
  let selectedRows = selectedRanges.getRanges();

  let r = 0;
  selectedRows.forEach(function(selectedRange) {
    let startRow = selectedRange.getRow() - r;

    sheet.deleteRow(startRow);
    r++;
  });
}

function deleteProductsOlx(c = 'pt') {
  if(sheet.getSheetName() == 'Advert') {
    let selectedRanges = sheet.getActiveRangeList();
    let selectedRows = selectedRanges.getRanges();

    let ids = [];
    selectedRows.forEach(function(selectedRange) {
      ids.push(selectedRange.getValues()[0][23]);
    });

    for(let k in ids) {
      deleteAdvert(ids[k], c);
    }

    toast('Успішно', 'Обрані рядки успішно видалені з OLX');
  }
}

function insertCitySelector() {
  let sheetCategories = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cities');

  let sheetRows = sheet.getDataRange().getValues();
  let geos = [];
  for(let a = 1; a < sheetRows.length; a++){
    if(!geos.includes(sheetRows[a][0].toLowerCase())) geos.push(sheetRows[a][0].toLowerCase());
  }

  geos.forEach(function(geo) {
    let fromto = getRangeList(geo);

    let range = sheetCategories.getRange(fromto[0]+':'+fromto[1]);
  
    if(range.getValues().length > 0) {
      let dropdown = SpreadsheetApp.newDataValidation()
        .requireValueInRange(range)
        .setAllowInvalid(false)
        .build();
      
      for(let a = 2; a <= sheetRows.length; a++){
        if(sheet.getRange(a, 1).getValue() == geo.toUpperCase()) {
          sheet.getRange(a, 32).setDataValidation(dropdown);
          sheet.getRange(a, 32).setWrap(true);
          sheet.getRange(a, 32).setHorizontalAlignment("center");
        }
      }
    }
  });
}

function insertCatSelector() {
  let sheetCategories = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');

  let lastInRange = 'B' + sheetCategories.getDataRange().getLastRow();

  let range = sheetCategories.getRange('B2:'+lastInRange);
  
  if(range.getValues().length > 0) {
    let dropdown = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range)
      .setAllowInvalid(false)
      .build();
    
    let sheetRows = sheet.getDataRange().getValues();

    for(let a = 2; a <= sheetRows.length; a++) {
      sheet.getRange(a, 9).setDataValidation(dropdown);
      sheet.getRange(a, 9).setWrap(true);
      sheet.getRange(a, 9).setHorizontalAlignment("center");
    }
  }
}

function updateStats(c = 'pt') {
  let advertRows = sheet.getDataRange().getValues();
  for(let v = 2; v <= advertRows.length; v++){
    if(sheet.getRange(v, 24).getValue() != '' && sheet.getRange(v, 1).getValue() == c.toUpperCase()) {
      let stat = getAdvertStats(sheet.getRange(v, 24).getValue(), c);
      sheet.getRange(v, 42).setValue(stat.data.advert_views);
      sheet.getRange(v, 43).setValue(stat.data.phone_views);
      sheet.getRange(v, 44).setValue(stat.data.users_observing);
    }
  }
}

function updateMessages(c = 'pt') {
  let threadsQ = getThreads(c);
  let threads = threadsQ.data;

  let messageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Message');
  let messageRow = messageSheet.getDataRange().getValues().length;;
  let messageIds = getMessageIds();
  let clientIds = getClientIds();
  if(threads.length > 0) {
    for(let t = 0; t < threads.length; t++) {
      let messages = getMessages(threads[t].id, c);
      if(messages.length > 0) {
        let new_messages = 0;
        for(let m = 0; m < messages.length; m++) {
          if(!messageIds.includes(messages[m].id)) {
            messageRow = messageRow+1;
            if(!clientIds.includes(threads[t].interlocutor_id)) {
              let clientData = getUser(threads[t].interlocutor_id, c);
              let cl = addNewClient(clientData, c);
              messageSheet.getRange(messageRow, 6).setValue(cl.name);
            }

            messageSheet.getRange(messageRow, 1).setValue(threads[t].advert_id);
            messageSheet.getRange(messageRow, 4).setValue(messages[m].thread_id);
            messageSheet.getRange(messageRow, 5).setValue(threads[t].interlocutor_id);
            messageSheet.getRange(messageRow, 7).setValue(messages[m].text);
            messageSheet.getRange(messageRow, 8).setValue(messages[m].id);
            messageSheet.getRange(messageRow, 9).setValue(messages[m].created_at);
            messageSheet.getRange(messageRow, 10).setValue(c.toUpperCase());

            new_messages++;
          }
        }
        if(new_messages > 0) {
          toast('Повідомлення', 'Завантажено ' + new_messages + ' нових повідомлень');
        } else {
          toast('Повідомлення', 'Немає нових повідомлень');
        }
      } else {
        toast('Повідомлення', 'Ще немає повідомлень');
      }
    }
  } else {
    toast('Повідомлення', 'Ще немає тем');
  }
}

function getMessageIds() {
  let messageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Message');
  let rows = messageSheet.getDataRange().getValues();
  let messageIds = [];
  if(rows.length > 1) {
    for (let i = 1; i < rows.length; i++) {
      messageIds.push(rows[i][7]);
    }
  }

  return messageIds;
}

function getClientIds() {
  let clientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  let rows = clientsSheet.getDataRange().getValues();
  let clientIds = [];
  if(rows.length > 1) {
    for (let i = 1; i < rows.length; i++) {
      clientIds.push(rows[i][2]);
    }
  }

  return clientIds;
}

function addNewClient(data, geo) {
  let clientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  let rows = clientsSheet.getDataRange().getValues();
  let clientlength = rows.length;
  for (let i = 1; i < rows.length; i++) {
    clientlength = clientlength + 1;
    clientsSheet.getRange(clientlength, 1).setValue(geo);
    clientsSheet.getRange(clientlength, 2).setValue(data.id);
    clientsSheet.getRange(clientlength, 3).setValue(data.name);
  }
}

function sendChosenMessage(c = 'pt') {
  let messageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Message');
  let selectedRanges = messageSheet.getActiveRangeList();
  let selectedRows = selectedRanges.getRanges();

  if(selectedRows.length > 1) {
    alert('Оберіть лише 1 рядок, на який Ви хочете відповісти');
  } else if(selectedRows.length < 1) {
    alert('Оберіть хоча б 1 рядок, на який Ви хочете відповісти з вкладки Message');
  } else {
    let selectedRow = selectedRows[0];
    let row = selectedRow.getRowIndex();
    let sr = selectedRow.getValues();
    let template = HtmlService.createTemplateFromFile('send_message');
    template.c = c;
    template.thread = sr[0][3];
    template.row = row;

    let html = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Відправка повідомлення');
  }
}

function sendTypedMessage(thread, message, row, c) {
  let messageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Message');

  messageSheet.getRange(row, 11).setValue(message);

  sendMessages(thread, message, c);
}

function translate() {
  if(sheet.getSheetName() == 'Advert') {
    let selectedRanges = sheet.getActiveRangeList();
    let selectedRows = selectedRanges.getRanges();

    selectedRows.forEach(function(selectedRange) {
      let r = selectedRange.getValues();
      let lang = r[0][0].toLowerCase();
      let title = r[0][6];
      let text = r[0][17];

      r[0][6] = LanguageApp.translate(title, '', lang);
      r[0][17] = LanguageApp.translate(text, '', lang);

      selectedRange.setValues(r);
    });
  }
}

function currencyModal() {
  if(sheet.getSheetName() == 'Advert') {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('exchange_value')
      .setHeight(250);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Конвертація суми відповідно до валюти');
  }
}

function exchangeValue(currencies, markup) {
  if(sheet.getSheetName() == 'Advert') {
    let selectedRanges = sheet.getActiveRangeList();
    let selectedRows = selectedRanges.getRanges();

    let data = getRates(currencies.to);
    let exchangeRate = data.rates[currencies.from];

    selectedRows.forEach(function(selectedRange) {
      let r = selectedRange.getValues();
      let sum = r[0][19];

      let resultExchange = sum / exchangeRate;

      if(markup != '' && markup != 0) resultExchange = resultExchange + resultExchange * (markup / 100);
      r[0][19] = resultExchange.toFixed(2).replace('.', ',');

      selectedRange.setValues(r);
    });
  }
}