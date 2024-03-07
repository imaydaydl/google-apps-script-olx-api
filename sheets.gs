var sheets = ['Advert', 'Message', 'Clients', 'List', 'Categories', 'Cities', 'Tokens'];

function createSheets() {
  for(let s = 0; s < sheets.length; s++) {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sh = spreadsheet.getSheetByName(sheets[s]);
    
    if (sh == null) {
      spreadsheet.insertSheet(sheets[s]);
      insertDefaults(sheets[s]);
    }
  }
}

function insertDefaults(sh) {
  let columns = [];
  switch(sh) {
    case 'Advert':
      columns = [['Geo',	'Account OLX',	'Language',	'Currency',	'ID/SKU',	'Url', 'Name', 'Parsed category', 'Category',	'Photo1',	'Photo2',	'Photo3',	'Photo4',	'Photo5',	'Photo6',	'Photo7',	'Photo8',	'Description',	'Price',	'Is price',	'Is Negotiable',	'Is Free',	'Is Exchange',	'OLX ID',	'OLX URL', 'OLX status', 'Is Private', 'Is Business',	'Is New',	'Is Used',	'Autorenewal',	'City',	'City ID',	'Latitude',	'Longitude'], ['Seller email',	'Seller phone'], ['Date create', 'Date renewal', 'Advert views', 'Phone views', 'Users observing']];
      break;
    case 'Message':
      columns = [['Advert Id',	'Advert Url',	'Name', 'Advert',	'Thread Id',	'Client Id',	'Client name',	'Message',	'Message Id',	'Message date',	'Geo',	'Answer']];
      break;
    case 'Clients':
      columns = [['Geo',	'Client Id',	'Name',	'City',	'Street',	'House',	'Flat',	'Department',	'Comment',	'Phone',	'Email']];
      break;
    case 'List':
      columns = [['Geo',	'Language',	'Currency',	'Account OLX',	'Email',	'Phone',	'Name']];
      break;
    case 'Categories':
      columns = [['ID',	'Name',	'Parent Id',	'Photos Limit',	'Is Leaf',	'GEO']];
      break;
    case 'Cities':
      columns = [['Country	ID',	'City Name',	'Latitude',	'Longitude', 'Municipality']];
      break;
    case 'Tokens':
      columns = [['ID',	'GEO',	'Code',	'Status',	'Active to date']];
      break;
  }

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newsheet = spreadsheet.getSheetByName(sh);
  let firstRowRange = newsheet.getRange("1:1");
  firstRowRange.setFontWeight("bold");

  let colNum = 1;
  for(let b = 0; b < columns.length; b++) {
    for(let c = 0; c < columns[b].length; c++) {
      newsheet.getRange(1, colNum).setValue(columns[b][c]);
      colNum++;
    }
    colNum++;
  }
}