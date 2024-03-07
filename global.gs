function getBlock(html, tag, idxStart) {
  let openTag = '<' + tag;
  let lenOpenTag = openTag.length;
  let closeTag = '</' + tag + '>';
  let lenCloseTag = closeTag.length;
  let countCloseTags = 0;
  let iMax = html.length;
  let idxEnd = 0;

  if (html.slice(idxStart, idxStart + lenOpenTag) != openTag) {
    idxStart = html.lastIndexOf(openTag, idxStart);
    if (idxStart == -1) return ["Can't to find openTag " + openTag + ' !', -1];
  };

  idxStart = html.indexOf('>', idxStart) + 1;
  let i = idxStart;
  
  while (i <= iMax) {
    i++;
    if (i === iMax) {
      return ['Could not find closing tag for ' + tag, -1];
    };
    let carrentValue = html[i];
    if (html[i] === '<'){
      let closingTag = html.slice(i, i + lenCloseTag);
      let openingTag = html.slice(i, i + lenOpenTag);
      if (html.slice(i, i + lenCloseTag) === closeTag) {
        if (countCloseTags === 0) {
          idxEnd = i - 1;
          break;
        } else {
          countCloseTags -= 1;
        };
      } else if (html.slice(i, i + lenOpenTag) === openTag) {
        countCloseTags += 1;
      };
    };
  };
  return [html.slice(idxStart,idxEnd + 1).trim(), idxEnd];
}

function getAttrName(html, attr, i) {
  let idxStart = html.indexOf(attr , i);
  if (idxStart == -1) return "Can't to find attr " + attr + ' !';
  idxStart = html.indexOf('"' , idxStart) + 1;
  let idxEnd = html.indexOf('"' , idxStart);
  return html.slice(idxStart,idxEnd).trim();
}

function getOpenTag(html, tag, idxStart) {
  let openTag = '<' + tag;
  let lenOpenTag = openTag.length;

  if (html.slice(idxStart, idxStart + lenOpenTag) != openTag) {
    idxStart = html.lastIndexOf(openTag, idxStart);
    if (idxStart == -1) return "Can't to find openTag " + openTag + ' !';
  };

  let idxEnd = html.indexOf('>', idxStart) + 1;
  if (idxStart == -1) return "Can't to find closing bracket '>' for openTag!";
  return html.slice(idxStart,idxEnd).trim();
}

function getAccountData(acc) {
  let accSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List');
  let accRows = accSheet.getDataRange().getValues();
  for(let i = 1; i < accRows.length; i++) {
    if(accRows[i][4] == acc) {
      let accData = {};
      accData.geo = accRows[i][0];
      accData.language = accRows[i][1];
      accData.currency = accRows[i][2];
      accData.name = accRows[i][6];
      if(accRows[i][5] != '') accData.phone = accRows[i][5];

      return accData;
    }
  }
}

function getCategoryId(catName) {
  let catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  let catRows = catSheet.getDataRange().getValues();
  for(let i = 1; i < catRows.length; i++) {
    if(catRows[i][1] == catName) {
      return catRows[i][0];
    }
  }
}

var alert = function(text) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(text);
}

var prompt = function(text) {
  var ui = SpreadsheetApp.getUi();
  ui.prompt(text);
}

var toast = function(title, text) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast(text, title, 5);
}