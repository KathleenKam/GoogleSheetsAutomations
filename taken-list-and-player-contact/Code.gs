var RANGE;
var CHARA_SHEET_NAME = "Characters_" + RANGE;
var KEY_SHEET_NAME = "Key_" + RANGE;
var TEMPLATE_NAME = "index_" + RANGE;

var NAME_INDEX = 1;
var PLAYER_INDEX = 3;
var SORT_INDEX = 0;
var ALIGNMENT_INDEX = 4;

// for HTML escaping

var ESC_MAP = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
};

function escapeHTML(s, forAttribute) {
    return s.replace(forAttribute ? /[&<>'"]/g : /[&<>]/g, function(c) {
        return ESC_MAP[c];
    });
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automations')
      .addItem('Generate Taken A-M', 'generateTakenAM')
      .addItem('Generate Taken N-Z', 'generateTakenNZ')
      .addItem('Generate Player Contact', 'generatePlayerContact')
      .addToUi();
}

function generatePlayerContact() {
  var output = HtmlService
     .createTemplateFromFile("players")
     .evaluate();
  
  var codeOutput = output.getContent();
  var ui = SpreadsheetApp.getUi();
  
  var app = UiApp.createApplication().setWidth(400).setHeight(400);
  var html = app.createHTML();
  
  html.setText(codeOutput);
  app.add(html);
  ui.showModelessDialog(app,'C/P this thing (click on the text, then Ctrl + A, Ctrl + C)');
}

function getPlayerData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sortCriteria = ss.getSheetByName("Players").getDataRange().getValues();
  
  var rawData = [];
  
  for (var i = 1; i < sortCriteria.length; i++) {
    var obj = {};
    obj.sortCriteria = sortCriteria[i][0];
    obj.playerData = sortCriteria[i];
    obj.characterData = [];
    rawData.push(obj);
  }
  
  sortCharacterDataByPlayer(ss, rawData, "Characters_A_M");
  sortCharacterDataByPlayer(ss, rawData, "Characters_N_Z");
  return rawData;
}

function sortCharacterDataByPlayer(ss, rawData, ssName) {
  var ssData = ss.getSheetByName(ssName).getDataRange().getValues();
  
  var tempData;
  var colorHex;
  
  for (var i = 0; i < ssData.length; i++) {
    for (var j = 0; j < rawData.length; j++) {
      if (rawData[j].sortCriteria === ssData[i][PLAYER_INDEX]) {
        tempData = ssData[i];
        rawData[j].characterData.push(tempData);
        break;
      }
    }
  }


}

function generateTakenAM() {
  RANGE = "A_M";
  CHARA_SHEET_NAME = "Characters_" + RANGE;
  KEY_SHEET_NAME = "Key_" + RANGE;
  TEMPLATE_NAME = "index_" + RANGE;
  generateTaken();
}

function generateTakenNZ() {
  RANGE = "N_Z";
  CHARA_SHEET_NAME = "Characters_" + RANGE;
  KEY_SHEET_NAME = "Key_" + RANGE;
  TEMPLATE_NAME = "index_" + RANGE;
  generateTaken();
}

function generateTaken() {
  var output = HtmlService
     .createTemplateFromFile(TEMPLATE_NAME)
     .evaluate();
  
  var codeOutput = output.getContent();
  var ui = SpreadsheetApp.getUi();
  
  var app = UiApp.createApplication().setWidth(400).setHeight(400);
  var html = app.createHTML();
  
  html.setText(codeOutput);
  app.add(html);
  ui.showModelessDialog(app,'C/P this thing (click on the text, then Ctrl + A, Ctrl + C)');
}

function getData() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sortCriteria = ss.getSheetByName(KEY_SHEET_NAME).getRange("A:A").getValues();
  
  var rawData = [];
  
  for (var i = 0; i < sortCriteria.length; i++) {
    var obj = {};
    obj.sortCriteria = sortCriteria[i][0];
    obj.data = [];
    rawData.push(obj);
  }
  
  var ssData = ss.getSheetByName(CHARA_SHEET_NAME).getDataRange().getValues();
  
  var tempData;
  var colorHex;
  
  for (var i = 0; i < ssData.length; i++) {
    for (var j = 0; j < rawData.length; j++) {
      if (rawData[j].sortCriteria === ssData[i][SORT_INDEX]) {
        tempData = ssData[i];
        // alignment colors
        
        if (tempData[ALIGNMENT_INDEX] === "Aiada") {
          tempData[ALIGNMENT_INDEX] = "Ai";
          colorHex = "#ACD3C1";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Daimonia") {
          tempData[ALIGNMENT_INDEX] = "Da";
          colorHex = "#F8C6CF";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Elios") {
          tempData[ALIGNMENT_INDEX] = "El";
          colorHex = "#DC6A5E";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Peromei") {
          tempData[ALIGNMENT_INDEX] = "Pe";
          colorHex = "#FFEFC2";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Piphron") {
          tempData[ALIGNMENT_INDEX] = "Pi";
          colorHex = "#9AB2D3";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Sosyne") {
          tempData[ALIGNMENT_INDEX] = "So";
          colorHex = "#B3ACC4";
        }
        else if (tempData[ALIGNMENT_INDEX] === "Thras") {
          tempData[ALIGNMENT_INDEX] = "Th";
          colorHex = "#FFC469";
        }
        
        tempData.push(colorHex);
        rawData[j].data.push(tempData);
        break;
      }
    }
  }
  
  return rawData;
}