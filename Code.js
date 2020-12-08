function onOpen(e) {
  //create menu items
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Start', 'openSidebar')
      .addToUi();
  
  //check if this is a valid sheet
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('v2.1').getRange('AQ4').getValue() != "2.1") {
    Browser.msgbox("Please make sure you are using GSheet v2.1")
  }
}

//make sure that onOpen is completed on install
function onInstall(e) {
  onOpen(e);
}

//Open the sidebar
function openSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle("Dice Roller");
  SpreadsheetApp.getUi().showSidebar(ui);
}

function RollDice(size, pic, mod) {
  //generate random number
  var roll = Math.floor(Math.random() * size) + 1 + parseInt(mod)
  
  //get the user
  var user = getUser()
  
  //JSON payload
  var payload = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(
      {
        "username": user,
        "avatar_url": pic,
        "embeds": [
          {
            "title": "**" + roll + "**",
            "thumbnail": {"url": pic},
            "description": "d" + size
          }
        ]
      })
  };
  
  postMessageToDiscord(payload)
};

function d4(){
  RollDice(4, 'https://i.imgur.com/3ddlkjE.png', 0)
}

function d6(){
  RollDice(6, 'https://i.imgur.com/RYR5pEn.png', 0)
}

function d8(){
  RollDice(8, 'https://i.imgur.com/YFZxpbp.png', 0)
}

function d10(){
  RollDice(10, 'https://i.imgur.com/TiccOq2.png', 0)
}

function d12(){
  RollDice(12, 'https://i.imgur.com/gDLc7I4.png', 0)
}

function d20(){
  RollDice(20, 'https://i.imgur.com/8DZXfoR.png', 0)
}

function d100(){
  RollDice(100, 'https://i.imgur.com/CU1raT4.png', 0)
}

function Roll(){
  //get active cell
  var activeCell = SpreadsheetApp.getActiveSheet().getActiveCell()
  
  //get the user
  var user = getUser()
 
  if (activeCell.getRow() >= 32 && activeCell.getRow() <= 36 && activeCell.getColumn() === 29) {
    rollDamage(user)
  } else {
    rolld20(SpreadsheetApp.getActiveSheet().getActiveCell().getDisplayValue(), user) //roll a d20 with the modifier found in the active cell
  }
}

function rolld20(active, user) { //roll a d20 with mod  
  if (active.indexOf("+") == 0) { //is this a valid cell 
    var result = Math.floor(Math.random() * 20) + 1 + parseInt(active)
    var pic = 'https://i.pinimg.com/originals/48/cb/53/48cb5349f515f6e59edc2a4de294f439.png'
    var roll = 'd20' + active
    
    //JSON payload
    var payload = {
      'method' : 'post',
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string.
      'payload' : JSON.stringify(
        {
          "username": user,
          "avatar_url": pic,
          "embeds": [
            {
              "title": "**" + result + "**",
              "thumbnail": {"url": pic},
            "description": roll
            }
          ]
        })
    };
    
    postMessageToDiscord(payload)
  } else {
    SpreadsheetApp.getUi().alert('Please make sure to select a valid cell before rolling.');
  }
}
  
function rollDamage(user) {
  var s = SpreadsheetApp.getActiveSheet()
  var activeRow = s.getActiveCell().getRow() //get the currently active row
  
  //get info about selected weapon
  var name = s.getRange(activeRow, 43).getDisplayValue()
  var formula = s.getRange(activeRow, 44).getDisplayValue()
  var damageType = s.getRange(activeRow, 45).getDisplayValue()
  var attackType = s.getRange(activeRow, 46).getDisplayValue()
  var desc = s.getRange(activeRow, 53).getDisplayValue()
  var mod = 0
    
  var result = 0 
  var pic = setPicture(attackType)
  
  getModifier(name, activeRow);  
  
  if (formula.indexOf('d') + 1) { //is there a dice to roll
    var numberDice = formula.substring(0, formula.indexOf('d')); 
    var diceSize = formula.substring(formula.indexOf('d') + 1);
    
    //roll the dice
    for (i = 0; i < numberDice; i++) {
      result += Math.floor(Math.random() * diceSize) + 1 + parseInt(mod);
    }     
    } else { //is there just a number
      result = formula
    }
  
  //create the expression to show the user
  var expression
  if (mod != 0) {
    expression = formula + mod
  } else {
    expression = formula
  }
  
  //JSON payload
  var payload = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(
      {
        "username": user,
        "avatar_url": pic,
        "embeds": [
          {
            "title": "**" + result + "**",
            "thumbnail": {"url": pic},
          "description": name + ": " + expression + " " + damageType + " damage \n *" + desc + "*"
          }
        ]
      })
  };

  //post to Discord
  postMessageToDiscord(payload)
}

function postMessageToDiscord(payload) {
  var discordUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK');
  UrlFetchApp.fetch(discordUrl, payload);
}

function getUser() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('v2.1').getRange('AE5').getValue()
}

function setPicture(attackType) {
  var pic

  if (attackType === "Spell") {
    pic = 'https://images.vexels.com/media/users/3/211391/isolated/preview/2af92aec47b1fa4289a190c5fa7ad94c-magic-spell-book-icon-by-vexels.png'
  } else if (attackType === "Ranged") {
    pic = 'https://cdn.pixabay.com/photo/2020/02/18/05/01/bow-4858463_1280.png'
  } else {
    pic = 'https://webstockreview.net/images/clipart-sword-vector-1.png' 
  } 
  return pic
}

function getModifier(searchFor, activeRow) {
  //Search for the weapon in Attack Info and find which row it's on
  var lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attack Info').getLastRow()
  var arrayToSearch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attack Info').getRange(7, 2, lastRow, 39).getValues()

  for(i = 0; i < arrayToSearch.length; i++){
    if (arrayToSearch[i][0] == searchFor){
      break
    }
  }

  Browser.msgBox(arrayToSearch[i][38]);

  //return the modifier if we are supposed to add it to the attack
  if (arrayToSearch[i][38] === '◉') {
    return SpreadsheetApp.getActiveSheet.getRange(activeRow, 55).getDisplayValue()
  }
}