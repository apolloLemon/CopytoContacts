/* If a Reviewer is looking at this code:
I didn't realise that it has to be "ready to ship"
I'm mainly trying to hacktogether a tool for me dad.
*/

var data = PropertiesService.getScriptProperties();
var fields = ["EMAIL","LAST_NAME","FIST_NAME","PHONE"];
var range = ["START","END"];


function onInstall(e) {
  onOpen(e);
  data.setProperty("GROUP", "System Group: My Contacts");
}

function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Update Contacts', functionName: 'updateContacts'},
    {name: 'Add All', functionName: 'CopyAllToContacts'},
    {name: 'Set all param',functionName: 'setALL'},
    {name: 'Set Fields', functionName: 'setFields'},
    {name: 'Set First Line',functionName: 'setRangeSTART'},
    {name: 'Set Group',functionName: 'setGROUP'},
    {name: 'logFiels', functionName: 'logFields'}
  ];
  spreadsheet.addMenu('Copy To Contacts', menuItems);
  
}

function logFields () {
  Logger.log(data.getProperties());
}

/*function onEdit(e) {
  var line = e.range;
  Logger.log(line);
  Logger.log(edits);
}*/

//function incrEdits () {edits++;}

function setALL () {
  setFields();
  setRangeSTART();
  setGROUP();
}
 
function setFields () {
  var spreadsheet = SpreadsheetApp.getActive();
  
  for(var i in fields) {
    Logger.log(i +': '+ fields[i]);
    var selectedCol = Browser.inputBox(fields[i] + ' Column',
                                      'Please enter a column letter for the '+fields[i]+' field',
                                       Browser.Buttons.OK_CANCEL);
    if(selectedCol=='cancel') continue;
    data.setProperty(fields[i], selectedCol);
  }
  
  Logger.log(data.getProperties());
}

function setRangeSTART () {
  var spreadsheet = SpreadsheetApp.getActive();
  
  var selected = Browser.inputBox('First Line',
                                      'Which is the first line of data?',
                                       Browser.Buttons.OK_CANCEL);
  if(selected=='cancel') return;
  data.setProperty(range[0], selected);
  
  Logger.log(data.getProperties());
}

function setGROUP () {
  var spreadsheet = SpreadsheetApp.getActive();
  
  var selected = Browser.inputBox('Contacts Group',
                                      'Which Contact Group will these contacts belong too?',
                                       Browser.Buttons.OK_CANCEL);
  if(selected=='cancel') return;
  else if (selected=='default'||selected=='none') data.setProperty("GROUP", "System Group: My Contacts"); 
  else {
    data.setProperty("GROUP", selected);
  }
    
  Logger.log(data.getProperties());
}

function setGroup_Default () {
  data.setProperty("GROUP", "System Group: My Contacts");
}


function setToZero () {
  setActiveValue(0);
}

function getActiveValue() {
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return cell.getValue();
}

function setActiveValue(value) {
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  cell.setValue(value);
}

function getActiveCol() {
  return SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
}

function getCellValue(i,j){
  var cell = SpreadsheetApp.getActiveSheet().getRange(i, j);
  return cell.getValue();
}

function getCellValue(code){
  var cell = SpreadsheetApp.getActiveSheet().getRange(code);
  return cell.getValue();
}

function CopyAllToContacts(){
  Logger.log("Copying All to Contacts");
  var contactdata = [];
  var firstline = parseInt(data.getProperty("START"));
  var lastline = SpreadsheetApp.getActive().getLastRow();
  
  Logger.log(firstline +" "+ lastline);
  
  for(var i=firstline;i<=lastline;i++){
    for(var j in fields) {
      Logger.log(data.getProperty(fields[j]) + String(i) + " = " + getCellValue(
        data.getProperty(fields[j]) + String(i)
      ));
      contactdata[j] = getCellValue(
        data.getProperty(fields[j]) + String(i)
      );
    }
    Logger.log(contactdata);
    CopyToContacts(contactdata);
  }
}

function CopyTestContact(){
  CopyToContacts(["kappa@bob.com","kappa","boby","696969"]);
}

function CopyToContacts(Cells) {
  Logger.log(Cells);
  var email = Cells[0];
  var lastName = Cells[1];
  var firstName = Cells[2];
  var phoneNum = Cells[3];
  Logger.log(email);
  var contactGroup = ContactsApp.getContactGroup(data.getProperty("GROUP"));
  
  var Temp = ContactsApp.getContact(email);
  /*if(Temp==null) return;
  Logger.log(Temp.getFullName());
  Logger.log(Temp.getEmails()[0].getAddress());*/
  if(Temp==null){
    Logger.log(Temp);
    Temp = ContactsApp.createContact(firstName, lastName, email);
    //Temp = ContactsApp.createContact("none", "yet", "not@yet.c");
    if(phoneNum != "") Temp.addPhone("Phone",phoneNum);
    //Temp.addToGroup(contactGroup);
    if (contactGroup != null) contactGroup.addContact(Temp);
 
  }
}