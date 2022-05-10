//@OnlyCurrentDoc
// above line is about permissions for running the script

// script name: googlesheets-apps-script-for-midi-commander-custom.gs
//
// this code should be pasted in "Extensions" > "Apps Script" from Google Spreadsheet
// it is intended to format the spreadsheet for better reading and editing of the settings
// for MeloAudio Midi-Commander-Custom project
//

// Install instructions (full process):
// 1. Open google spreadsheets (https://docs.google.com/spreadsheets)
// 2. Select Start a New Spreadsheet > Blank
// 3. Menu > File > Import > Select your CSV file 
// 4. At "Import File" popup > Import Location >  Replace current sheet 
//    At "Separator type" you may let it Detect automatically or select "comma"
//    You may check "Convert text to numbers, dates and formulas
//    Click "Import data"
// 5. Menu > Extensions > Apps Script 
//    Paste this entire file into the "code.gs" window
//	  Save the project in the icon at the left of ">Run" button 
//	  Click "Run" button to execute the project permissions
//	  After authorization, the code is run 
// 6. Go back to the spreadsheet window and click the new menu item
//    Menu > Midi-Commander-Custom > Format All	 
//	  You should have a formatted spreadsheet 
// 7. After editing, you may export to a CSV file
//    Menu > File > Download > Comma Separatede File (csv)

//colors
const LIGHTBLUE1 = '#cfe2f3';
const LIGHTBLUE2 = '#c9daf8';
const MAGENT1    = '#d5a6bd';
const MAGENT2    = '#ff00ff';
const YELLOW1    = '#fff2cc';
const BLACK      = '#000000';
const RED1       = '#dd7e6b';
const DARKGREEN1 = '#6aa84f';
const GRAY1      = '#bbbbbb';
const WHITE      = '#ffffff';
const GREEN1     = '#6aa84f';
const CHERY1     = '#e6b8af';
const RED2       = '#f72876';


var sheet, data_range, columnA;
var MAX_COLUMNS;
var MAX_ROWS;
var DEBUG = false;

set_globals();


//
// executed on spreadsheet open
//
function onOpen() {
  // adds menu entry to the spreadsheet
  SpreadsheetApp
    .getUi()
    .createMenu("Midi-Commander-Custom")
    .addItem("Format ALL", "run_all")
    .addItem("Only Colors", "set_conditional_format")
    .addItem("Only Aligns", "set_aligns")
    .addItem("Only Button Settings", "button_settings_header")
    .addToUi();
}

//
// The event handler triggered when editing the spreadsheet.
// @param {Event} e The onEdit event.
// @see https://developers.google.com/apps-script/guides/triggers#onedite
//
function onEdit(e) {
  var column = e.range.getColumn();
  var row = e.range.getRow();
  var val = e.range.getValue();                                              
  var is_keyword = false;
                                                  DEBUG && Logger.log('onEdit col:%s row:%s => %s (type:%s)' , column, row, val, typeof(val) );

  // sets reserved keywords to their expected case format                                                  
  if (typeof val != 'object') {
      val = val.toUpperCase();
      switch (val) {
        case 'CC':
        case 'PC':
        case 'PB':
        case 'Y' :
        case 'N' :
                    is_keyword = true;
                    break;
        case 'NOTE':
                    val = 'Note'
                    is_keyword = true;
                    break;        
        case 'START':
                    val = 'Start'
                    is_keyword = true;
                    break;        
        case 'STOP':
                    val = 'Stop'
                    is_keyword = true;
                    break;

      }
      if (is_keyword) { e.range.setValue(val); }
  }
}

//
//  
//
function set_globals() {

    sheet = SpreadsheetApp.getActiveSheet();      //WARNING: THIS TENDS TO RETURN THE FIRST SHEET, 
                                                  //SEE https://stackoverflow.com/a/54719592/873650
                                                  DEBUG &&   Logger.log('sheet name = %s', sheet.getSheetName());
    data_range = sheet.getDataRange();
    MAX_ROWS = data_range.getLastRow();
    MAX_COLUMNS = data_range.getLastColumn();
                                                  DEBUG &&   Logger.log('MAX_ROWS = %s\nMAX_COLUMNS = %s', MAX_ROWS, MAX_COLUMNS);
    //some decisions are made on column A 
    columnA    = sheet.getRange(1,1,MAX_ROWS,1);

}
//
// for "Apps Script" easy debugging
//
function debug_all() {

  DEBUG = true; 
  run_all();

}
//
// MAIN: streamline of actions
//
function run_all() {

  set_conditional_format();
  set_aligns();
  button_settings_header();

}

//
// find the header row for the Button_Settings session and 
// format it to be vertical text, so that we can get a better view
// 
function button_settings_header() {
  var tosearch = "Bank_Number";
  var all = columnA.createTextFinder(tosearch).findAll();     //all is a range
    
  for (var i = 0; i < all.length; i++) {
                                                      DEBUG && Logger.log('Sheet %s, cell %s = %s.', all[i].getSheet().getName(), all[i].getA1Notation(), all[i].getValue()); 
      next_column_cel = all[i].offset(0,1);                   
                                                      DEBUG && Logger.log('next_column is: ' + next_column_cel.getValue());
                                                      //test if is row of titles
      if (next_column_cel.getValue() == "Button_Identifier") {
                                                      DEBUG && Logger.log('next_column is Button_Identifier');
          var rowNum = all[i].getRow();                                            
          var row = sheet.getRange( rowNum, 1, 1, MAX_COLUMNS );                                                              
          row.setTextRotation(90);
          row.setHorizontalAlignment('center');
          row.setFontWeight('bold');
          
          sheet.setColumnWidths( 3, MAX_COLUMNS, 33);
          //vertical borders for dividing messages (A, B, C, ...)
          sheet.getRange(  'L' + rowNum + ':L' + MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'V' + rowNum + ':V' + MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'AF'+ rowNum + ':AF'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'AP'+ rowNum + ':AP'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'AZ'+ rowNum + ':AZ'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'BJ'+ rowNum + ':BJ'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'BT'+ rowNum + ':BT'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'CD'+ rowNum + ':CD'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          sheet.getRange(  'CN'+ rowNum + ':CN'+ MAX_ROWS  ).setBorder(false,false,false,true,false,false);
          
          row.setBorder(true,true,true,true,true,true);
      }
  }
}

// 
// Adds a conditional format rule to a sheet 
// 
function set_conditional_format(cellsInA1Notation) {
  var rule;
  var rules = [];
                                                  DEBUG && Logger.log('inside setConditionalFormat');
  // when cell = 0 (zero)
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('0') 
            .setBackground(LIGHTBLUE1)
            .setRanges([data_range])
            .build();
  rules.push(rule);
  
  // when cell > 0 
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(0) 
            .setBackground(LIGHTBLUE2)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  // when cell = PC (Program Change)
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('PC') 
            .setBackground(YELLOW1)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  // when cell = 'CC' 
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('CC') 
            .setBackground(MAGENT1)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  // when cell = 'Y' 
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Y') 
            .setBackground(GREEN1)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  // when cell = 'N'
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('N') 
            .setBackground(RED1)
            .setRanges([data_range])
            .build();
  rules.push(rule); 

  // when cell = 'Note'
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Note') 
            .setBackground(CHERY1)
            .setRanges([data_range])
            .build();
  rules.push(rule); 

  // when cell = 'PB'
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('PB') 
            .setBackground(MAGENT2)
            .setRanges([data_range])
            .build();
  rules.push(rule); 

  // when cell = 'Start'
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Start') 
            .setBackground(DARKGREEN1)
            .setRanges([data_range])
            .build();
  rules.push(rule); 

  // when cell = 'Stop'
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo('Stop') 
            .setBackground(RED2)
            .setRanges([data_range])
            .build();
  rules.push(rule); 

  // when cell = starts with '*' (sub-table)
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=IF( LEFT( TRIM($A1); 1 ) = "*"; true; false)'  )
            .setBackground(BLACK)
            .setFontColor(WHITE)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  // when cell = starts with '#' (comments)
  rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=IF( LEFT( TRIM($A1); 1 ) = "#"; true; false)'  )
            .setFontColor(GRAY1)
            .setRanges([data_range])
            .build();
  rules.push(rule);

  sheet.setConditionalFormatRules(rules);
 
}
//
// align rules. these are executed only once.
//
function set_aligns() {
    
  findRegexAlign('^[a-zA-Z].*', 'left');    //generictext
  findRegexAlign('^[0-9].*',    'right');   //numbers
  findRegexAlign('^[ABCD]{1}',  'right');   //banks
  findExactlyAlign('Y',    'center');
  findExactlyAlign('N',    'center');
  findExactlyAlign('CC',   'center');
  findExactlyAlign('PC',   'center');
  findExactlyAlign('Note', 'center');
  findExactlyAlign('PB',   'center');
  findExactlyAlign('Start','center');
  findExactlyAlign('Stop', 'center');
  findStartAlign('#', 'left');
  findStartAlign('*', 'left');

}
// 
// find regexp <text> and align it
// 
function findRegexAlign(text, align) {

  var rg = sheet.createTextFinder(text)
          .useRegularExpression(true)
          .matchEntireCell(true)
          .findAll();  
  for (var i = 0; i < rg.length; i++) {
          rg[i].setHorizontalAlignment(align);
  }

}
// 
// find exactly <text> and align it
// Case Sensitive = true
// 
function findExactlyAlign(text, align) {   
  var rg = sheet.createTextFinder(text)
           .matchEntireCell(true)
           .matchCase(true)
           .findAll();  
  for (var i = 0; i < rg.length; i++) {
          rg[i].setHorizontalAlignment(align);
  }
}
// 
// find text starting with <text> and align it
// Case Sensitive = true
// 
function findStartAlign(text, align) {
  var rg = sheet.createTextFinder(text)
           .matchCase(true)
           .findAll();  
  for (var i = 0; i < rg.length; i++) {
      if (rg[i].getValue().startsWith(text)) { 
          rg[i].setHorizontalAlignment(align); 
      }
  }
}