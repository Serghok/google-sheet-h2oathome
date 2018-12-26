/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Prepare sheet...', functionName: 'prepareSheet_'},
    {name: 'Create new workshop...', functionName: 'createWorkshop'}
  ];
  spreadsheet.addMenu('H2O at home', menuItems);
}

function createWorkshop() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('NewWorkshop');
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Workshop'));
  Logger.log(spreadsheet.getSheetByName('Workshop').getRange("D4").getNumberFormat())
  ui.showModalDialog(html, "New workshop");
}

function insertWorkshop(form){    
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet= spreadsheet.getSheetByName('Workshop');
  
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1).setValue(form.idworkshop);
  sheet.getRange(2, 2).setValue(form.host);
  sheet.getRange(2, 3).setValue(form.date);
  sheet.getRange(2, 1, 1, 9).setBackground("#FBE5CC");
  
  var columns = ["Contact", "Prix", "Date de versement", "Commande sur"];
  sheet.insertRowAfter(2);
  createGrid(sheet, columns, false, 3, 3);
  var numParticipant = Number(form.numparticipant);
  for(var i = 0; i < numParticipant; i++){    
    sheet.insertRowAfter(3);  
    sheet.getRange(4, 6, 1, 1).setFormula("=C4");
  }
  sheet.getRange(3, 3, 1, 7).setBackground("#B2D7A7");
  sheet.getRange(3, 1, numParticipant + 1, 2).setBackground("#EFEFEF");
  sheet.getRange(4, 3, numParticipant, 7).setBackground("#FFFFFF");
  sheet.getRange(4, 3, numParticipant, 7).setFontWeight("Normal");
  sheet.getRange(4, 4, numParticipant, 1).setValue(0);
  sheet.getRange(2, 1, 1, 6).setFontWeight("Normal");
  // Set formula to get the sum of the workshop sale  
  sheet.getRange("D2").setFormula("=SUM(D4:D"+ (4 + numParticipant - 1) +")");
  sheet.getRange("E2").setFormula("=SUMIFS(D4:D"+ (4 + numParticipant - 1)+";E4:E"+ (4 + numParticipant - 1) +";\"<>\")");  
  sheet.getRange("F2").setFormula("=MAX(E4:E"+ (4 + numParticipant - 1) +")");
  sheet.getRange("H2").setFormula("=MULTIPLY(G2; 0,25)");
  
  sheet.getRangeList(["D2:E2","G2:H2"]).setNumberFormat("#,##0.00\ [$€-1]");  
  sheet.getRange("D4:D"+ (4 + numParticipant - 1)).setNumberFormat("#,##0.00\ [$€-1]");
  sheet.getRange("E4:E"+ (4 + numParticipant - 1)).setNumberFormat("DD/mm/YYYY");
  sheet.getRangeList(["C2","F2", "I2"]).setNumberFormat("DD/mm/YYYY");
  
  
  // Add conditional rule in the data sheet to display when someone doesn't paid or the workshop isn't completely paid    
  var rules = sheet.getConditionalFormatRules();  
  var ruleContact = null;
  var ruleAmount = null;
  var rulesToKeep = [];
  for(var i = 0; i < rules.length; i++)
  {
    rule = rules[i];
    var range = rule.getRanges();
    if(range[i].getA1Notation().indexOf("E") == 0)      
      ruleAmount = rule.copy();
    else if(range[i].getA1Notation().indexOf("C") == 0)
      ruleContact = rule.copy();
    else
      rulesToKeep.push(rule);
  }
  if(ruleAmount == null)
  {
    ruleAmount = SpreadsheetApp.newConditionalFormatRule()
       .whenFormulaSatisfied("=(D7 - E7) > 0")
       .setBackground("#F4C7C3")
       .setFontColor("#CC0000");
  }
  if(ruleContact == null)
  {
    ruleContact = SpreadsheetApp.newConditionalFormatRule()
       .whenFormulaSatisfied("=(E19 = \"\")")
       .setBackground("#F4C7C3")
       .setFontColor("#CC0000");
  }
  var range = ruleContact.getRanges();
  range.push(sheet.getRange("C4:C" + (4 + numParticipant - 1)));
  ruleContact.setRanges(range);
  
  range = ruleAmount.getRanges();
  range.push(sheet.getRange("E2"));
  ruleAmount.setRanges(range);
  
  rulesToKeep.push(ruleContact.build());
  rulesToKeep.push(ruleAmount.build());
  sheet.setConditionalFormatRules(rulesToKeep);  
  
  createWorkshopData(form);  
}

function createWorkshopData(form){  
  Logger.log("create data");
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet= spreadsheet.getSheetByName('Data');
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue(form.date);
  sheet.getRange(2, 2).setValue("Vente @ "+ form.host);
  sheet.getRange(2, 3).setFormula("=0-SUMIF(Workshop!A2:A10000;F2;Workshop!D2:D10000)");
  sheet.getRange(2, 4).setValue("Atelier vente");
  sheet.getRange(2, 5).setValue(form.date).setNumberFormat("mm-YY");
  sheet.getRange(2, 6).setValue(form.idworkshop);
  
  
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setFormula("=MAXIFS(Workshop!F2:F10000; Workshop!A2:A10000;F2)");
  sheet.getRange(2, 2).setValue("Paiement @ "+ form.host);  
  sheet.getRange(2, 3).setFormula("=SUMIF(Workshop!A2:A10000;F2;Workshop!E2:E10000)");
  sheet.getRange(2, 4).setValue("Atelier remboursement");
  sheet.getRange(2, 5).setValue(form.date).setNumberFormat("mm-YY");
  sheet.getRange(2, 6).setValue(form.idworkshop);
  
  
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setFormula("=MAXIFS(Workshop!I2:I10000; Workshop!A2:A10000;F3)");  
  sheet.getRange(2, 2).setValue("Commission @ "+ form.host);  
  sheet.getRange(2, 3).setFormula("=SUMIFS(Workshop!H2:H10000;Workshop!A2:A10000;F2;Workshop!I2:I10000;\"<>\")");
  sheet.getRange(2, 4).setValue("Commission");
  sheet.getRange(2, 5).setValue(form.date).setNumberFormat("mm-YY");
  sheet.getRange(2, 6).setValue(form.idworkshop);
  
  sheet.getRange("A2:H4").setBackground("#d9ead3");
}


/**
 * This method will check that Workshop, Data and Gran total sheet are created 
 * otherwise it will create them with correct formula and conditional format
 */
function prepareSheet_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet= spreadsheet.getSheetByName('Gran total');
  var needbuildGranTotal = false;
  if(sheet == null){
    sheet = spreadsheet.insertSheet('Gran total', 0);
    needbuildGranTotal = true;    
  }
  
  createDataSheet();
  createWorkshopSheet();
  Logger.log("Gran total test: "+ needbuildGranTotal);
  if(needbuildGranTotal)
    buildGranTotal();
  
  sheet= spreadsheet.getSheetByName('Workshop');
  var rules = sheet.getConditionalFormatRules();
  Logger.log("Rules length = " + rules.length);
  for(var i = 0; i < rules.length; i++){
    Logger.log(rules[i].getRanges()[0].getA1Notation());
    var booleanCondition = rules[i].getBooleanCondition();
    if(booleanCondition != null)
    {
      Logger.log("Boolean condition");
      Logger.log(booleanCondition.getCriteriaType());
      var color = booleanCondition.getBackground();
      Logger.log("The background color for rule is " + color);
      Logger.log("The font color for rule is " + booleanCondition.getFontColor());
      Logger.log("The criteria for rule is " + booleanCondition.getCriteriaValues());
    }
  }
  
}

/**
 * This method will create the data sheet that will contains all movement done in the account
 **/
function createDataSheet()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Data');
  var columns = ["Date", "Description", "Montant", "Type", "Mois d'application", "Atelier link", "Facture"];
  if(sheet == null)
  {
    sheet = spreadsheet.insertSheet('Data');
    createGrid(sheet, columns, true, 1);    
    sheet.getRange(1, 1, 1, columns.length).setBackground("#C9DAF8");
    sheet.getRange(2, 1, 10000, columns.length).setBackground("#FFFFFF");
    sheet.getRange("A1:A10000").setNumberFormat("DD/mm/YYYY");
    sheet.getRange("C1:C10000").setNumberFormat("#,##0.00\ [$€-1]");    
    sheet.getRange(2, 1, 10000, columns.length).setFontWeight("Normal");
    
    // Add conditional rule in the data sheet to display when an invoice is required but not yet classified
    var range = sheet.getRange("G2:G10000");
    var rules = sheet.getConditionalFormatRules();
    var ruleInvoice = SpreadsheetApp.newConditionalFormatRule();
    ruleInvoice.whenFormulaSatisfied("=AND(OR(D2=\"UCM\"; D2=\"Frais\"; D2=\"Commission\"); G2<> \"Oui\")")
        .setBackground("#F4C7C3")
        .setRanges([range])
        .build();
    rules.push(ruleInvoice);
    sheet.setConditionalFormatRules(rules);
  }
  
}

/**
 * This method will create the workshop sheet that will contains all data about workshop: contact who paied...
 **/
function createWorkshopSheet(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Workshop');
  var columns = ["Atelier",	"Hote", "Date", "Vente", "Rendu", "Date paiement", "Commissionable", "Commission", "Date commission"];
  if(sheet == null){
    sheet = spreadsheet.insertSheet('Workshop');
    createGrid(sheet, columns, true, 1);
    sheet.getRange(1, 1, 1, columns.length).setBackground("#C9DAF8");
  }
  
}

function createGrid(sheet, columns, isBold, startIndex, startColumn){
  if(startColumn == null)
    startColumn = 1;
  for(var i = 0; i < columns.length; i++)
      sheet.getRange(startIndex,(i+startColumn)).setValue(columns[i]);
  
  if(isBold){
    sheet.getRange(1, startColumn, 1, columns.length).setFontWeight("Bold");
  }
}

function buildGranTotal(){  
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetData = "Data";
  
  var pivotTableParams = {};
  pivotTableParams.source = {
    sheetId: spreadsheet.getSheetByName(sheetData).getSheetId()
  };
  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{
    sourceColumnOffset: 4,
    showTotals: true,
    sortOrder: "DESCENDING"
  }, {
    sourceColumnOffset: 4,
    showTotals: true,
    sortOrder: "DESCENDING"
  }, {
    sourceColumnOffset: 5,  
    showTotals: true,
    sortOrder: "ASCENDING"
  }];
  pivotTableParams.columns = [{
    sourceColumnOffset: 3,
    sortOrder: "ASCENDING"
  }];
  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [{
    summarizeFunction: "SUM",
    sourceColumnOffset: 2
  }];
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": spreadsheet.getSheetByName("Gran total").getSheetId()
      },
      "fields": "pivotTable"
    }
  };

  //Sheets.newPivotTable();
  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, spreadsheet.getId());
  
}