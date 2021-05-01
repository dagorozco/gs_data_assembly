function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('Advanced')
  .addItem('Copy from file', 'UpdateDB')
  .addItem('Update Collar', 'collar')
  .addItem('Update Survey', 'survey')
  .addItem('Update Lithology', 'lithology')
  .addItem('Update Assays', 'assays')
  .addToUi();
 }

function UpdateDB () {

  var database = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = database.getSheetByName('Property');
  var fileId = sheet.getActiveCell().getValue();
  Logger.log(fileId);
  var source = SpreadsheetApp.openById(fileId.toString());

// Update Collar tab
  var db_collar = database.getSheetByName("ImportedCollar");
  var src_collar = source.getSheetByName("collar");
  var LastRowColl = src_collar.getLastRow()
  var collar = src_collar.getRange(2, 1, src_collar.getLastRow(),7).getValues();
  db_collar.getRange(1, 1, db_collar.getLastRow()+LastRowColl, 7).setValues(collar);
  


//Update Survey tab
  var db_survey = database.getSheetByName("ImportedSurvey");
  var src_svy = source.getSheetByName("survey");
  var LastRowSvy = src_svy.getLastRow();
  var LastRowSurvey = db_survey.getLastRow();
  var survey = src_svy.getRange(2, 1, src_svy.getLastRow(),5).getValues();
  db_survey.getRange(1, 1, db_survey.getLastRow()+LastRowSvy, 5).setValues(survey);
 
//Update Assays tab 
  var db_assy = database.getSheetByName("ImportedAssays");
  var src_assy = source.getSheetByName("data");
  var LastRowAssy = src_assy.getLastRow();
  var LastRowAssay = db_assy.getLastRow()+1;
  var assays = src_assy.getRange(2, 1, src_assy.getLastRow(),7).getValues();
  db_assy.getRange(1, 1, db_assy.getLastRow()+LastRowAssy, 7).setValues(assays);

//Update Geology tab
  var db_geo = database.getSheetByName("ImportedLithology");
  var src_geo = source.getSheetByName("lithology");
  var LastRowGeo = src_geo.getLastRow()
  var LastRowGeology = db_geo.getLastRow()+1;
  var lithology = src_geo.getRange(2, 1, src_geo.getLastRow(),6).getValues();
  db_geo.getRange(1, 1, db_geo.getLastRow()+LastRowGeo, 6).setValues(lithology);
  
 }
 

function collar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ImportedCollar');
  var outputSheet = ss.getSheetByName('Collar');
  var range = sheet.getDataRange()
  var values = range.getValues();
  Logger.log(values);
  var fixed = values.join().split(',');
  Logger.log(values.length);
  for(var x=0;x<values.length;x++){
  outputSheet.appendRow(values[x]);
  }
  sheet.clear()
}

function survey(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ImportedSurvey');
  var outputSheet = ss.getSheetByName('Survey');
  var range = sheet.getDataRange()
  var values = range.getValues();
  Logger.log(values);
  var fixed = values.join().split(',');
  Logger.log(values.length);
  for(var x=0;x<values.length;x++){
  outputSheet.appendRow(values[x]);
  }
  sheet.clear()
 }

function lithology(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ImportedLithology');
  var outputSheet = ss.getSheetByName('Lithology');
  var range = sheet.getDataRange()
  var values = range.getValues();
  Logger.log(values);
  var fixed = values.join().split(',');
  Logger.log(values.length);
  for(var x=0;x<values.length;x++){
  outputSheet.appendRow(values[x]);
  }
  sheet.clear()
 }
 
function assays(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ImportedAssays');
  var outputSheet = ss.getSheetByName('Assays');
  var range = sheet.getDataRange()
  var values = range.getValues();
  Logger.log(values);
  var fixed = values.join().split(',');
  Logger.log(values.length);
  for(var x=0;x<values.length;x++){
  outputSheet.appendRow(values[x]);
  }
  sheet.clear()
 }
