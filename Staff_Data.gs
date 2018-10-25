// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Staff_Data.gs
// =============
//
// Dev: AndrewRoberts.net
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet
//
// The filename is prepended with _API as the Github chrome extension won't 
// push a file with the same name as the project.

var Log_

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls; they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

var EVENT_HANDLERS_ = {

//                           Name                            onError Message                          Main Functionality
//                           ----                            ---------------                          ------------------

  initialize:                ['initialize()',                'Failed to initialize',                  initialize_],
  initialized:               ['initialized()',               'Failed to get initialized',             initialized_],  
  clearConfig:               ['clearConfig()',               'Failed to clear config',                clearConfig_],    
  staffFolders:              ['staffFolders()',              'Failed to sort staff folders',          staffFolders_],
  maintainPromotionCalendar: ['maintainPromotionCalendar()', 'Failed to maintain promotion calendar', maintainPromotionCalendar_],
  onEdit:                    ['onEdit()',                    'Failed to process edit',                onEdit_],
}

function initialize(arg1)                {return eventHandler_(EVENT_HANDLERS_.initialize,                arg1)}
function initialized(arg1)               {return eventHandler_(EVENT_HANDLERS_.initialized,               arg1)}
function clearConfig(arg1)               {return eventHandler_(EVENT_HANDLERS_.clearConfig,               arg1)}
function staffFolders(arg1)              {return eventHandler_(EVENT_HANDLERS_.staffFolders,              arg1)}
function maintainPromotionCalendar(arg1) {return eventHandler_(EVENT_HANDLERS_.maintainPromotionCalendar, arg1)}
function onEdit(arg1)                    {return eventHandler_(EVENT_HANDLERS_.onEdit,                    arg1)} 

// Private Functions
// =================

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {Array} config:
 *   [0] {Function} prefunction
 *   [1] {String} eventName
 *   [2] {String} onErrorMessage
 *   [3] {Function} mainFunction
 *
 * @param {Object}   arg1       The argument passed to the top-level event handler
 */

function eventHandler_(config, arg1) {

  try {

    var userEmail = Session.getEffectiveUser().getEmail()
    
    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    
    Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    
    // Call the main function
    return config[2](arg1)
    
  } catch (error) {
  
    var handleError = Assert.HandleError.DISPLAY_FULL

    if (!PRODUCTION_VERSION_) {
      handleError = Assert.HandleError.THROW
    }

    var assertConfig = {
      error:          error,
      userMessage:    config[1],
      log:            Log_,
      handleError:    handleError, 
      sendErrorEmail: SEND_ERROR_EMAIL_, 
      emailAddress:   ADMIN_EMAIL_ADDRESS_,
      scriptName:     SCRIPT_NAME,
      scriptVersion:  SCRIPT_VERSION, 
    }

    Assert.handleError(assertConfig) 

  }
  
} // eventHandler_()

// Private event handlers
// ----------------------

/*

// onChange
{authMode=FULL, changeType=REMOVE_ROW, source=Spreadsheet, user=andrewr1969@gmail.com, triggerUid=822831230}

// onEdit
{authMode=FULL, range=Range, oldValue=1.0, source=Spreadsheet, value=2, user=andrewr1969@gmail.com, triggerUid=1616728844}

*/

function onEdit_(event) {

  Log_.functionEntryPoint();
  Log_.fine('event: %s', event);
  var runRecreateQuickLook = false;
  var sp;
    
  if (!event.hasOwnProperty('triggerUid')) {
    
    // Ignore the simple edit trigger, as the installable is the only one that has auth 
    return;
  }
  
  if (!event.hasOwnProperty('range')) {
    
    // We've not been told where the change is so do it anyway
    runRecreateQuickLook = true;
    
    Log_.warning('No range property');
    return
    
  } else {
    
    var range = event.range;
    var sh = range.getSheet();
    var shName = sh.getName();
    
    if (shName !== STAFF_DATA_SHEET_NAME_) {
      return
    }
    
    sp = sh.getParent();
    var strValue = event.value;
    var oldStrValue = event.oldValue;
    
    Log_.fine('strValue: ' + strValue)
    Log_.fine('oldStrValue: ' + oldStrValue)
    Log_.fine('sheet name: ' + shName)
    
    if (strValue !== oldStrValue) {
      
      if (range.getColumn() === STATUS_COLUMN_NUMBER_ && strValue === 'No longer employed') {        
        sh.hideRows(range.getRow(), 1);
      }  
      
      runRecreateQuickLook = true;
    }
  }
  
  if (runRecreateQuickLook) {
    recreateQuickLook();      
  }
  
  return;
  
  // Private Functions
  // -----------------
    
  function recreateQuickLook() {
  
    var sh = sp.getSheetByName(STAFF_DATA_SHEET_NAME_);
    var shDst = sp.getSheetByName(QUICK_LOOK_SHEET_NAME_);
    var shTmpl = sp.getSheetByName(QUICK_LOOK_TEMPLATE_SHEET_NAME_);
    var data = sh.getDataRange().getValues();
    
    var arrDest = []
    var counter = 1;
    
    for (var i = 2; i < data.length; i++) {
    
      var firstName  = data[i][0];
      var lastName   = data[i][1];
      var jobTitle   = data[i][4]
      var status     = data[i][5];
      var ucStatus   = '' + status.trim().toUpperCase() // cast as string    
      var extension  = data[i][6];
      var cell       = data[i][7];
      var secondCell = data[i][9];
      
      if (firstName != "" && 
          (ucStatus == "FULL-TIME" || ucStatus == "PART-TIME" || ucStatus == "VOLUNTEER")) {
          
        var row = [counter, extension, lastName + ', ' + firstName, jobTitle, cell, secondCell, status];
        arrDest.push(row);
        counter++;     
      }    
    }
    
    shDst.clear();
    var lr = shDst.getMaxRows();
    
    if (lr > 1) {
      shDst.deleteRows(2, lr - 1);
    }
    
    var lastTemplateRow = shTmpl.getLastRow()
    
    var rng = shTmpl.getRange(1, 1, lastTemplateRow, 9);
    rng.copyTo(shDst.getRange("A1"));
    var dates = shDst.getRange("A17").getValue();
    var timeZone = Session.getScriptTimeZone()
    dates = ('' + dates).replace('{date_updated_in_MM_/_DD_/_YY}', Utilities.formatDate(new Date(), timeZone, "MM/dd/yy"))
    shDst.getRange("A17").setValue(dates)
  
    if (arrDest.length !== 0) {
  
      if (arrDest.length > 1) {
        shDst.insertRows(4, (arrDest.length - 1));
      }
    
      shDst.getRange(4, 2, arrDest.length, arrDest[0].length).setValues(arrDest);
    }
    
    Log_.info('Updated "Quick Look" tab')
    
  } // onEdit_.recreateQuickLook()  
  
} // onEdit_()

/**
 * Clear the config
 */

function clearConfig_() {

  Log_.functionEntryPoint()
  
  var properties = PropertiesService.getDocumentProperties()
  var triggerId = properties.getProperty('triggerId')
  
  if (triggerId !== null) {
    properties.deleteProperty('triggerId')
    Log_.info('Cleared "onInstallableEdit" onEdit trigger id:' + triggerId)
  }

} // clearConfig_() 

/**
 * Initialize the library
 */
 
function initialize_() {

  Log_.functionEntryPoint()
  
  var staffDataSpreadsheetId = Config.get('STAFF_DATA_GSHEET_ID')  
  Log_.fine('staffDataSpreadsheetId: ' + staffDataSpreadsheetId)
  
  var triggerId = ScriptApp
    .newTrigger('onInstallableEdit')
    .forSpreadsheet(staffDataSpreadsheetId)
    .onEdit()
    .create()
    .getUniqueId()

  var properties = PropertiesService.getDocumentProperties()
  
  if (properties.getProperty('triggerId') !== null) {
    throw new Error('There is already a trigger ID stored')
  }

  properties.setProperty('triggerId', triggerId)

  Log_.info('Created "onInstallableEdit" onEdit trigger: ' + triggerId)

} // initialize_()

/**
 * Check if the library is initialized
 */
 
function initialized_(e) {

  Log_.functionEntryPoint()
  
  var found = false
  
  if (PropertiesService.getDocumentProperties().getProperty('triggerId') !== null) {
    found = true
  }

  return found

} // initialize_() 
