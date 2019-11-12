// JSHint - 6th Nov 2019 12.17 GMT
/* jshint asi: true */

(function() {"use strict"})()

// Staff_Data.gs
// =============
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

  onInstall:                 ['onInstall()',                 'Failed to install',                     onInstall_],
  initialize:                ['initialize()',                'Failed to initialize',                  initialize_],
  staffFolders:              ['staffFolders()',              'Failed to sort staff folders',          staffFolders_],
  maintainPromotionCalendar: ['maintainPromotionCalendar()', 'Failed to maintain promotion calendar', maintainPromotionCalendar_],
  onEdit:                    ['onEdit()',                    'Failed to process edit',                onEdit_],
  onChange:                  ['onChange()',                  'Failed to process change',              onChange_],
}

function onInstall(arg1)                 {return eventHandler_(EVENT_HANDLERS_.onInstall,                 arg1)}
function initialize(arg1)                {return eventHandler_(EVENT_HANDLERS_.initialize,                arg1)}
function staffFolders(arg1)              {return eventHandler_(EVENT_HANDLERS_.staffFolders,              arg1)}
function maintainPromotionCalendar(arg1) {return eventHandler_(EVENT_HANDLERS_.maintainPromotionCalendar, arg1)}
function onEdit(arg1)                    {return eventHandler_(EVENT_HANDLERS_.onEdit,                    arg1)}
function onChange(arg1)                  {return eventHandler_(EVENT_HANDLERS_.onChange,                  arg1)}

// These can be opened in various authModes so it needs to be outside eventHandler_() 
function onOpen(arg1) {onOpen_(arg1)}

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

    var userEmail = 'unknown email'
    var initializeLog = false

    if (arg1 !== undefined && arg1 instanceof Object && arg1.hasOwnProperty('authMode')) {
    
      // arg1 is an event so need to check authMode
      if (arg1.authMode !== ScriptApp.AuthMode.NONE) { // LIMITED or FULL

        userEmail = Session.getEffectiveUser().getEmail()
        initializeLog = true
      }
      
    } else {

      // arg1 is not an event so assume we have sufficient auth to do the following
      userEmail = Session.getEffectiveUser().getEmail()
      initializeLog = true   
    }
    
    if (initializeLog) {
    
      Log_ = BBLog.getLog({
        level:                DEBUG_LOG_LEVEL_, 
        displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
      })
      
      Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    }

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

function onInstall_(event) {

  Log_.functionEntryPoint()
  
  var properties = PropertiesService.getDocumentProperties()
  var triggerEditId = properties.getProperty('triggerEditId')
  var triggerChangeId = properties.getProperty('triggerChangeId')
  
  if (triggerEditId !== null) {
    properties.deleteProperty('triggerEditId')
    Log_.info('Cleared "onSDInstallableEdit" onEdit trigger id:' + triggerEditId)
  }
  
  if (triggerChangeId !== null) {
    properties.deleteProperty('triggerChangeId')
    Log_.info('Cleared "onSDInstallableChange" onChange trigger id:' + triggerChangeId)
  }
  
  onOpen_(event)
  
} // onInstall_()

function onOpen_(event) {

  var menu = SpreadsheetApp.getUi().createMenu('CloudFire')
  
  if (event.authMode === ScriptApp.AuthMode.NONE) {
  
    menu.addItem('Start', 'initialize')   

  } else { // LIMITED or FULL
  
    var triggerEditId = PropertiesService.getDocumentProperties().getProperty('triggerEditId')
    var triggerChangeId = PropertiesService.getDocumentProperties().getProperty('triggerChangeId')
    var triggerId = PropertiesService.getDocumentProperties().getProperty('triggerId')
    
    if (triggerEditId !== null && triggerChangeId !== null) {
      
      // Check if a previous version of the onEdit has been installed by checking if a triggerId exists
      // If triggerId exists run the initialize again
      if (triggerId !== null) {
              
        menu.addItem('Start', 'initialize')
      } else {
  
        menu
          .addItem("Update Staff Folders in Google Drive", "staffFolders")
          .addItem("Update Event Sponsorship Pages for Teams", "maintainPromotionCalendar")   
      }
  
    } else {
  
      menu.addItem('Start', 'initialize')   
    }
  }
          
  menu.addToUi()
}

/*

// onChange
{authMode=FULL, changeType=REMOVE_ROW, source=Spreadsheet, user=andrewr1969@gmail.com, triggerUid=822831230}
*/

function onChange_(event) {
  
  Log_.functionEntryPoint()
  
  var sp
  
  Log_.fine('triggerUid ' + event.triggerUid)
  Log_.fine('Change Type: ' + event.changeType)

  // The "event" object does not seem to be derived from Object, so 
  // hasOwnProperty() can't be used
  if (event.triggerUid === undefined) {
    
    // Ignore the simple edit trigger, as the installable is the only one that has auth    
    Log_.fine('No triggerUid')
    return
  }
  
  if (event.changeType === undefined) {
    
    // The change type is undefined
    Log_.fine('Change Type: undefined')
    return
  }
    
  sp = SpreadsheetApp.getActive()
  recreateQuickLook_(sp)      
  
  return
}


/*
// onEdit
{authMode=FULL, range=Range, oldValue=1.0, source=Spreadsheet, value=2, user=andrewr1969@gmail.com, triggerUid=1616728844}

*/

function onEdit_(event) {
  
  Log_.functionEntryPoint()
  
  var runRecreateQuickLookEdit = false
  var sp

  Log_.fine('triggerUid ' + event.triggerUid)
  
  // The "event" object does not seem to be derived from Object, so 
  // hasOwnProperty() can't be used 
  if (event.triggerUid === undefined) {
    
    // Ignore the simple edit trigger, as the installable is the only one that has auth     
    Log_.fine('No triggerUid')
    return
  }
  
  if (event.range === undefined) {
    
    // We've not been told where the change is so do it anyway   
    runRecreateQuickLookEdit = true    
    
  } else {
    
    var range = event.range
    var sh = range.getSheet()
    var shName = sh.getName()
    
    if (shName !== STAFF_DATA_SHEET_NAME_) {
      return
    }
    
    sp = sh.getParent()
    var strValue = event.value
    var oldStrValue = event.oldValue
    
    Log_.fine('strValue: ' + strValue)
    Log_.fine('oldStrValue: ' + oldStrValue)
    Log_.fine('sheet name: ' + shName)
    
    if (strValue !== oldStrValue) {
      
      if (range.getColumn() === STATUS_COLUMN_NUMBER_ && strValue === 'No longer employed') {        
        sh.hideRows(range.getRow(), 1)
      }  
      
      runRecreateQuickLookEdit = true
    }
  }
  
  if (runRecreateQuickLookEdit) {
    
    recreateQuickLook_(sp)      
  }
  
  return
  
} // onEdit_()

// Private Functions
// -----------------
    
function recreateQuickLook_(sp) {
  
  var sh = sp.getSheetByName(STAFF_DATA_SHEET_NAME_)
  var shDst = sp.getSheetByName(QUICK_LOOK_SHEET_NAME_)
  var shTmpl = sp.getSheetByName(QUICK_LOOK_TEMPLATE_SHEET_NAME_)
  var data = sh.getDataRange().getValues()
    
  // Get a script lock, because we're about to modify a shared resource.
  var lock = LockService.getDocumentLock()
  
  // Wait for up to 10 milliseconds for other processes to finish.  
  if (!lock.tryLock(10)) {
    Log_.fine('Could not obtain lock.')
    return
  }
  
  var arrDest = []
  var counter = 1
    
  for (var i = 2; i < data.length; i++) {
    
    var firstName  = data[i][0]
    var lastName   = data[i][1]
    var jobTitle   = data[i][4]
    var status     = data[i][5]
    var ucStatus   = '' + status.trim().toUpperCase() // cast as string    
    var extension  = data[i][6]
    var cell       = data[i][7]
    var secondCell = data[i][9]
     
    if (firstName != "" && 
        (ucStatus == "FULL-TIME" || ucStatus == "PART-TIME" || ucStatus == "VOLUNTEER")) {
         
      var row = [counter, extension, lastName + ', ' + firstName, jobTitle, cell, secondCell, status]
      arrDest.push(row)
      counter++     
    }    
  }
    
  shDst.clear()
  var lr = shDst.getMaxRows()
    
  if (lr > 1) {
    shDst.deleteRows(2, lr - 1)
  }
  
  var lastTemplateRow = shTmpl.getLastRow()
    
  var rng = shTmpl.getRange(1, 1, lastTemplateRow, 9)
  rng.copyTo(shDst.getRange("A1"))
  var dates = shDst.getRange("A17").getValue()
  var timeZone = Session.getScriptTimeZone()
  dates = ('' + dates).replace('{date_updated_in_MM_/_DD_/_YY}', Utilities.formatDate(new Date(), timeZone, "MM/dd/yy"))
  shDst.getRange("A17").setValue(dates)
  
  if (arrDest.length !== 0) {
  
    if (arrDest.length > 1) {
      shDst.insertRows(4, (arrDest.length - 1))
    }
    
    shDst.getRange(4, 2, arrDest.length, arrDest[0].length).setValues(arrDest)
  }
  
  lock.releaseLock()
   
} // recreateQuickLook_()  

/**
 * Initialize the library
 *
 * @param {object} event
 */
 
function initialize_(event) {

  Log_.functionEntryPoint()
  
  var staffDataSpreadsheetId = SpreadsheetApp.getActive().getId()
  
  // Delete all project triggers
  ScriptApp.getProjectTriggers().forEach(function(trigger){

    var triggerId = trigger.getUniqueId()
    
    ScriptApp.deleteTrigger(trigger)
    Log_.info('Deleted trigger: ' + triggerId)
  })
   
  var properties = PropertiesService.getDocumentProperties()
  
  //Delete the trigger ID properties also
  properties.deleteProperty('triggerEditId')
  properties.deleteProperty('triggerChangeId')
  properties.deleteProperty('triggerId')
  
  var triggerEditId = ScriptApp
    .newTrigger('onSDInstallableEdit')
    .forSpreadsheet(staffDataSpreadsheetId)
    .onEdit()
    .create()
    .getUniqueId()
  
  var triggerChangeId = ScriptApp
    .newTrigger('onSDInstallableChange')
    .forSpreadsheet(staffDataSpreadsheetId)
    .onChange()
    .create()
    .getUniqueId()

  if (properties.getProperty('triggerEditId') !== null) {
    throw new Error('There is already a trigger Edit ID stored')
  }
  
  if (properties.getProperty('triggerChangeId') !== null) {
    throw new Error('There is already a trigger Change ID stored')
  }

  properties.setProperty('triggerEditId', triggerEditId)
  properties.setProperty('triggerChangeId', triggerChangeId)
  
  // Refresh the menu (initialize is only ever called from menu, when authMode is FULL)
  onOpen_({authMode: ScriptApp.AuthMode.FULL})

  Log_.info('Created "onSDInstallableEdit" onEdit trigger: ' + triggerEditId)
  Log_.info('Created "onSDInstallableChange" onChange trigger: ' + triggerChangeId)
  
  // Create the Quick Look after initialization
  recreateQuickLook_(SpreadsheetApp.getActive())   

} // initialize_()