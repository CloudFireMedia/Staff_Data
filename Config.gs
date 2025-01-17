// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 6th Nov 2019 12.17 GMT
/* jshint asi: true */

(function() {"use strict"})()

// Code review all files - TODO
// JSHint review (see files) - Done
// Unit Tests - TODO
// System Test (Dev) - Done
// System Test (Prod) - Done

// Config.gs
// =========
//
// All the constants and configuration settings

// Configuration
// =============

var SCRIPT_NAME = 'Staff_Data'
var SCRIPT_VERSION = 'v1.9'

var PRODUCTION_VERSION_ = true

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? BBLog.Level.INFO : BBLog.Level.ALL
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? BBLog.DisplayFunctionNames.NO : BBLog.DisplayFunctionNames.YES

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false
var HANDLE_ERROR_ = Assert.HandleError.THROW
var ADMIN_EMAIL_ADDRESS_ = 'dev@andrewroberts.net'

// Test
// ----

var TEST_STAFF_DATA_SHEET_ID_ = '1oaf0fHVuvOhHb1MsZbuhwN_AQfhYFG2UHa05-Rz1n7A'  

var TEST_DELETE_ES_TAB = true
var TEST_SEND_EMAILS = true

if (PRODUCTION_VERSION_) {
  if (!TEST_DELETE_ES_TAB || !TEST_SEND_EMAILS) {
    throw new Error('Test flags set in production mode')
  }
}

// Constants/Enums
// ===============

var STAFF_DATA_SHEET_NAME_          = 'Staff Directory'
var QUICK_LOOK_SHEET_NAME_          = 'Staff Directory - Quick Look'
var QUICK_LOOK_TEMPLATE_SHEET_NAME_ = '[ TEMPLATE ] Staff Directory - Quick Look'

var STATUS_COLUMN_NUMBER_ = 6

var ES_TEMPLATE_SHEET_NAME_='TEMPLATE'

var NOTIFICATION_SUBJECT = 'Action required! Event Sponsorship calendar for "%s" has been deleted'

// Function Template
// -----------------

/**
 *
 *
 * @param {Object} 
 *
 * @return {Object}
 */
/* 
function functionTemplate() {
  Log_.functionEntryPoint()
  
  
} // functionTemplate() 
*/