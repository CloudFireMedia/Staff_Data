// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Code review all files - TODO
// JSHint review (see files) - TODO
// Unit Tests - TODO
// System Test (Dev) - TODO
// System Test (Prod) - TODO

// Config.gs
// =========
//
// All the constants and configuration settings

// Configuration
// =============

var SCRIPT_NAME = 'Staff_Data'
var SCRIPT_VERSION = 'v0.dev_ajr'

var PRODUCTION_VERSION_ = false

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? BBLog.Level.INFO : BBLog.Level.FINER
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? BBLog.DisplayFunctionNames.NO : BBLog.DisplayFunctionNames.YES

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false
var HANDLE_ERROR_ = Assert.HandleError.THROW
var ADMIN_EMAIL_ADDRESS_ = ''

// Constants/Enums
// ===============

// var STAFF_FOLDER_ID_ = '1jGF-Md5vsJHP41FuUZ8XWtcSHKvRn87i'; // Live Staff folder
var STAFF_FOLDER_ID_ = '1NySJljnL7vyK6yHWf9LWaHBOikHJaQdR' // Andrew's Projects > Staff_Data > Staff

// var STAFF_DATA_SHEET_ID_ = '1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'; // Live CCN Staff Data Sheet
var STAFF_DATA_SHEET_ID_ = '1qx_Ok8F1-WOS6CnFvTeLUDtbkGTvFBDBSbsdZl_ANEc'; // Test Copy of CCN Staff Data Sheet (andrewr)

var STAFF_DATA_SHEET_NAME_ = 'Staff';
var QUICK_LOOK_SHEET_NAME_  = 'Quick Look';
var TEMPLATE_SHEET_NAME_    = 'Sheet Name';

var STATUS_COLUMN_NUMBER_ = 6;

// var CCN_EVENTS_PROMOTION_CALENDAR_ID_ = '1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc'; // Live CCN Events Promotion Calendar
var CCN_EVENTS_PROMOTION_CALENDAR_ID_ = '1Q3M3HOqrWeKchmlIYfz-f5T2GyFSvxveV_1xnoVf73U'; // Test Copy of  CCN Events Promotion Calendar (andrewr)

// var EVENT_TEMPLATE_ID_ = '1Idf44phe12hFr6TRk-h8qDFTnu_BanOP0H0swnLAARw'; // Live TEMPLATE CCN Events Promotion Calendar (DELETED - 29June2018)
var EVENT_TEMPLATE_ID_ = '1JgTf-2bD7LmAKmq8aczwlFxPPpNzLUKIUZA9ccK6dic'; // Test Copy of TEMPLATE CCN Events Promotion Calendar (andrewr)

var TEMPLATE_SHEET_NAME_='TEMPLATE';

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