var winston= require ( 'winston' )
require ( 'winston-daily-rotate-file' )
const moment = require('moment-timezone')
const packagejson = require ( './package.json' )
var fs = require('fs')
const path = require('path')
const util = require('util')

let baseLogPath = undefined
let logPath = undefined

if ( process.type === 'browser') { // 'browser' means we are in the main process
	if ( PROD ) { // Production (Main process)
		baseLogPath = require('electron').app.getPath('appData')
		logPath = path.join ( baseLogPath, packagejson.productName, 'log' )
	} else {  // Development (Main process)
		baseLogPath = require('electron').app.getAppPath()
		logPath = path.join ( baseLogPath, 'log' )
	}
} else { // if not in main process, we must be in the renderer process process.type will then be 'renderer'
	if ( global.PROD ) { // Production (Renderer process)
    	baseLogPath = require('@electron/remote').app.getPath('appData')
		logPath = path.join ( baseLogPath, packagejson.productName, 'log' )
	} else { // Production (Renderer process)
		baseLogPath = require('@electron/remote').app.getAppPath()
		logPath = path.join ( baseLogPath, 'log' )
	}
}

if ( !fs.existsSync ( logPath ) )
	fs.mkdirSync ( logPath, { recursive: true } )

const LOG_FILE = logPath + '/%DATE%logfile.txt';//error: 0, warn: 1, info: 2, verbose: 3, debug: 4, silly: 5

function getFormattedTimestamp ( withTimezone ) {
	var ts = moment().valueOf()
	var utcOffset = moment().utcOffset()

	if ( withTimezone ) {
		let tmp_str = moment(ts).utcOffset(utcOffset).format('YY-MM-DD HH:mm:ss.SSS UTCZ')
		
		if ( tmp_str.substr ( -3 ) === ':00' )
			tmp_str = tmp_str.substr ( 0, tmp_str.length -3 )

		return tmp_str
	} else
		return moment(ts).utcOffset(utcOffset).format('YY-MM-DD HH:mm:ss.SSS')
}

global.zipLogFileName = function () {
	let ts = moment().valueOf()
	let utcOffset = moment().utcOffset()
	return 'logs_' + moment(ts).utcOffset(utcOffset).format('YYYY_MM_DD') + '.zip'
}

function getLineNumber () {
	let stackError = new Error()
	let lines      = stackError.stack.split(/^\s+at/m)
	let lineNumber = "????"

	// Get line number fromstacktrace data
	for ( let i = 0 ; i < lines.length ; i++ ) {
		if ( lines[i].includes("global.log") ) {
			var elements   = lines[i+1].split(':')
			lineNumber = elements[elements.length-2]
		}
	}

	while ( lineNumber.toString().length < 4 )
		lineNumber = lineNumber + ' '

	return lineNumber
}

const wlog = winston.createLogger({
	format: winston.format.printf ( info => `${getFormattedTimestamp(true)} Ln ${getLineNumber()} [${info.level.toUpperCase()}] ${info.message}`),
	transports: [
		new winston.transports.DailyRotateFile ({
			filename    : LOG_FILE,
			datePattern : 'YYYY-MM-DD_',
			maxFiles    : '30d'
		})
	]
});

let inspectOptions = {
	showHidden:true,
	depth: null,
	getters: true
}

global.log = function () {
	console.log ( getFormattedTimestamp(false), "Ln " + getLineNumber() + " [INFO] ", ...arguments ) // same like console.info
	wlog.info   ( util.formatWithOptions(inspectOptions, ...arguments) ) // Write to logfile
}

global.logErr = function () {
	console.error ( getFormattedTimestamp(false), "Ln " + getLineNumber() + " [ERROR] ", ...arguments )
	wlog.error    ( util.formatWithOptions(inspectOptions, ...arguments) ) // Write to logfile
}

global.logWarn = function () {
	console.warn ( getFormattedTimestamp(false), "Ln " + getLineNumber() + " [WARN] ", ...arguments )
	wlog.warn    ( util.formatWithOptions ( inspectOptions, ...arguments) )
}

global.logInfo = function () {
	global.log ( ...arguments )
}