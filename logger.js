var winston= require ( 'winston' )
require ( 'winston-daily-rotate-file' )
const moment = require('moment-timezone')
var fs = require('fs')

if ( process.type === 'browser') // 'browser' means we are in the main process
    global.ABS_APP_ROOT = require('electron').app.getAppPath()
else // if not in main process, we must be in the renderer process process.type will then be 'renderer'
    global.ABS_APP_ROOT = require('@electron/remote').app.getAppPath()

ABS_APP_ROOT = ABS_APP_ROOT.replace ( /\/resources\/app.asar|\\resources\\app.asar/, '' ); // If running with asar packaged app directory
ABS_APP_ROOT = ABS_APP_ROOT.replace ( /\/resources\/app|\\resources\\app/          , '' ); // If running with unpacked app directory

if ( !fs.existsSync ( ABS_APP_ROOT + '/log' ) )
	fs.mkdirSync ( ABS_APP_ROOT + '/log' );

const LOG_FILE = ABS_APP_ROOT + '/log/%DATE%logfile.txt';//error: 0, warn: 1, info: 2, verbose: 3, debug: 4, silly: 5

global.ERROR   = 0;
global.WARNING = 1;
global.INFO    = 2;

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

const wlog = winston.createLogger({
	format: winston.format.combine (
		winston.format.timestamp(),
		winston.format.printf ( info => `${getFormattedTimestamp(true)} [${info.level.toUpperCase()}] ${info.message}`)
	),
	transports: [
		new winston.transports.DailyRotateFile ({
			filename    : LOG_FILE,
			datePattern : 'YYYY-MM-DD_',
			maxFiles    : '30d'
		})
	]
});

global.log = function ( str, type ) {
	let type_str = "";

	if ( type !== undefined ) {
		switch ( type ) {
			case global.ERROR  : type_str  = "ERROR"  ; break;
			case global.WARNING: type_str  = "WARNING"; break;
			case global.INFO   :
			default: type_str = "INFO";
		}
	} else
		type_str = "INFO";

	var stackError = new Error()
	var lines      = stackError.stack.split(/^\s+at/m)
	var lineNumber = "????"

	for ( let i = 0 ; i < lines.length ; i++ ) {
		if ( lines[i].includes("global.log") ) {
			var elements   = lines[i+1].split(':')
			lineNumber = elements[elements.length-2]
		}
	}

	while ( lineNumber.toString().length < 4 )
		lineNumber = lineNumber + ' '

	switch ( type_str ) {
		case "ERROR":
			console.error ( getFormattedTimestamp(false), "Ln " + lineNumber + " [" + type_str + "] ", str );
			wlog.error    ( str ); // Write to logfile
			break;
		case "WARNING":
			console.warn ( getFormattedTimestamp(false), "Ln " + lineNumber + " [" + type_str + "] ", str );
			wlog.warn    ( str ); // Write to logfile
			break;
		case "INFO":
		default:
			console.log ( getFormattedTimestamp(false), "Ln " + lineNumber + " [" + type_str + "] ", str ); // same like console.info
			wlog.info   ( str ); // Write to logfile
			break;
	}
}