const {app, BrowserWindow, Menu, dialog, ipcMain } = require('electron')
const path = require('path')
const packagejson = require ( './package.json' )
const { productName, version } = require ( './package.json' )
const electronLocalshortcut = require ( 'electron-localshortcut' )
require('@electron/remote/main').initialize()
const Configstore = require ( 'configstore'    )

let config = new Configstore ( packagejson.name, {} )
global.recentProjects = config.get ( 'recentProjects' )

if ( !global.recentProjects )
	global.recentProjects = []

global.PROD = false

let wasChanged = false

if ( 'ELECTRON_IS_PROD' in process.env ) { // If env variable was set manually ...
	// Set global.PROD according the value of ELECTRON_IS_DEV
	global.PROD = Number.parseInt(process.env.ELECTRON_IS_PROD) === 1
} else {
	// Electron Builder cannot set environment variables. Thus in the finally packaged app,
	// we cannot detect the environment based on the env variable (unless someone would set
	// it manually). But we can detect if the app is running in producion mode by checking the
	// isPackaged property of app
	global.PROD = app.isPackaged
}

require ( './logger.js' ) // Stay below the definition of global.PROD. global.PROD is needed to switch the logging path
log ( 'Running in ' + (global.PROD?'production':'development') + ' mode!' )

let mainWindow = undefined

function createWindow () {
  	mainWindow = new BrowserWindow({
    	width: global.PROD?1200:1000,
    	height: global.PROD?800:1000,
    	webPreferences: {
      		preload: path.join(__dirname, 'preload.js'),
      		nodeIntegration: true,
      		contextIsolation: false
    	}
	})

	require("@electron/remote/main").enable(mainWindow.webContents)

	mainWindow.loadFile('index.html')
	mainWindow.setTitle ( productName + " V" + version )
	let wc = mainWindow.webContents
	
	if ( !global.PROD )
		wc.openDevTools()

	var menuJSON = []

	if ( process.platform === 'darwin' )
    	menuJSON.push ({ label: 'App Menu', submenu: [{ role: 'quit'}] })
  
	var projectMenuJSON = { label: 'Project', submenu: [
		{ label: "New               CTRL-N", click () { wc.send ( 'PROJECT_NEW' ) } },
		{ label: "Open             CTRL-O", click () { openProject() } },
		{ label: "Save               CTRL-S", click () { wc.send ( 'PROJECT_SAVE') }	},
		{ label: "Save As", click () { saveAs() } },
		{ type : "separator" },
		{ label: "Export EXCEL", click () { exportExcel() } },
		{ type : "separator" }
	]}

	for ( let recentProject of global.recentProjects )
		projectMenuJSON.submenu.push ( {label: recentProject.path, click () { wc.send ( 'PROJECT_OPEN', recentProject.path ) } })

	projectMenuJSON.submenu.push ({ type : "separator" })
	projectMenuJSON.submenu.push ({ label: "Exit                 CTRL-Q", click () { app.exit() } })

	function saveAs () {
		dialog.showSaveDialog ({
			title: "Save Project As",
			filters: [ {name: "OGP", extensions: ["ogp"]} ]
		}).then ( (res) => {
			if ( !res.canceled )
				wc.send ( 'PROJECT_SAVE_AS', res.filePath )
		})
	}

	function exportExcel () {
		dialog.showSaveDialog ({
			title: "EXCEL Export",
			filters: [ {name: "XLSX", extensions: ["xlsx"]} ]
		}).then ( (res) => {
			if ( !res.canceled )
				wc.send ( 'EXPORT_EXCEL', res.filePath )
		})
	}

	function openProject () {
		dialog.showOpenDialog ({
			title: "Open Project",
			filters: [ {name: "OGP", extensions: ["ogp"]} ]
		}).then ( (res) => {
			if ( !res.canceled )
				wc.send ( 'PROJECT_OPEN', res.filePaths[0] )
		})
	}

	menuJSON.push ( projectMenuJSON )
	Menu.setApplicationMenu ( Menu.buildFromTemplate ( menuJSON ) )

	ipcMain.on ( 'TRIGGER_PROJECT_SAVE_AS', (ev, data) => saveAs ())
	ipcMain.on ( "setWasChanged", ( event, val ) => wasChanged = val )

	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+R', () => checkIfSaved ('RELOAD') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+N', () => wc.send ( 'PROJECT_NEW') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+O', () => openProject () )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+S', () => wc.send ( 'PROJECT_SAVE') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+Q', () => checkIfSaved ('EXIT') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+Alt+Shift+I', () => wc.openDevTools() )
}

function checkIfSaved ( action ) {
	let btnIndex = undefined

	if ( wasChanged ) {
		btnIndex = dialog.showMessageBoxSync ( mainWindow, {
			type: 'question',
			buttons: ['Yes', 'No'],
			defaultId: 1,
			title: ' ',
			message: 'Unsaved data!',
			detail: `You have unsaved changes! Do you really want to ${action==='EXIT'?'quit':'reload'}?`
		})
	}

	if ( !wasChanged || btnIndex === 0 ) {
		wasChanged = false

		switch ( action ) {
			case 'EXIT': app.exit(); break
			case 'RELOAD': mainWindow.reload(); break
			default: log ("ERROR: Unsupported action for checkIfSaved() function!", ERROR )
		}
	}
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})