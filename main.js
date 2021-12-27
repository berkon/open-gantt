const {app, BrowserWindow, Menu, dialog, ipcMain } = require('electron')
const path = require('path')
const packagejson = require ( './package.json' )
const { productName, version } = require ( './package.json' )
const electronLocalshortcut = require ( 'electron-localshortcut' )
require ( './logger.js' )
require('@electron/remote/main').initialize()
const Configstore = require ( 'configstore'    )

let config = new Configstore ( packagejson.name, {} )
global.recentProjects = config.get ( 'recentProjects' )

if ( !global.recentProjects )
	global.recentProjects = []

global.PROD = false

function createWindow () {
  	const mainWindow = new BrowserWindow({
    	width: global.PROD?1200:800,
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

	ipcMain.on ( 'TRIGGER_PROJECT_SAVE_AS', (ev, data) => {
		saveAs ()
	})

	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+R', () => mainWindow.reload() )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+N', () => wc.send ( 'PROJECT_NEW') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+O', () => openProject () )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+S', () => wc.send ( 'PROJECT_SAVE') )
	electronLocalshortcut.register ( mainWindow, 'CommandOrControl+Q', () => app.exit() )
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