"use strict";

const datepicker  = require ( 'js-datepicker'  )
const packagejson = require ( './package.json' )
const Configstore = require ( 'configstore'    )
const { ipcRenderer } = require ( 'electron' )
const { app, BrowserWindow, dialog } = require ( '@electron/remote' )
const fs = require ('fs')
const prompt = require('electron-prompt')
const excel  = require ( 'exceljs' )

require ( './logger.js' )

var PROD = require ('@electron/remote').getGlobal('PROD')
var recentProjects = require ('@electron/remote').getGlobal('recentProjects')
let config = new Configstore ( packagejson.name, {} )

var DATA_CELL_PADDING_VERTICAL   = 3
var DATA_CELL_PADDING_HORIZONTAL = 10
var GANTT_LINE_HEIGHT = 18 + DATA_CELL_PADDING_VERTICAL *2 + 1
var GANTT_CELL_WIDTH  = 16
var GANTT_BAR_HANDLE_SIZE = Math.floor ( GANTT_CELL_WIDTH / 3 )

var START_DATE_OBJ = undefined
var END_DATE_OBJ   = undefined

var mouseDownData = undefined

const MOUSE_ACTION_DATA_RESIZE_COLUMN   = 1
const MOUSE_ACTION_GANTT_BAR_DRAG_START = 2
const MOUSE_ACTION_GANTT_BAR_DRAG_BODY  = 3
const MOUSE_ACTION_GANTT_BAR_DRAG_END   = 4

const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
const weekDays = ["S", "M", "T", "W", "T", "F", "S"]

var project  = {}
var pickersStart = []
var pickersEnd   = []
var daysPerProject  = 0
var contextMenu = undefined
var dragIdx = undefined

document.addEventListener ( "DOMContentLoaded", function ( event ) {
    let lastProject = config.get ( 'lastProject' )

    if ( lastProject && lastProject.path )
        openProject ( lastProject )
    else
        createNewProject ( {}, true )

    document.addEventListener ( 'mousemove', function (ev) {
        if ( !mouseDownData )
            return

        switch ( mouseDownData.action ) {
            case MOUSE_ACTION_DATA_RESIZE_COLUMN  : resizeDataColumn( ev         ); break
            case MOUSE_ACTION_GANTT_BAR_DRAG_START: mouseMoveAction ( ev, 'start'); break
            case MOUSE_ACTION_GANTT_BAR_DRAG_END  : mouseMoveAction ( ev, 'end'  ); break
            case MOUSE_ACTION_GANTT_BAR_DRAG_BODY : mouseMoveAction ( ev, 'body' ); break
            default:
        }
    })

    document.addEventListener ( 'mouseup', ev => mouseDownData = undefined )

    function resizeDataColumn ( ev, width ) {
        let diffX = ev.pageX - mouseDownData.x
        let dataTableWidth = document.getElementById('data-table').offsetWidth

        let newWidth = (width?width:(mouseDownData.width + diffX) - DATA_CELL_PADDING_HORIZONTAL*2 )
        mouseDownData.elem.style.width = newWidth + 'px'

        let foundColumn = project.columnData.find ( col => col.attributeName === mouseDownData.elem.id.split('_')[1] )
        foundColumn.width = newWidth

        for ( let idx = 0 ; idx < project.taskData.length ; idx++ ) {
            let elem = document.getElementById ( 'data-cell_' + idx + '_' + mouseDownData.elem.innerText )
            elem.style.width = newWidth + 'px'
        }

        let ganttTableHeaderSvg = document.getElementById('gantt-table-header-svg')
        ganttTableHeaderSvg.style.left = (dataTableWidth - mouseDownData.ganttHeaderOffset).toPx()

        let scrollBarWrapper = document.getElementById ('gantt-table-scrollbar-wrapper')
        scrollBarWrapper.style.left = dataTableWidth + 'px'
        scrollBarWrapper.style.width = 'calc(100% - '+dataTableWidth+'px)'

        let ganttTableWrapper = document.getElementById('gantt-table-wrapper')
        ganttTableWrapper.style.width = 'calc(100% - '+dataTableWidth+'px)'
    }

    function updateChartBar ( idx ) {
        let ganttBar = document.getElementById( 'gantt-bar_'+ idx )
        let daysToGanttBarStart = getOffsetInDays ( START_DATE_OBJ, project.getTask(idx).Start )
        let daysGanttBarLength  = getOffsetInDays ( project.getTask(idx).Start, project.getTask(idx).End )

        if ( !ganttBar ) { // Create new Bar if not existing
            let ganttBar = createSVGRect (
                document.getElementById ( 'gantt-table-svg' ),
                daysToGanttBarStart * GANTT_CELL_WIDTH,
                GANTT_LINE_HEIGHT * (idx + 2) + 1,
                daysGanttBarLength * GANTT_CELL_WIDTH,
                GANTT_LINE_HEIGHT - 1,
                'gantt-cell-active',
                'gantt-bar_' + idx
            )
            ganttBar.setAttribute ( 'rx', '5' )
            ganttBar.setAttribute ( 'ry', '5' )
            ganttBar.addEventListener ( 'mousedown', function (ev) {
                let dataTableWidth = document.getElementById('data-table-wrapper').offsetWidth
                let xStartPix = ganttBar.getBBox().x
                let xEndPix   = xStartPix + ganttBar.getBBox().width

                mouseDownData  = {
                    x    : ev.offsetX,
                    y    : ev.offsetY,
                    elem : ev.target,
                    taskStartDate: project.getTask(idx).Start,
                    taskEndDate  : project.getTask(idx).End,
                    ganttBarStart: xStartPix,
                    ganttBarEnd  : xEndPix,
                    ganttBarWidth: daysGanttBarLength,
                    dataTableWidth,
                    idx
                }

                if ( ev.offsetX >= xStartPix && ev.offsetX <= xStartPix + GANTT_BAR_HANDLE_SIZE )
                    mouseDownData.action = MOUSE_ACTION_GANTT_BAR_DRAG_START
                else if ( ev.offsetX > xStartPix + GANTT_BAR_HANDLE_SIZE && ev.offsetX < xEndPix - GANTT_BAR_HANDLE_SIZE ) {
                    mouseDownData.action = MOUSE_ACTION_GANTT_BAR_DRAG_BODY
                } else if ( ev.offsetX >= xEndPix - GANTT_BAR_HANDLE_SIZE && ev.offsetX <= xEndPix )
                    mouseDownData.action = MOUSE_ACTION_GANTT_BAR_DRAG_END
            })

            // Trigger mouseover of the corresponding gantt line by dispatching a 'mouseover' event to it.
            // That listener will then do the highlighting for both data and gantt table.
            ganttBar.addEventListener ( 'mouseover', function (ev) {
                document.getElementById ( 'gantt-line_' + this.id.lineIndex() ).dispatchEvent ( new Event('mouseover') )
            })

            // Trigger mouseout of the corresponding gantt line by dispatching a 'mouseout' event to it.
            // That listener will then do the un-highlighting for both data and gantt table.
            ganttBar.addEventListener ( 'mouseout', function (ev) {
                document.getElementById ( 'gantt-line_' + this.id.lineIndex() ).dispatchEvent ( new Event('mouseout') )
            })

            // Set mouse cursor icon
            ganttBar.addEventListener ( 'mousemove', function (ev) {
                if ( mouseDownData ) // Don't change icon anymore once user holds donw the left mouse button.
                    return

                let xStartPix = ganttBar.getBBox().x
                let xEndPix   = xStartPix + ganttBar.getBBox().width
        
                if ( ev.offsetX >= xStartPix && ev.offsetX <= xStartPix + GANTT_BAR_HANDLE_SIZE )
                    this.style.cursor = 'col-resize'
                else if ( ev.offsetX > xStartPix + GANTT_BAR_HANDLE_SIZE && ev.offsetX < xEndPix - GANTT_BAR_HANDLE_SIZE )
                    this.style.cursor = 'w-resize'
                else if ( ev.offsetX >= xEndPix - GANTT_BAR_HANDLE_SIZE && ev.offsetX <= xEndPix )
                    this.style.cursor = 'col-resize'
            })
        } else { // Gantt bar already exists => update it
            ganttBar.setAttribute ( "x"     , daysToGanttBarStart * GANTT_CELL_WIDTH )
            ganttBar.setAttribute ( "width" , daysGanttBarLength * GANTT_CELL_WIDTH )

            if ( daysToGanttBarStart < 0 ) {
                getProjectDateBounds ( project.taskData )
                updateGanttTable()
            }
        }
    }

    function mouseMoveAction ( ev, action ) {
        // ev.offsetX is the x-mouse coordinate in relation to the element's relative 0,0 coordinate
        // independent whether the element is scrolled or not
        let diffX = ev.offsetX - mouseDownData.x
        let deltaDays = Math.floor ( diffX / GANTT_CELL_WIDTH )
        let tmpDate = undefined

        switch ( action ) {
            case 'start':
                tmpDate = new Date ( convertDate ( mouseDownData.taskStartDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays )
                
                if ( compareDate ( tmpDate, mouseDownData.taskEndDate ) > 0 )
                    tmpDate = mouseDownData.taskEndDate
                
                project.setTask ( mouseDownData.idx, { Start: convertDate(tmpDate,'string') } )
                log ( `Moving Gantt start in line ${mouseDownData.idx} to: ` + project.getTask(mouseDownData.idx).Start )
                break
            case 'end':
                tmpDate = new Date ( convertDate ( mouseDownData.taskEndDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays +1)
                
                if ( compareDate ( tmpDate, mouseDownData.taskStartDate ) < 0 )
                    tmpDate = mouseDownData.taskStartDate
                
                project.setTask ( mouseDownData.idx, { End: convertDate(tmpDate,'string') } )
                log ( `Moving Gantt end in line ${mouseDownData.idx} to: ` + project.getTask(mouseDownData.idx).End )
                break
            case 'body':
                tmpDate = new Date ( convertDate ( mouseDownData.taskStartDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays )
                project.setTask ( mouseDownData.idx, { Start: convertDate(tmpDate,'string') } )

                tmpDate = new Date ( convertDate ( mouseDownData.taskEndDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays )
                project.setTask ( mouseDownData.idx, { End: convertDate(tmpDate,'string') } )
                log ( `Moving Gantt in line ${mouseDownData.idx} to: ` + project.getTask(mouseDownData.idx).Start + ' ' + project.getTask(mouseDownData.idx).End )
                break
            default:
        }

        if ( compareDate ( project.getTask(mouseDownData.idx).Start, START_DATE_OBJ ) < 1 ||
             compareDate ( project.getTask(mouseDownData.idx).End  , END_DATE_OBJ   ) > 1)
            updateGanttTable()
        else
            updateChartBar ( mouseDownData.idx )
    }
    
    // Prevent blur event from executing if Enter was pressed
    let dontBlur = false

    document.addEventListener ( "keydown", ( ev ) => {
        if ( ev.key === 'Enter' ) {
            dontBlur = true

            if ( ev.target.id.includes ('data-cell_') )
                saveDataCellChanges ( ev )
            if ( ev.target.id.includes ('data-col_') ) {
                let attr = ev.target.id.lineIndex()

                // Do not allow to change the name of certain columns
                if ( attr !== 'Task' && attr !== 'Start' && attr !== 'End' )
                    saveHeaderCellChanges ( ev )
            }
        }
    })

    function saveDataCellChanges ( ev ) {
        ev.preventDefault()
        let elem = ev.target
        let idx = elem.id.lineIndex()
        let taskAttr = elem.id.colIdentifier()
        let taskData = {}
        taskData[taskAttr] = elem.innerText
        project.setTask ( idx, taskData )        
        elem.contentEditable = 'false' // must be a string !!
        elem.scrollLeft = 0
        elem.classList.add ('text-readonly')
    }

    function saveHeaderCellChanges ( ev ) {
        ev.preventDefault()
        let elem = ev.target
        let attr = elem.id.lineIndex()

        let foundElem = project.columnData.find ( el => el.attributeName === attr )
        foundElem.displayName = elem.innerText
        foundElem.attributeName = elem.innerText

        elem.id = 'data-col_' + elem.innerText
        elem.contentEditable = 'false' // must be a string !!
        elem.scrollLeft = 0
        elem.classList.add ('text-readonly')

        // Rename attribute in tasks
        for ( let [idx, task] of project.taskData.entries() ) {
            task[elem.innerText] = task[attr]
            delete task[attr]
        }

        updateDataTable()
    }

    function updateGanttTable () {
        let ganttTableWrapper  = document.getElementById ( 'gantt-table-wrapper' )
        let ganttTableHeaderSvg= document.getElementById ( 'gantt-table-header-svg' )
        let ganttTableSvg      = document.getElementById ( 'gantt-table-svg' )
        let scrollPos = 0

        if ( ganttTableSvg ) {
            scrollPos = ganttTableWrapper.scrollLeft
            ganttTableWrapper.removeChild ( ganttTableSvg )
        }

        if ( ganttTableHeaderSvg )
            ganttTableWrapper.removeChild ( ganttTableHeaderSvg )

        let textTableWidth  = document.getElementById ( 'data-table-wrapper' ).offsetWidth
        let textTableHeight = document.getElementById ( 'data-table' ).offsetHeight

        ganttTableWrapper.style.top    = 0 //'-' + textTableHeight + 'px'
        ganttTableWrapper.style.left   = 0// textTableWidth + 'px'
        ganttTableWrapper.style.width  = 'calc(100% - '+textTableWidth+'px)'
        ganttTableWrapper.style.height = textTableHeight + GANTT_LINE_HEIGHT*2 + 16 + 'px'

        ganttTableSvg = document.createElementNS ( "http://www.w3.org/2000/svg", 'svg' )
        ganttTableSvg.setAttribute ( "width" , "100%" )
        ganttTableSvg.setAttribute ( "height", GANTT_LINE_HEIGHT * (project.taskData.length + 2) )
        ganttTableSvg.setAttribute ( "xmlns" , "http://www.w3.org/2000/svg" )
        ganttTableSvg.setAttribute ( "id"   , "gantt-table-svg" )
        ganttTableWrapper.appendChild ( ganttTableSvg )

        ganttTableHeaderSvg = document.createElementNS ( "http://www.w3.org/2000/svg", 'svg' )
        ganttTableHeaderSvg.setAttribute ( "width" , "100%" )
        ganttTableHeaderSvg.setAttribute ( "height", GANTT_LINE_HEIGHT * 2 + 'px' )
        ganttTableHeaderSvg.setAttribute ( "xmlns" , "http://www.w3.org/2000/svg" )
        ganttTableHeaderSvg.setAttribute ( "id"   , "gantt-table-header-svg" )
        ganttTableHeaderSvg.setAttribute ( "style", "position: fixed; z-index 1;" )
        ganttTableWrapper.appendChild ( ganttTableHeaderSvg )
        ganttTableHeaderSvg.style.top  = document.getElementById ( 'header-area' ).offsetHeight + 'px'
        ganttTableHeaderSvg.style.left = document.getElementById ( 'data-table'  ).offsetWidth + 'px'
        ganttTableWrapper.removeEventListener ( 'scroll', scrollSvgHeader )
        ganttTableWrapper.addEventListener    ( 'scroll', scrollSvgHeader )

        let dateCounterObj  = new Date ( convertDate (START_DATE_OBJ, 'number') )
        let oldDateCounterObj  = {}

        daysPerProject = 0

        // Count days in project
        while ( compareDate ( dateCounterObj, END_DATE_OBJ ) < 0 ) {
            daysPerProject++
            increaseDate ( dateCounterObj )
        }

        daysPerProject++ // make 1-based
        ganttTableSvg.setAttribute       ( "width", GANTT_CELL_WIDTH * daysPerProject )
        ganttTableHeaderSvg.setAttribute ( "width", GANTT_CELL_WIDTH * daysPerProject )
        ganttTableWrapper.scrollLeft = scrollPos
        let oldLeft = ganttTableHeaderSvg.style.left.toVal()
        ganttTableHeaderSvg.style.left = (oldLeft - scrollPos) + 'px'
        document.getElementById('gantt-table-scrollbar').style.width = (GANTT_CELL_WIDTH * daysPerProject) + 'px'

        // Add gantt calendar background lines
        createSVGRect ( ganttTableHeaderSvg, 0, 0, GANTT_CELL_WIDTH * daysPerProject, GANTT_LINE_HEIGHT, 'gantt-line-header' )  // Month - Year
        createSVGRect ( ganttTableHeaderSvg, 0, GANTT_LINE_HEIGHT, GANTT_CELL_WIDTH * daysPerProject, GANTT_LINE_HEIGHT, 'gantt-line-header' ) // Date

        // Add gantt lines ( odd / even background )
        for ( let [ idx, task ] of project.taskData.entries() ) {
            let row = createSVGRect (
                ganttTableSvg,
                0,
                GANTT_LINE_HEIGHT * (idx + 2) + 1,
                GANTT_CELL_WIDTH  * daysPerProject,
                GANTT_LINE_HEIGHT,
                idx%2?'gantt-line-odd gantt-line':'gantt-line-even gantt-line',
                'gantt-line_' + idx
            )

            // Highlight/un-highlight gantt line and also the corresponding line in the data table
            // this is the only place where the highlighting/un-highlighting is done! All other
            // mouseover/mouseout listeners for any table lines just trigger this listener by
            // dispatching a mouseover/mouseout event to it
            row.addEventListener ( 'mouseover', function (ev) {
                this.classList.add ('gantt-line-highlighted')
                document.getElementById ( 'data-line_' + this.id.lineIndex() ).classList.add ('data-line-highlighted')
            })
            row.addEventListener ( 'mouseout', function (ev) {
                this.classList.remove ('gantt-line-highlighted')
                document.getElementById ( 'data-line_' + this.id.lineIndex() ).classList.remove ('data-line-highlighted')
            })
        }

        dateCounterObj = new Date ( START_DATE_OBJ.getTime() ) // reset date counter
        let dayCnt = 0

        // Draw weekend and today column(s) only (see explanations at next while below)
        while ( compareDate ( dateCounterObj, END_DATE_OBJ ) <= 0 ) {
            if ( isWeekend(dateCounterObj) ) {
                createSVGRect ( ganttTableHeaderSvg, GANTT_CELL_WIDTH * dayCnt, GANTT_LINE_HEIGHT + 1, GANTT_CELL_WIDTH, GANTT_LINE_HEIGHT, 'gantt-col-header-weekend' ) // right cell border in gantt grid
                createSVGRect ( ganttTableSvg, GANTT_CELL_WIDTH * dayCnt, GANTT_LINE_HEIGHT * 2 + 1, GANTT_CELL_WIDTH, GANTT_LINE_HEIGHT * (project.taskData.length + 1), 'gantt-col-weekend' ) // right cell border in gantt grid
            }

            if ( isToday(dateCounterObj) ) {
                createSVGRect ( ganttTableHeaderSvg, GANTT_CELL_WIDTH * dayCnt, GANTT_LINE_HEIGHT + 1, GANTT_CELL_WIDTH, GANTT_LINE_HEIGHT, 'gantt-col-header-today' ) // right cell border in gantt grid
                createSVGRect ( ganttTableSvg, GANTT_CELL_WIDTH * dayCnt, GANTT_LINE_HEIGHT * 2 + 1, GANTT_CELL_WIDTH, GANTT_LINE_HEIGHT * (project.taskData.length + 1), 'gantt-col-today' ) // right cell border in gantt grid
            }

            increaseDate ( dateCounterObj )
            dayCnt++
        }

        dayCnt = 0
        dateCounterObj = new Date ( convertDate (START_DATE_OBJ, 'number') ) // reset date counter
        let monthStartPos = 0

        // Add gantt grid lines and calendar text
        while ( compareDate ( dateCounterObj, END_DATE_OBJ ) <= 0 ) { // Not doing this in the while loop above (which seems possible), because SVG overlays objects in the order they are drawn
            // Add date text and grid
            createSVGRect ( ganttTableHeaderSvg, GANTT_CELL_WIDTH * (dayCnt+1), GANTT_LINE_HEIGHT, 1, GANTT_LINE_HEIGHT, 'gantt-line-header-border' ) // right cell border in gantt header
            createSVGRect ( ganttTableSvg, GANTT_CELL_WIDTH * (dayCnt+1), GANTT_LINE_HEIGHT * 2, 1, GANTT_LINE_HEIGHT * (project.taskData.length + 1), 'gantt-line-border' ) // right cell border in gantt grid

            if ( GANTT_CELL_WIDTH >= 16 ) {
                let svgText = createSVGText ( ganttTableHeaderSvg, [dateCounterObj.getDate(), weekDays[dateCounterObj.getDay()]], GANTT_CELL_WIDTH * dayCnt, GANTT_LINE_HEIGHT * 2, 'gantt-line-calendar_' + convertDate ( dateCounterObj, 'string' ) )
                centerSvgObj ( svgText, GANTT_CELL_WIDTH, GANTT_LINE_HEIGHT )
            }

            dayCnt++
            oldDateCounterObj = new Date ( convertDate ( dateCounterObj, 'number') ) // Need to create a new object! Otherwise just getting a reference with =
            increaseDate ( dateCounterObj )

            // Add Month/Year text and grid
            if ( dateCounterObj.getDate() < oldDateCounterObj.getDate() || compareDate ( dateCounterObj, END_DATE_OBJ ) > 0 ) { // check if a new month starts
                let curPos = GANTT_CELL_WIDTH * dayCnt
                let monthWidth = curPos - monthStartPos
                let y = oldDateCounterObj.getFullYear()
                let m = oldDateCounterObj.getMonth()
                let contentStr = ""

                if ( monthWidth < 32 )
                    contentStr = ""
                else if ( monthWidth < 54 )
                    contentStr = monthNamesShort[m]
                else if ( monthWidth < 74 )
                    contentStr = monthNamesShort[m] + ' ' + y.toString().substring (2)
                else if ( monthWidth < 96 )
                    contentStr = monthNamesShort[m] + ' ' + y
                else
                    contentStr = monthNames[m] + ' ' + y

                createSVGRect ( ganttTableHeaderSvg, curPos, 0, 1, GANTT_LINE_HEIGHT, 'gantt-line-header-border' )
                let svgText = createSVGText ( ganttTableHeaderSvg, contentStr, monthStartPos , GANTT_LINE_HEIGHT - 2 )
                //setTimeout ( () => centerSvgObj ( svgText, monthWidth, GANTT_LINE_HEIGHT ), 500 )
                centerSvgObj ( svgText, monthWidth, GANTT_LINE_HEIGHT )
                monthStartPos = curPos
            }
        }

        // Add bottom border for gantt lines
        for ( let [ idx, task ] of project.taskData.entries() ) {
            createSVGRect ( 
                ganttTableSvg,
                0,
                GANTT_LINE_HEIGHT * (idx + 3),
                GANTT_CELL_WIDTH  * daysPerProject,
                1,
                'gantt-line-border',
                'gantt-line-border_' + idx
            )
        }

        // Month - Year bottom border
        createSVGRect ( ganttTableHeaderSvg, 0, GANTT_LINE_HEIGHT, GANTT_CELL_WIDTH * daysPerProject, 1, 'gantt-line-header-border' )

        // Loop through tasks to add gantt chart bars
        for ( let idx = 0 ; idx < project.taskData.length ; idx++ )
            updateChartBar ( idx )
    }

    let scrollBarWrapper = document.getElementById ( 'gantt-table-scrollbar-wrapper' )
    let textTableWidth   = document.getElementById ( 'data-table-wrapper' ).offsetWidth
    scrollBarWrapper.style.width = 'calc(100% - '+textTableWidth+'px)'
    scrollBarWrapper.style.height = '16px'
    scrollBarWrapper.style.left = textTableWidth + 'px'

    let scrollBar = document.getElementById('gantt-table-scrollbar')
    scrollBar.style.width = (GANTT_CELL_WIDTH * daysPerProject) + 'px'
    scrollBar.style.height = '20px'

    // This boolean prevents the scroll listeners for influencing themselves in a loop
    // That means that when scrolling the table by tilting the mouse wheel, the artificial
    // scrollbar shall be positioned correctly, but its own listener must not fire in this case!
    // If it would, it would try to position the table by itself and thus would start
    // a feedback loop between both listeners! 
    let preventScrollListenerLoopEffekt = false

    // Scroll gantt table when artificial scrollbar is scrolled
    let ganttTableWrapper = document.getElementById ('gantt-table-wrapper')
    scrollBarWrapper.addEventListener ( 'scroll', function (ev) {
        if ( preventScrollListenerLoopEffekt ) {
            preventScrollListenerLoopEffekt = false
            return
        }

        ganttTableWrapper.scrollLeft = scrollBarWrapper.scrollLeft
    })

    // Scroll artificial scrollbar when gantt table is scrolled by tilting the mouse wheel
    ganttTableWrapper.addEventListener ( 'scroll', function (ev) {
        preventScrollListenerLoopEffekt = true
        scrollBarWrapper.scrollLeft = ganttTableWrapper.scrollLeft
    })

    document.getElementById('header-area').addEventListener ('click', () => openDonateWindow() )

    function openDonateWindow () {
		let win = new BrowserWindow ({
			webPreferences: {
                nodeIntegration: true,
                contextIsolation: false
            },
			width: 650,
			height: 650
		})
		win.setMenuBarVisibility ( false )
		win.loadURL("file://" + __dirname + "/donate.html")
        //win.webContents.openDevTools()
	}

    function insertLineAboveOrBelow ( idx, isAbove ) {
        log ( `Insert line ${isAbove?'above':'below'} ${idx}` )
        
        if ( !isAbove )
            idx++

        project.taskData.splice ( idx, 0, {} )
        project.setTask ( idx, {
            Task : "",
            Start: convertDate ( new Date (), 'string' ),
            End  : convertDate ( new Date (), 'string' )
        })
    }

    function removeLine ( idx ) {
        log ( "Remove line " + idx )
        project.deleteTask ( idx )
    }

    function insertColBeforeOrAfter ( colName, isBefore ) {
        log ( `Insert column ${isBefore?'before':'after'} '${colName}'` )
        let idx = project.columnData.findIndex ( elem => elem.attributeName === colName )
        prompt({
            label: 'Column Name',
            inputAttrs: { type: 'text', required: true },
            type: 'input',
            alwaysOnTop: true
        })
        .then ( (newColName) => {
            if ( newColName ) {
                let exists = false

                for ( let col of project.columnData ) {
                    if ( col.attributeName === newColName )
                        exists = true
                }

                if ( exists ) {
                    dialog.showMessageBox ( null, {
                        title: ' ',
                        type: 'error',
                        buttons: ['Ok'],
                        defaultId: 0,
                        message: "Error",
                        detail: "A column with this name already exists!"
                    })
                } else {
                    if ( !isBefore )
                        idx++

                    project.columnData.splice ( idx, 0, {
                        displayName: newColName,
                        attributeName: newColName,
                        width: "50"
                    })

                    for ( let task of project.taskData )
                        task[newColName] = ""

                    updateTables()
                }
            }
        })
        .catch ( console.error )
    }

    function removeColumn ( colName, isBefore /*not needed in this case */ ) {
        if ( colName === 'Task' || colName === 'Start' || colName === 'End') { // These columns must not be deleted!
            log ( `Not allowed to remove column ${colName}`)
            return
        }
        log ( "Remove column " + colName )
        let idx = project.columnData.findIndex ( elem => elem.attributeName === colName )
        project.columnData.splice ( idx, 1 )

        for ( let task of project.taskData )
            delete task[colName]

        updateTables()
    }

    function updateTables () {
        updateDataTable()
        updateGanttTable()
    }

    HTMLElement.prototype.addContextMenuOptions4DataTableLines = function ( parentElem ) {
        let options = [{
            icon:  "./icons/insert_line_above.png",
            label: "Insert line above",
            isAbove: true,
            action: insertLineAboveOrBelow
        },{
            icon:  "./icons/insert_line_below.png",
            label: "Insert line below",
            isAbove: false,
            action: insertLineAboveOrBelow
        },{
            icon:  "./icons/remove.png",
            label: "Remove line",
            action: removeLine
        }]

        for ( let option of options ) {
            let elem = document.createElement ( 'div' )
            elem.classList.add ('context-menu-option-line')
            elem.innerHTML = `<span><img src=\"${option.icon}\"/></span><span class=\"context-menu-option-text\">${option.label}</span>`
            elem.addEventListener ( "click", (ev) => {
                option.action ( parentElem.id.lineIndex(), option.isAbove )
                updateTables()
            })
            this.appendChild ( elem )
        }
    }

    HTMLElement.prototype.addContextMenuOptions4DataTableHeader = function ( headerElem ) {
        let options = []

        if ( headerElem.id.lineIndex() !== project.columnData[0].attributeName ) {
            options.push ({
                icon:  "./icons/insert_col_before.png",
                label: "Insert column before",
                isBefore: true,
                action: insertColBeforeOrAfter
            })
        }

        options.push ({
            icon:  "./icons/insert_col_after.png",
            label: "Insert column after",
            isBefore: false,
            action: insertColBeforeOrAfter
        },{
            icon:  "./icons/remove.png",
            label: "Remove column",
            isBefore: undefined,
            action: removeColumn
        })

        for ( let option of options ) {
            let elem = document.createElement ( 'div' )
            elem.classList.add ('context-menu-option-line')
            elem.innerHTML = `<span><img src=\"${option.icon}\"/></span><span class=\"context-menu-option-text\">${option.label}</span>`
            elem.addEventListener ( "click", (ev) => option.action ( headerElem.id.lineIndex(), option.isBefore ) )
            this.appendChild ( elem )
        }
    }

    function addClickListenerText ( elem ) {
        elem.addEventListener ('click', ( ev ) => {
            if ( contextMenu ) {
                contextMenu.remove()
                contextMenu = undefined
            }

            elem.setAttribute ( 'contenteditable', true )
            elem.classList.remove ('text-readonly')
            elem.focus()
        })
    }

    function addContextMenuListener ( elem, elemType ) {
        elem.addEventListener ( 'contextmenu', function (ev) {
            contextMenu = document.getElementById('context-menu')

            if ( contextMenu ) {
                contextMenu.remove()
                contextMenu = undefined
            }

            contextMenu = document.createElement("div")
            contextMenu.classList.add("context-menu")
            contextMenu.id = "context-menu"
            contextMenu.style.visibility ="visible"
            contextMenu.style.left = ev.pageX + "px"
            switch ( elemType ) {
                case 'data-table-line'  : contextMenu.addContextMenuOptions4DataTableLines  ( ev.target.parentElement ); break
                case 'data-table-header': contextMenu.addContextMenuOptions4DataTableHeader ( ev.target ); break
            }
            let body = document.getElementsByTagName('body')[0]
            body.appendChild ( contextMenu ) // need to render first, before doing height calculations
            let bodyHeight = document.getElementsByTagName('body')[0].offsetHeight

            // Make sure context menu is fully visible
            if ( ev.pageY + contextMenu.offsetHeight > bodyHeight )
                contextMenu.style.bottom = 0
            else
                contextMenu.style.top = ev.pageY + 'px'
        })
    }

    function addBlurListener ( elem ) {
        elem.addEventListener ( 'blur', ( ev ) => {
            if ( dontBlur ) {
                dontBlur = false
                return
            }

            if ( ev.target.id.includes ('data-cell_') )
                saveDataCellChanges ( ev )
            if ( ev.target.id.includes ('data-col_') )
                saveHeaderCellChanges ( ev )
        })
    }

    document.addEventListener ('click', (ev) => {
        if ( contextMenu && ev.target.id !== 'context-menu' ) {
            contextMenu.remove()
            contextMenu = undefined
        }
    })

    function updateDataTable () {
        let dataTable = document.getElementById ( 'data-table' )

        // Remove all childs if existing
        while ( dataTable.hasChildNodes() )
            dataTable.removeChild ( dataTable.lastChild )

        // Create header line for column labels
        let row = createAndAppendElement ( 'data-table', 'tr' , {
            class: 'row',
            style: [
                'width: 100%',
                'line-height:' + ((GANTT_LINE_HEIGHT - DATA_CELL_PADDING_VERTICAL) * 2) + 'px',
                'position:fixed',
                'z-index:1'
            ]
        })

        let attrNr = 0

        // Create data table column labels in header line
        for ( let col of project.columnData ) {
            let isLastAttr = attrNr===project.columnData.length-1
            let headerCell = createAndAppendElement ( row, 'th', {
                class: [ 'fixed-col', 'cell' ],
                style: [
                    'min-width:' + col.width + 'px',
                    'width:' + col.width + 'px',
                    'font-weight:bold',
                    'background-color:#cdcdcd',
                    'padding:' + DATA_CELL_PADDING_VERTICAL + 'px ' + DATA_CELL_PADDING_HORIZONTAL + 'px',
                    isLastAttr?'border-right:1px solid #666':'border-right:1px solid #b8b8b8',
                    'position:relative'
                ],
                id: "data-col_" + col.attributeName,
                content: col.displayName
            })

            addContextMenuListener ( headerCell, 'data-table-header' )
            
            // Do not allow to rename certain columns
            if ( col.attributeName !== 'Task' && col.attributeName !== 'Start' && col.attributeName !== 'End' ) {
                addClickListenerText ( headerCell )
                addBlurListener ( headerCell )
            }

            let resizeHandle = createResizeHandle ( headerCell, 'data-resize-handle_' + attrNr, 'RIGHT' )
            resizeHandle.addEventListener ( 'mousedown', (ev) => {
                let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
                let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
                mouseDownData  = {
                    action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
                    x     : ev.pageX,
                    elem  : ev.target.parentElement,
                    width : ev.target.parentElement.offsetWidth,
                    ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
                }
            })

            resizeHandle.addEventListener ( 'dblclick', (ev) => {
                let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
                let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
                mouseDownData = {
                    action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
                    x     : ev.offsetX,
                    elem  : ev.target.parentElement,
                    ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
                }

                resizeDataColumn ( ev, 50 )
                document.dispatchEvent ( new Event ('mousemove') ) // Strange! Need to trigger mousemove manually, otherwise SVG header is not redrawn
                mouseDownData = undefined
            })
            attrNr++
        }

        if ( !project.taskData.length )
            return

        let dataTableBody = document.getElementById ( 'data-table-body' )
        
        if ( dataTableBody )
            document.getElementById('data-table-body').remove()
    
        // Create data table body
        let tbody = createAndAppendElement ( 'data-table', 'tbody', {
            style: [
                'top:' + (GANTT_LINE_HEIGHT * 2 + 1) + 'px',
                'position:relative'
            ],
            id: 'data-table-body'
        })

        // Add tasks to data table
        for ( let [ idx, task ] of project.taskData.entries() ) {
            // Add data table rows
            let row = createAndAppendElement ( tbody, 'tr' , {
                class: [
                    'row',
                    idx % 2?'data-line-odd':'data-line-even'
                ],
                style: ['width: 100%'],
                id: "data-line_" + idx
            })

            row.setAttribute ("draggable", "true")
            // Trigger mouseover of the corresponding gantt line by dispatching a 'mouseover' event to it.
            // That listener will then do the highlighting for both data and gantt table.
            row.addEventListener ( 'mouseover', (ev) => {
                // "ev.currentTarget" contains the element (in our case <tr>) to which the handler was attached originally.
                // It is only available in this listener context!
                // "ev.target" contains the child element <td> which received the click and thus NOT the element to which
                // the handler was attached to!
                let ganttLine = document.getElementById ( 'gantt-line_' + ev.currentTarget.id.lineIndex() )
                ganttLine.dispatchEvent ( new Event ('mouseover') )
            })

            // Trigger mouseout of the corresponding gantt line by dispatching a 'mouseout' event to it.
            // That listener will then do the un-highlighting for both data and gantt table.
            row.addEventListener ( 'mouseout' , (ev) => {
                // "ev.currentTarget" contains the element (in our case <tr>) to which the handler was attached originally.
                // It is only available in this listener context!
                // "ev.target" contains the child element <td> which received the click and thus NOT the element to which
                // the handler was attached to!
                let ganttLine = document.getElementById ( 'gantt-line_' + ev.currentTarget.id.lineIndex() )
                ganttLine.dispatchEvent ( new Event ('mouseout') )
            })

            addDragListener ( row )
            let attrNr = 0

            // Add data table cells
            for ( let colData of project.columnData ) {
                let taskAttr = colData.attributeName
                let foundColumn = project.columnData.find ( col => col.attributeName === taskAttr )
                let isLastAttr = attrNr===project.columnData.length-1
                let cell = createAndAppendElement ( row, 'td', {
                    class: [
                        'text-overflow',
                        'fixed-col',
                        'cell',
                        'text-readonly'
                    ],
                    style: [
                        'min-width:' + foundColumn.width + 'px',
                        'width:' + foundColumn.width + 'px',
                        isLastAttr?'border-right:1px solid #666':'border-right:1px solid #b8b8b8',
                        'padding:' + DATA_CELL_PADDING_VERTICAL + 'px ' + DATA_CELL_PADDING_HORIZONTAL + 'px',
                        'min-height:' + (GANTT_LINE_HEIGHT - DATA_CELL_PADDING_VERTICAL*2 - 1) + 'px'
                    ],
                    id: 'data-cell_' + idx + '_' + foundColumn.attributeName,
                    content: task[taskAttr]
                })

                addBlurListener ( cell )
                addContextMenuListener ( cell, 'data-table-line' )
                addDropListener ( cell )

                if ( taskAttr === 'Start' || taskAttr === 'End' ) {
                    cell.style.textAlign = "center"
                    let picker = datepicker ( cell, {
                        onSelect: ( pickerInstance, date ) => {
                            let idx = pickerInstance.el.id.lineIndex()
                            let taskAttrFromId = pickerInstance.el.id.colIdentifier()
                            let taskOptions = {}
                            taskOptions[taskAttrFromId] = convertDate ( date, 'string' )
                            project.setTask ( idx, taskOptions )

                            if ( taskAttrFromId === 'Start' && compareDate ( date, START_DATE_OBJ ) < 0 ) {
                                getProjectDateBounds ( project.taskData )
                                updateGanttTable()
                            }

                            if ( taskAttrFromId === 'End'   && compareDate ( date, END_DATE_OBJ   ) > 0 ) {
                                getProjectDateBounds ( project.taskData )
                                updateGanttTable()
                            }
                        },
                        dateSelected: convertDate ( project.getTask(idx)[taskAttr], 'object' ),
                        startDay: 1,
                        showAllDates: true
                    })
                    picker.calendarContainer.style.setProperty ( 'font-family', 'OpenSans' )
                    picker.calendarContainer.style.setProperty ( 'font-size', '12px' )

                    if ( taskAttr === 'Start' ) pickersStart.push ( picker )
                    if ( taskAttr === 'End'   ) pickersEnd.push   ( picker )
                } else {
                    addClickListenerText ( cell )
                }

                attrNr++
            }
        }
    }

    // Extend height to make sure that last line is fully visible when table was scrolled down
    document.getElementById ('table-area').style.height = 'calc(100% - ' + 23 + 'px)'

    function addDragListener ( elem ) {
        elem.addEventListener ( "dragstart" , ev => dragIdx = ev.currentTarget.id.lineIndex() )
    }

    function addDropListener ( elem ) {
        // Canceling dragover is required! Otherwise drop won't fire
        // https://stackoverflow.com/questions/32084053/why-is-ondrop-not-working
        elem.addEventListener ( "dragover" , ev => ev.preventDefault() )

        elem.addEventListener ( "drop" , function (ev) {
            let idx = ev.currentTarget.id.lineIndex()

            if ( idx === dragIdx )
                return

            log ( `Moving line ${dragIdx} to line ${idx} ...`)
            let taskData = {}

            for ( let attr in project.getTask(dragIdx) ) {
                log ( project.getTask(dragIdx)[attr])
                taskData[attr] = project.getTask(dragIdx)[attr]
            }

            removeLine ( dragIdx )
            insertLineAboveOrBelow ( idx, true )
            project.setTask ( idx, taskData )
            updateTables()
        })
    }

    ipcRenderer.on ( 'PROJECT_NEW'    , (ev, mesg) => createNewProject( {}, true ) )
    ipcRenderer.on ( 'PROJECT_OPEN'   , (ev, path) => openProject ({path: path, name: "Unnamed"}) )
    ipcRenderer.on ( 'PROJECT_SAVE'   , (ev, mesg) => {
        if ( !project.path )
            ipcRenderer.send ( 'TRIGGER_PROJECT_SAVE_AS' )
        else
            saveProject ()
    })
    ipcRenderer.on ( 'PROJECT_SAVE_AS', (ev, path) => {
        project.path = path
        config.set ( 'lastProject.path', path )
        saveProject ()
    })
    ipcRenderer.on ( 'EXPORT_EXCEL', (ev, path) => {
        exportExcel ( path )
    })

    function openProject ( prj ) {
        if ( prj === undefined ) {
            log ( "Unable to open project! Project data is empty!", ERROR )
            return
        }

        createNewProject ({name: prj.name, path: prj.path})

        if ( prj.path ) {
            try {
                let loadedData = fs.readFileSync ( prj.path )
                let jsonData = JSON.parse ( loadedData )

                for ( let data in jsonData )
                    project[data] = jsonData[data]

                GANTT_CELL_WIDTH = parseInt(project.ganttCellWidth)
            } catch ( err ) {
                log ( err )
            }
        }

        getProjectDateBounds ( project.taskData )
        updateTables ()

        config.set ( 'lastProject.path', prj.path )
        addToRecentProjects ( recentProjects, prj.path )
    }

    function saveProject () {
        // Blur all lines (header and data cells) to make sure that data is saved
        for ( let col of project.columnData )
            document.getElementById ( 'data-col_' + col.attributeName ).blur()

        for ( let idx = 0 ; idx < project.taskData.length ; idx++ ) {
            for ( let attr in project.taskData[idx] )
            document.getElementById('data-cell_'+idx+'_'+attr).blur()
        }

        let dataToSave = {}
        dataToSave.name       = project.name
        dataToSave.columnData = project.columnData
        dataToSave.taskData   = project.taskData
        dataToSave.ganttCellWidth = GANTT_CELL_WIDTH.toString()
        let serialized = JSON.stringify ( dataToSave, null, "    " )
        fs.writeFileSync ( project.path, serialized )
        addToRecentProjects ( recentProjects, project.path )
    }

    function exportExcel ( path ) {
        let workbook = new excel.Workbook()
        let worksheet  = workbook.addWorksheet ( 'Gantt Data' )
        let excelColumns = []

        for ( let col of project.columnData ) {
            excelColumns.push ({
                header: col.displayName  ,
                key   : col.attributeName,
                width : 20
            })
        }

        worksheet.columns = excelColumns
        worksheet.getRow(1).font = { name: 'Arial', family: 4, size: 11, bold: true }
		worksheet.views = [{
			state    : 'frozen',
			xSplit   : 0,
			ySplit   : 1,
			zoomScale: 85
		}]

        for ( let task of project.taskData ) {
            let row = {}

            for ( let col of project.columnData )
                row[col.attributeName] = task[col.attributeName]

            worksheet.addRow ( row )
        }

        fs.open ( path, 'w+', function ( err, fd ) {
			if ( err ) {
				log ( `Unable to open ${path}! Error: ` + err, ERROR );

				if ( err.code === 'EBUSY' )
					log ( "ERROR: File busy!" )
				else
                    log ( "ERROR: " + err )

				if ( fd )
					fs.close ( fd )

				return
			}

			workbook.xlsx.writeFile ( path ).then ( function () {
                log ( "EXCEL written successfully!" )
				fs.close ( fd )
			})
		})
    }

    function createNewProject ( options, shallUpdate ) {
        GANTT_CELL_WIDTH = 16 // set back to default value

        project = {
            getTask: (idx) => project.taskData[idx],
            setTask: (idx, newData) => {
                if ( newData ) {
                    for ( let attr in newData ) {
                        switch ( attr ) {
                            case 'Start':
                            case 'End':
                                project.taskData[idx][attr] = newData[attr]
                                break
                
                            default:
                                project.taskData[idx][attr] = newData[attr]
                        }
                    }

                    if ( compareDate ( project.taskData[idx].End, project.taskData[idx].Start ) < 0 ) {
                        project.taskData[idx].End = project.taskData[idx].Start
                        document.getElementById ( 'data-cell_' + idx + '_End' ).innerHTML = project.taskData[idx].End
                        pickersEnd[idx].setDate ( convertDate (project.taskData[idx].End, 'object') , true)
                    }

                    updateTables ( idx )
                }
            },
            deleteTask : (idx) => project.taskData.splice ( idx, 1 ),
            name: '',
            taskData: [{
                Task : "",
                Start: convertDate ( new Date (), 'string' ),
                End  : convertDate ( new Date (), 'string' )
            }]
        }

        START_DATE_OBJ = new Date ()
        END_DATE_OBJ   = new Date ()
        decreaseDate ( START_DATE_OBJ, 15 )
        increaseDate ( END_DATE_OBJ  , 15 )

        project.columnData = [{
            "attributeName": "Task",
            "displayName": "Task",
            "width":"200"
        },{
            "attributeName": "Start",
            "displayName": "Start",
            "width":"50"
        },{
            "attributeName": "End",
            "displayName": "End",
            "width":"50"
        }]

        if ( options ) {
            if ( options.name )
                project.name = options.name

            if ( options.path )
                project.path = options.path

            if ( options.columnData ) {
                for ( let col of options.columnData ) {
                    let foundColumn = project.columnData.find ( elem => elem.attributeName === col.attributeName )

                    if ( foundColumn ) {
                        for ( let attr in col )
                            foundColumn[attr] = col[attr]
                    } else {
                        let newColumn = {}

                        for ( let attr in col )
                            newColumn[attr] = col[attr]

                        project.columnData.push ( newColumn )
                    }
                }
            }
        }

        if ( shallUpdate )
            updateTables ()
    }

    document.addEventListener ( "wheel", (ev) => {
        if ( ev.ctrlKey ) {
            if ( ev.deltaY < 0 )
                GANTT_CELL_WIDTH += GANTT_CELL_WIDTH
            else
                GANTT_CELL_WIDTH /= 2 
            
            if ( GANTT_CELL_WIDTH <  2 ) GANTT_CELL_WIDTH = 2
            if ( GANTT_CELL_WIDTH > 64 ) GANTT_CELL_WIDTH = 64

            updateGanttTable()
        }
    })
})

// Keep outside of DOMContentLoaded listener to avoid issues with function being undefined!
var scrollSvgHeader = function (ev) {
    let val = (document.getElementById ( 'data-table-wrapper' ).offsetWidth - this.scrollLeft) + 'px'
    document.getElementById ( 'gantt-table-header-svg').style.left = val
}