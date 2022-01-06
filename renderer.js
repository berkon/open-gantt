"use strict";

const datepicker  = require ( 'js-datepicker'  )
const packagejson = require ( './package.json' )
const Configstore = require ( 'configstore'    )
const { ipcRenderer } = require ( 'electron' )
const { app, BrowserWindow, dialog } = require ( '@electron/remote' )
const fs = require ('fs')
const prompt = require('electron-prompt')
const excel  = require ( 'exceljs' )
const eventListeners = require ( './event-listeners.js')

var PROD = require ('@electron/remote').getGlobal('PROD')
require ( './logger.js' ) // Stay below the definition of PROD. PROD is needed to switch the logging path

var recentProjects = require ('@electron/remote').getGlobal('recentProjects')
let config = new Configstore ( packagejson.name, {} )

var DATA_CELL_PADDING_VERTICAL   = 3
var DATA_CELL_PADDING_HORIZONTAL = 10
var GANTT_LINE_HEIGHT = 18 + DATA_CELL_PADDING_VERTICAL *2 + 1
var GANTT_CELL_WIDTH  = 16
var GANTT_BAR_HANDLE_SIZE = Math.floor ( GANTT_CELL_WIDTH / 3 )

const DRAG_TYPE_COL = 1
const DRAG_TYPE_ROW = 2

var START_DATE_OBJ = undefined
var END_DATE_OBJ   = undefined

var mouseDownData = undefined

const MOUSE_ACTION_DATA_RESIZE_COLUMN   = 1
const MOUSE_ACTION_GANTT_BAR_DRAG_START = 2
const MOUSE_ACTION_GANTT_BAR_DRAG_BODY  = 3
const MOUSE_ACTION_GANTT_BAR_DRAG_END   = 4

const GROUP_INDENT_SIZE = 18

const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
const weekDays = ["S", "M", "T", "W", "T", "F", "S"]

var project  = {}
var daysPerProject  = 0
var contextMenu = undefined
var dragIdx = undefined
var dragColAttr = undefined
var regexMatch = true
var picker = undefined
var queriedElements = undefined
var dragType = undefined


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

    async function resizeDataColumn ( ev, width ) {
        if ( width )
            width = parseInt ( width )

        let diffX = ev.pageX - mouseDownData.x
        let newColWidth = (width?width:(mouseDownData.width + diffX) - DATA_CELL_PADDING_HORIZONTAL*2 )
        mouseDownData.elem.style.width = newColWidth + 'px'

        let foundColumn = project.columnData.find ( col => col.attributeName === mouseDownData.elem.id.split('_')[1] )
        foundColumn.width = newColWidth.toString()
    
        // "display = none" in the line below, removes the table body from the DOM render queue. After doing all recalculations
        // it will be enabled further down again. This greatly improves UI performance with huge tables.
        document.getElementById('data-table-body').style.display = 'none'
        
        let newColWidthPx = newColWidth.toPx()

        for ( let elem of queriedElements )
            elem.style.width = newColWidthPx

        // As described above, here the table's body is inserted back into the DOM's render queue
        document.getElementById('data-table-body').style.display= 'table-row-group'

        let dataTableWidth = document.getElementById('data-table').offsetWidth
        let ganttTableHeaderSvg = document.getElementById('gantt-table-header-svg')
        ganttTableHeaderSvg.style.left = (dataTableWidth - mouseDownData.ganttHeaderOffset).toPx()

        let scrollBarWrapper = document.getElementById ('gantt-table-scrollbar-wrapper')
        scrollBarWrapper.style.left = dataTableWidth + 'px'
        scrollBarWrapper.style.width = 'calc(100% - '+dataTableWidth+'px)'

        let ganttTableWrapper = document.getElementById('gantt-table-wrapper')
        ganttTableWrapper.style.width = 'calc(100% - '+dataTableWidth+'px)'

        ipcRenderer.send( "setWasChanged", true )
    }

    function updateChartBar ( idx ) {
        let ganttBar = document.getElementById( 'gantt-bar_'+ idx )
        let daysToGanttBarStart = getOffsetInDays ( START_DATE_OBJ, project.getTask(idx).Start )
        let daysGanttBarLength  = getLengthInDays ( project.getTask(idx).Start, project.getTask(idx).End )
        let isGroup = project.taskData[idx].isGroup

        if ( !ganttBar ) { // Create new Bar if not existing
            let ganttBar = createSVGRect (
                document.getElementById ( 'gantt-table-svg' ),
                daysToGanttBarStart * GANTT_CELL_WIDTH,
                GANTT_LINE_HEIGHT * (idx + 2) + 1,
                daysGanttBarLength * GANTT_CELL_WIDTH,
                GANTT_LINE_HEIGHT - 1,
                isGroup?'gantt-bar-group':'gantt-bar-active',
                'gantt-bar_' + idx
            )
            ganttBar.setAttribute ( 'rx', '5' )
            ganttBar.setAttribute ( 'ry', '5' )

            if ( !isGroup ) {
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
            }
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

                project.setTask ( mouseDownData.idx, { Start: convertDate(tmpDate,'string') })
                log ( 'Moving Gantt bar start in line:', mouseDownData.idx, 'to:', project.getTask(mouseDownData.idx).Start )
                break
            case 'end':
                tmpDate = new Date ( convertDate ( mouseDownData.taskEndDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays +1)

                if ( compareDate ( tmpDate, mouseDownData.taskStartDate ) < 0 )
                    tmpDate = mouseDownData.taskStartDate

                project.setTask ( mouseDownData.idx, { End: convertDate(tmpDate,'string') })
                log ( 'Moving Gantt bar end in line:', mouseDownData.idx, 'to:', project.getTask(mouseDownData.idx).End )
                break
            case 'body':
                tmpDate = new Date ( convertDate ( mouseDownData.taskStartDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays )
                project.setTask ( mouseDownData.idx, { Start: convertDate(tmpDate,'string') })

                tmpDate = new Date ( convertDate ( mouseDownData.taskEndDate, 'number' ) )
                increaseDate ( tmpDate, deltaDays )
                project.setTask ( mouseDownData.idx, { End: convertDate(tmpDate,'string') })
                log ( 'Moving Gantt bar in line:', mouseDownData.idx, 'to new Start:', project.getTask(mouseDownData.idx).Start, 'and new End:', project.getTask(mouseDownData.idx).End )
                break
            default:
        }

        if ( compareDate ( project.getTask(mouseDownData.idx).Start, START_DATE_OBJ ) < 1 ||
             compareDate ( project.getTask(mouseDownData.idx).End  , END_DATE_OBJ   ) > 1)
            updateGanttTable()
        else {
            updateDataTable ( mouseDownData.idx )
            updateChartBar ( mouseDownData.idx )
        }
    }
    
    document.addEventListener ( "keyup", ( ev ) => {
        if ( ev.target.id.colIdentifier() === 'Start' || ev.target.id.colIdentifier() === 'End' ) {
            if ( /^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)(?:0?[13-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$/.test (ev.target.innerText) ) {
                ev.target.style.color = null
                regexMatch = true
            } else {
                ev.target.style.color = "red"
                regexMatch = false
            }
        }
    })

    document.addEventListener ( "keydown", ( ev ) => {
        switch ( ev.key ) {
            case 'Enter':
                log ( "EVENT: 'keydown'  <Enter> was pressed!")
                ev.target.contentEditable = 'false' // (must be a string) this trigger the blur event where all the stuff is handeled !!
                break

            case 'Escape':
                log ( "EVENT: 'keydown'  <Escape> was pressed!")
                if ( picker ) {
                    picker.remove()
                    picker = undefined
                }
                break

            case 'Tab':
                log ( "EVENT: 'keydown'  <Tab> was pressed!")
                let curLineIdx = undefined
                let curAttr    = undefined
                let curColIdx  = undefined
                let curTarget  = undefined

                let newLineIdx = undefined
                let newAttr    = undefined
                let newColIdx  = undefined
                let newTarget  = undefined

                curTarget  = ev.target
                curLineIdx = ev.target.id.lineIndex()
                curAttr    = ev.target.id.colIdentifier()
                curColIdx  = project.columnData.findIndex ( elem => elem.attributeName === curAttr )
                curTarget  = ev.target

                // Save changes of current text cell (no need to save data of date picker cell)
                if ( curAttr ) {
                    if ( curAttr !== 'Start' && curAttr !== 'End' )
                        saveDataCellChanges ( ev )
                    else
                        curTarget.blur()
                }

                newLineIdx = curLineIdx

                if ( !ev.shiftKey ) { // Go forward
                    newColIdx = curColIdx + 1

                    if ( newColIdx < project.columnData.length ) {
                        newAttr = project.columnData[newColIdx].attributeName
                    } else {
                        newAttr = project.columnData[1].attributeName
                        newLineIdx++

                        if ( newLineIdx === project.taskData.length )
                            insertLineAboveOrBelow ( curLineIdx, false )
                    }
                } else { // Go backwards
                    newColIdx = curColIdx - 1

                    if ( newColIdx >= 1 ) { // Not 0 because we want to exclude the '#' index column
                        newAttr = project.columnData[newColIdx].attributeName
                    } else {
                        newAttr = project.columnData[project.columnData.length-1].attributeName
                        newLineIdx--

                        if ( newLineIdx < 0 ) {
                            insertLineAboveOrBelow ( curLineIdx, true )
                            newLineIdx = 0
                        }
                    }
                }

                newTarget = document.getElementById ( 'data-cell_' + newLineIdx + '_' + newAttr )
                newTarget.setAttribute ( 'contenteditable', true )
                newTarget.classList.remove ('text-readonly')
                break
        }
    })

    function saveDataCellChanges ( ev ) {
        let elem = ev.target
        let idx = elem.id.lineIndex()
        let taskAttr = elem.id.colIdentifier()
        let taskData = {}

        if ( taskData[taskAttr] !== elem.innerText )
            ipcRenderer.send( "setWasChanged", true )

        taskData[taskAttr] = elem.innerText
        project.setTask ( idx, taskData, true )
        elem.contentEditable = 'false' // must be a string !!
        elem.scrollLeft = 0
        elem.classList.add ('text-readonly')
    }

    function saveHeaderCellChanges ( ev ) {
        let elem = ev.target
        let attr = elem.id.lineIndex()

        let foundElem = project.columnData.find ( el => el.attributeName === attr )

        if ( foundElem.displayName !== elem.innerText )
            ipcRenderer.send( "setWasChanged", true )

        foundElem.displayName   = elem.innerText
        foundElem.attributeName = elem.innerText

        elem.id = 'data-col_' + elem.innerText
        elem.contentEditable = 'false' // must be a string !!
        elem.scrollLeft = 0
        elem.classList.add ('text-readonly')

        // Rename attribute in tasks
        for ( let task of project.taskData ) {
            let data = task[attr]
            delete task[attr]
            task[elem.innerText] = data
        }

        updateDataTable()
    }

    function updateGanttTable ( mouseX, absDaysToMousePos ) {
        let ganttTableWrapper  = document.getElementById ( 'gantt-table-wrapper' )
        let ganttTableHeaderSvg= document.getElementById ( 'gantt-table-header-svg' )
        let ganttTableSvg      = document.getElementById ( 'gantt-table-svg' )
        let scrollPos = 0

        if ( ganttTableSvg ) {
            let ganttTableWrapper = document.getElementById ( 'gantt-table-wrapper' )
            let dataTableWidth    = document.getElementById ( 'data-table-wrapper' ).offsetWidth
            let delta             = mouseX - dataTableWidth            
            scrollPos             = (GANTT_CELL_WIDTH * absDaysToMousePos) - delta
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

        let oldDateCounterObj  = {}

        daysPerProject = getLengthInDays ( START_DATE_OBJ, END_DATE_OBJ ) // make 1-based
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

        let dateCounterObj = new Date ( START_DATE_OBJ.getTime() )
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

    document.getElementById('donate-button').addEventListener ('click', () => openDonateWindow() )

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

    function insertLineAboveOrBelow ( idx, isAbove, redraw ) {
        log ( 'Insert line', isAbove?'above':'below',  idx, '=> becomes', isAbove?idx:idx+1 )
        
        if ( !isAbove )
            idx++

        project.taskData.splice ( idx, 0, {} )
        project.setTask ( idx, {
            Task : "",
            Start: convertDate ( new Date (), 'string' ),
            End  : convertDate ( new Date (), 'string' ),
            isGroup: false,
            groupLevel: 0
        }, false )

        ipcRenderer.send( "setWasChanged", true )

        if ( redraw )
            updateTables()
    }

    function removeLine ( idx ) {
        log ( "Removing line", idx )
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
                        minWidth: "50",
                        width: "50"
                    })

                    for ( let task of project.taskData )
                        task[newColName] = ""

                    updateTables()
                }
            }

            ipcRenderer.send( "setWasChanged", true )
        })
        .catch ( console.error )
    }

    function removeColumn ( colName, isBefore /*not needed in this case */ ) {
        if ( colName === 'Task' || colName === 'Start' || colName === 'End') { // These columns must not be deleted!
            log ( `Not allowed to remove column: '${colName}'`)
            return
        }

        log ( `Removing column: '${colName}'` )
        let idx = project.columnData.findIndex ( elem => elem.attributeName === colName )
        project.columnData.splice ( idx, 1 )

        for ( let task of project.taskData )
            delete task[colName]

        ipcRenderer.send( "setWasChanged", true )
        updateTables()
    }

    function updateTables ( idx ) {
        if ( idx !== undefined ) {
            updateDataTable ( idx )
            updateChartBar ( idx )
        } else {
            updateDataTable ()
            updateGanttTable()
        }
    }

    function makeGroup ( idx ) {
        project.taskData[idx].isGroup = true
    }

    function unGroup ( idx ) {
        let idxCnt = idx

        if ( !project.taskData[idx].isGroup )
            return

        do {
            if ( project.taskData[idx].isGroup ) {
                project.taskData[idx].isGroup = false

                if ( idxCnt + 1 >= project.taskData.length || project.taskData[idxCnt + 1].groupLevel <= project.taskData[idxCnt] )
                    return
                else
                    idxCnt++
            }

            project.taskData[idxCnt].groupLevel--
            idxCnt++
        } while ( idxCnt < project.taskData.length && project.taskData[idxCnt].groupLevel > project.taskData[idxCnt-1].groupLevel )
    }

    HTMLElement.prototype.addContextMenuOptions4DataTableLines = function ( parentElem ) {
        let options = [{
            icon:  "./icons/insert_line_above.png",
            label: "Insert line above",
            isAbove: true,
            action: insertLineAboveOrBelow
        },{
            icon:  "./icons/remove.png",
            label: "Remove line",
            action: removeLine
        },{
            icon:  "./icons/group.png",
            label: "Make group",
            action: makeGroup
        },{
            icon:  "./icons/un_group.png",
            label: "Un-group",
            action: unGroup
        },{
            icon:  "./icons/insert_line_below.png",
            label: "Insert line below",
            isAbove: false,
            action: insertLineAboveOrBelow
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

        // No need to exclude '#' (idx 0) because context menu is not added at all to this colum header

        if ( headerElem.id.lineIndex() !== project.columnData[1].attributeName) { // Task name column
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
        })

        if ( headerElem.id.lineIndex() !== project.columnData[0].attributeName && headerElem.id.lineIndex() !== project.columnData[1].attributeName) {
            options.push ({
                icon:  "./icons/remove.png",
                label: "Remove column",
                isBefore: undefined,
                action: removeColumn
            })
        }

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
            log ( "EVENT: 'click'   TARGET: " + ev.target.id + "   CURRENT-TARGET: " + ev.currentTarget.id )

            if ( contextMenu ) {
                contextMenu.remove()
                contextMenu = undefined
            }

            let idx = ev.target.id.lineIndex()
            let taskAttr = ev.target.id.colIdentifier()

            if ( !regexMatch || (project.taskData[idx].isGroup && (taskAttr === 'Start' || taskAttr === 'End') ) )
                return

            elem.setAttribute ( 'contenteditable', true )
            elem.classList.remove ('text-readonly')
            log ( "FOCUS: " + elem.id )
            elem.focus()

            if ( taskAttr === 'Start' || taskAttr === 'End' ) {
                elem.style.textAlign = "center"

                if ( picker )
                    picker.remove()

                picker = datepicker ( elem, {
                    onSelect: ( pickerInstance, date ) => {
                        let idx = pickerInstance.el.id.lineIndex()
                        let taskAttrFromId = pickerInstance.el.id.colIdentifier()
                        let taskOptions = {}
                        taskOptions[taskAttrFromId] = convertDate ( date, 'string' )

                        // Move End date accordingly
                        if ( taskAttrFromId === 'Start' ) {
                            let task = project.getTask ( idx )
                            let days = getOffsetInDays ( task.Start, task.End )
                            let newEndDate = new Date ( date.getTime() )
                            newEndDate = increaseDate ( newEndDate, days )
                            taskOptions.End = convertDate ( newEndDate, 'string' )
                        }

                        project.setTask ( idx, taskOptions, true )
                        picker.remove()
                        picker = undefined
                        ipcRenderer.send( "setWasChanged", true )

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
                picker.setDate ( convertDate (elem.innerText, 'object') , true )
                picker.show()
            }
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
                case 'data-table-header': contextMenu.addContextMenuOptions4DataTableHeader ( ev.target ); break
                case 'data-table-line'  : contextMenu.addContextMenuOptions4DataTableLines  ( ev.target.parentElement ); break
            }

            if ( !contextMenu.hasChildNodes() )
                return

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
            log ( "EVENT: 'blur'    TARGET: " + ev.target.id + "   CURRENT-TARGET: " + ev.currentTarget.id )

            if ( !regexMatch )
                return

            if ( ev.target.id.includes ('data-cell_') ) {
                let attr = ev.target.id.colIdentifier()
                let task = project.getTask ( ev.target.id.lineIndex() )
                let days = getOffsetInDays ( task.Start, task.End )

                if ( attr === 'Start' || attr === 'End' ) {
                    let arr = ev.target.innerText.split('-')
                    if ( arr[0].length === 1 ) arr[0] = '0' + arr[0]
                    if ( arr[1].length === 1 ) arr[1] = '0' + arr[1]
                    if ( arr[2].length === 1 ) arr[2] = '0' + arr[2]
                    ev.target.innerText = arr[0] + '-' + arr[1] + '-' + arr[2]
                    saveDataCellChanges ( ev )

                    // Move End date accordingly
                    if ( attr === 'Start' ) {
                        let newEndDate = ev.target.innerText
                        newEndDate = increaseDate ( newEndDate, days )
                        project.setTask ( ev.target.id.lineIndex(), { End: convertDate(newEndDate, 'string') }, true )
                    }

                    if ( attr === 'Start' && compareDate ( ev.target.innerText, START_DATE_OBJ ) < 0 ) {
                        getProjectDateBounds ( project.taskData )
                        updateGanttTable()
                    }
                
                    if ( attr === 'End'   && compareDate ( ev.target.innerText, END_DATE_OBJ   ) > 0 ) {
                        getProjectDateBounds ( project.taskData )
                        updateGanttTable()
                    }
                } else
                    saveDataCellChanges ( ev )
            }

            if ( ev.target.id.includes ('data-col_') )
                saveHeaderCellChanges ( ev )
            
            updateGanttTable ()
        })
    }

    document.addEventListener ('click', (ev) => {
        log ( "EVENT: 'click'   TARGET: " + ev.target.id + "   CURRENT-TARGET: " + ev.currentTarget.id )

        if ( contextMenu && ev.target.id !== 'context-menu' ) {
            contextMenu.remove()
            contextMenu = undefined
        }

        ev.stopPropagation()
    })

    function updateDataTable ( idx ) {
        if ( idx !== undefined ) { // If idx is set, just update a single line and don't update whole table to improve performance
            for ( let attr in project.taskData[idx] ) {
                let elem = document.getElementById ( 'data-cell_' + idx + '_' + attr )
                
                if ( elem )
                    elem.innerText = project.taskData[idx][attr]
            }

            return
        }

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
                    'min-width:' + col.minWidth + 'px',
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

            headerCell.addEventListener ( 'dragover' , ev => ev.target.style.borderRightColor = 'red'       )
            headerCell.addEventListener ( 'dragleave', ev => {
                if ( isLastAttr )
                    ev.target.style.borderRightColor = '#666666'    
                else
                    ev.target.style.borderRightColor = '#b8b8b8'
            })

            if ( col.attributeName !== '#' ) {
                addContextMenuListener ( headerCell, 'data-table-header' )
                headerCell.setAttribute ("draggable", "true")
                addColDragListener ( headerCell )
            }

            addColDropListener ( headerCell )

            // Do not allow to rename certain columns
            if ( col.attributeName !== 'Task' && col.attributeName !== 'Start' && col.attributeName !== 'End' && col.attributeName !== '#') {
                addClickListenerText ( headerCell )
                addBlurListener ( headerCell )
            }

            let resizeHandle = createResizeHandle ( headerCell, 'data-resize-handle_' + attrNr, 'RIGHT' )
            eventListeners.addMouseDownListener ( resizeHandle )
            eventListeners.addMouseUpListener ( resizeHandle )
            eventListeners.addDblClickListener ( resizeHandle, resizeDataColumn )
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

            row.addEventListener ( 'dragover', (ev) => {
                for ( let cell of ev.target.parentElement.children )
                    cell.style.borderBottomColor = 'red'
            })

            row.addEventListener ( 'dragleave', (ev) => {
                for ( let cell of ev.target.parentElement.children )
                    cell.style.borderBottomColor = 'lightgray'
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

            addRowDragListener ( row )
            let attrNr = 0

            // Add data table cells
            for ( let colData of project.columnData ) {
                let taskAttr = colData.attributeName
                let foundColumn = project.columnData.find ( col => col.attributeName === taskAttr )
                let isLastAttr = attrNr===project.columnData.length-1
                let textIndent = 0
                let content = ""

                if ( taskAttr === 'Task' && task['groupLevel'] )
                    textIndent = task.groupLevel * GROUP_INDENT_SIZE

                if ( task.isGroup ) {
                    if ( taskAttr === 'Task' )
                       content = '▼&nbsp;&nbsp;' + content

                    if ( taskAttr === 'Start' ) {
                        let idxCnt = idx + 1

                        while ( idxCnt < project.taskData.length && project.taskData[idxCnt].groupLevel > project.taskData[idx].groupLevel ) {
                            if ( compareDate ( project.taskData[idxCnt].Start, project.taskData[idx].Start ) < 0 )
                                project.taskData[idx].Start = project.taskData[idxCnt].Start
                            idxCnt++
                        } 
                    }

                    if ( taskAttr === 'End' ) {
                        let idxCnt = idx + 1

                        while ( idxCnt < project.taskData.length && project.taskData[idxCnt].groupLevel > project.taskData[idx].groupLevel ) {
                            if ( compareDate ( project.taskData[idxCnt].End, project.taskData[idx].End ) > 0 )
                                project.taskData[idx].End = project.taskData[idxCnt].End
                            idxCnt++
                        } 
                    }
                }

                content += taskAttr==='#'?(idx+1).toString():task[taskAttr]    
                let cell = createAndAppendElement ( row, 'td', {
                    class: [
                        'text-overflow',
                        'fixed-col',
                        'cell',
                        'text-readonly',
                        (taskAttr==='Start'|| taskAttr==='End')?'cell-center':'',
                        task.isGroup?'task-group':''
                    ],
                    style: [
                        'min-width:' + foundColumn.minWidth + 'px',
                        'width:' + foundColumn.width + 'px',
                        isLastAttr?'border-right:1px solid #666':'border-right:1px solid #b8b8b8',
                        'padding:' + DATA_CELL_PADDING_VERTICAL + 'px ' + DATA_CELL_PADDING_HORIZONTAL + 'px',
                        'min-height:' + (GANTT_LINE_HEIGHT - DATA_CELL_PADDING_VERTICAL*2 - 1) + 'px',
                        'text-indent:' + textIndent + 'px'
                    ],
                    id: 'data-cell_' + idx + '_' + foundColumn.attributeName,
                    content: content
                })

                if ( taskAttr !== '#' ) {
                    addBlurListener ( cell )
                    addClickListenerText ( cell )
                }

                addContextMenuListener ( cell, 'data-table-line' )
                addRowDropListener ( cell )
                attrNr++
            }
        }
    }

    // Extend height to make sure that last line is fully visible when table was scrolled down
    document.getElementById ('table-area').style.height = 'calc(100% - ' + 23 + 'px)'

    function addRowDragListener ( elem ) {
        elem.addEventListener ( "dragstart" , (ev) => {
            dragType = DRAG_TYPE_ROW
            dragIdx = ev.currentTarget.id.lineIndex()
        })
    }

    function addRowDropListener ( elem ) {
        // Canceling dragover is required! Otherwise drop won't fire
        // https://stackoverflow.com/questions/32084053/why-is-ondrop-not-working
        elem.addEventListener ( "dragover" , ev => ev.preventDefault() )

        elem.addEventListener ( "drop" , function (ev) {
            for ( let cell of ev.target.parentElement.children )
                cell.style.borderBottomColor = 'lightgray'

            if ( dragType === DRAG_TYPE_COL ) {
                log ( "Not allowed to drop column on row!" )
                return
            }

            dragType = undefined
            let idx = ev.currentTarget.id.lineIndex()

            if ( idx === dragIdx )
                return

            ipcRenderer.send( "setWasChanged", true )
            log ( 'Moving line:', dragIdx, 'to current line:', idx )
            let taskDataArr = []
            let srcGroupLevel = project.taskData[dragIdx].groupLevel
            let dstGroupLevel = project.taskData[idx].groupLevel
            let idxCnt = dragIdx
            let idxEnd = undefined
            let addedLines = undefined
            let lineDeleteOffset = undefined
            let groupLevelShift = (dstGroupLevel - srcGroupLevel) + (project.taskData[idx].isGroup?1:0)

            do {
                let taskData = clone ( project.taskData[idxCnt] )
                taskData.groupLevel = taskData.groupLevel + groupLevelShift
                taskDataArr.push ( taskData )
                idxCnt++
            } while ( idxCnt < project.taskData.length && project.taskData[idxCnt].groupLevel > srcGroupLevel )

            idxEnd = idxCnt - 1
            addedLines = idxEnd - dragIdx + 1

            if ( idx > dragIdx && idx < dragIdx + addedLines ) {
                log ( "Not allowed to drop into own group!" )
                return
            }

            for ( idxCnt = idx ; idxCnt < idx + addedLines ; idxCnt++ ) {
                insertLineAboveOrBelow ( idxCnt, false )
                project.setTask ( idxCnt + 1, taskDataArr[idxCnt-idx], false )
            }

            if ( idx > dragIdx )
                lineDeleteOffset = 0
            else
                lineDeleteOffset = addedLines

            for ( idxCnt = 0 ; idxCnt < addedLines; idxCnt++ )
                removeLine ( dragIdx + lineDeleteOffset )

            dragIdx === undefined
            updateTables()
        })
    }

    function addColDragListener ( elem ) {
        elem.addEventListener ( "dragstart" , (ev) => {
            dragType = DRAG_TYPE_COL
            dragColAttr = ev.currentTarget.id.lineIndex()
        })
    }

    function addColDropListener ( elem ) {
        elem.addEventListener ( "dragover" , ev => ev.preventDefault() )

        elem.addEventListener ( "drop" , function (ev) {
            let toColAttr = ev.currentTarget.id.lineIndex()
            let colData = {}
            let fromColIdx = project.columnData.findIndex ( elem => elem.attributeName === dragColAttr )
            let toColIdx   = project.columnData.findIndex ( elem => elem.attributeName === toColAttr   )

            if ( toColIdx === project.columnData.length - 1 )
                ev.target.style.borderRightColor = '#666666'    
            else
                ev.target.style.borderRightColor = '#b8b8b8'

            if ( dragType === DRAG_TYPE_ROW ) {
                log ( "Not allowed to drop row on column!" )
                return
            }

            dragType = undefined

            if ( toColAttr === dragColAttr )
                return

            ipcRenderer.send( "setWasChanged", true )

            log ( 'Moving column:', fromColIdx, 'to:', toColIdx )

            for ( let attr in project.columnData[fromColIdx] )
                colData[attr] = project.columnData[fromColIdx][attr]

            project.columnData.splice ( toColIdx + 1, 0, colData )

            if ( toColIdx > fromColIdx )
                project.columnData.splice ( fromColIdx, 1 )
            else
                project.columnData.splice ( fromColIdx + 1, 1 )

            dragColAttr = undefined
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
            log.err ( "Unable to open project! Project data is empty!" )
            return
        }

        if ( prj.path ) {
            try {
                let loadedData = fs.readFileSync ( prj.path )
                let jsonData = JSON.parse ( loadedData )

                createNewProject ({name: prj.name, path: prj.path})

                for ( let data in jsonData )
                    project[data] = jsonData[data]

                GANTT_CELL_WIDTH = parseInt ( project.ganttCellWidth )
                document.getElementById('project-name').innerHTML = project.name
                getProjectDateBounds ( project.taskData )
                updateTables()
                config.set ( 'lastProject.path', prj.path )
                addToRecentProjects ( recentProjects, prj.path )    
            } catch ( err ) {
                log.err ( err )
            }
        }
    }

    function saveProject () {
        // Blur all lines (header and data cells) to make sure that data is saved
        for ( let col of project.columnData )
            document.getElementById ( 'data-col_' + col.attributeName ).blur()

        for ( let idx = 0 ; idx < project.taskData.length ; idx++ ) {
            for ( let attr in project.taskData[idx] ) {
                let elem = document.getElementById ( 'data-cell_'+idx+'_'+attr )

                if ( elem )
                    elem.blur()
            }
        }

        let dataToSave = {}
        dataToSave.name       = project.name
        dataToSave.columnData = project.columnData
        dataToSave.taskData   = project.taskData
        dataToSave.ganttCellWidth = GANTT_CELL_WIDTH.toString()
        let serialized = JSON.stringify ( dataToSave, null, "    " )
        fs.writeFileSync ( project.path, serialized )
        ipcRenderer.send( "setWasChanged", false )
        addToRecentProjects ( recentProjects, project.path )
    }

    function exportExcel ( path ) {
        let workbook = new excel.Workbook()
        let worksheet  = workbook.addWorksheet ( 'Gantt Data' )
        let excelColumns = []
        let width = undefined

        // Create data column headers
        for ( let [ idx, col ] of project.columnData.entries() ) {
            switch ( col.displayName ) {
                case '#'    : width = 4 ; break
                case 'Task' : width = 30; break
                case 'Start': width = 10; break
                case 'End'  : width = 10; break
                default     : width = 20
            }

            excelColumns.push ({
                header: col.displayName  ,
                key   : col.attributeName,
                width : width
            })
        }

        let projectDays = getLengthInDays ( START_DATE_OBJ, END_DATE_OBJ )
        let dateCounter = new Date ( convertDate (START_DATE_OBJ, 'number') )

        // Create calendar column headers
        for ( let col = 0 ; col < projectDays ; col++ ) {
            let dateArr = convertDate(dateCounter,'string').split('-')

            excelColumns.push ({
                header: dateArr[2] + '\n' + dateArr[1] + '\n' + dateArr[0],
                key: col,
                width: 3.2
            })

            increaseDate ( dateCounter )
        }

        // Add the columns to the worksheet
        worksheet.columns = excelColumns

        // Style header (first) line (Bold, Text-Wrap, etc)
        let headerRow = worksheet.getRow ( 1 )
        headerRow.height = 50
        headerRow.font = { name: 'Arial', family: 4, size: 11, bold: true }
        headerRow.eachCell ( function ( cell, rowNumber) {
            cell.alignment = { wrapText: true }
        })

        // Lock first line
		worksheet.views = [{
			state    : 'frozen',
			xSplit   : project.columnData.length,
			ySplit   : 1,
			zoomScale: 85
		}]

        let dateCounterObj = new Date ( START_DATE_OBJ.getTime() )
        let dayCnt = 1

        // Draw weekend and today column(s)
        while ( compareDate ( dateCounterObj, END_DATE_OBJ ) <= 0 ) {
            if ( isWeekend(dateCounterObj) ) {
                worksheet.getColumn(project.columnData.length + dayCnt).fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'efddef' } }
                worksheet.getColumn(project.columnData.length + dayCnt+1).fill = { type: 'pattern', pattern: 'none' , fgColor: { argb: 'ffffff' } }
            }

            if ( isToday(dateCounterObj) ) {
                worksheet.getColumn(project.columnData.length + dayCnt).fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'ff3737' } }
                worksheet.getColumn(project.columnData.length + dayCnt+1).fill = { type: 'pattern', pattern: 'none' , fgColor: { argb: 'ffffff' } }
            }

            increaseDate ( dateCounterObj )
            dayCnt++
        }

        // Add content
        for ( let [ idx, task ] of project.taskData.entries() ) {
            let row = {}

            // Add data table content
            for ( let col of project.columnData ) {
                if ( col.attributeName === '#' )
                    row[col.attributeName] = idx + 1
                else
                    row[col.attributeName] = task[col.attributeName]                
            }

            // Add Gantt table content
            let ganttBarStartOffset = getOffsetInDays ( START_DATE_OBJ, task.Start )
            let ganttBarLength      = getLengthInDays ( task.Start, task.End )

            worksheet.addRow ( row )
            row = worksheet.getRow ( idx + 2 ) // EXCEL Line index is 1-based and also skip first line

            let startCellIdx = ganttBarStartOffset + project.columnData.length + 1 // EXCEL Cell index is 1-based
            let endCellIdx   = startCellIdx + ganttBarLength

            for ( let cellIdx = startCellIdx ; cellIdx < endCellIdx ; cellIdx++ ) {
                row.getCell ( cellIdx ).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '80c080' }
                }
            }

        }

        fs.open ( path, 'w+', function ( err, fd ) {
			if ( err ) {
				log.err ( `Unable to open '${path}'! Error: `, err )

				if ( err.code === 'EBUSY' )
					log.err ( "File busy!" )
				else
                    log.err ( err )

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
            setTask: (idx, newData, shallUpdate) => {
                if ( newData ) {
                    for ( let attr in newData ) {
                        if ( project.taskData[idx][attr] !== newData[attr] )
                            ipcRenderer.send( "setWasChanged", true )

                        project.taskData[idx][attr] = newData[attr]
                    }

                    if ( compareDate ( project.taskData[idx].End, project.taskData[idx].Start ) < 0 ) {
                        project.taskData[idx].End = project.taskData[idx].Start
                        document.getElementById ( 'data-cell_' + idx + '_End' ).innerHTML = project.taskData[idx].End
                    }

                    if ( shallUpdate )
                        updateDataTable ( idx )
                    
                    updateChartBar (idx)
                }
            },
            deleteTask : (idx) => {
                project.taskData.splice ( idx, 1 )
                ipcRenderer.send( "setWasChanged", true )
            },
            name: 'No Name',
            taskData: [{
                Task : "",
                Start: convertDate ( new Date (), 'string' ),
                End  : convertDate ( new Date (), 'string' ),
                isGroup   : false,
                groupLevel: 0
            }]
        }

        let projectName = document.getElementById('project-name')
        projectName.innerHTML = project.name
        
        projectName.addEventListener ( 'click', (ev) => {
            if ( contextMenu ) {
                contextMenu.remove()
                contextMenu = undefined
            }

            ev.target.setAttribute ( 'contenteditable', true )
            ev.target.classList.remove ('text-readonly')
            ev.target.focus()
        })

        projectName.addEventListener ( 'blur', (ev) => {
            project.name = projectName.innerText
            ev.target.setAttribute ( 'contenteditable', false )
            ev.target.classList.add ('text-readonly')
            ev.target.blur()
        })

        project.columnData = [{
            "attributeName": "#",
            "displayName": "#",
            "minWidth": "20",
            "width": "20"
        },{
            "attributeName": "Task",
            "displayName": "Task",
            "minWidth": "200",
            "width":"200"
        },{
            "attributeName": "Start",
            "displayName": "Start",
            "minWidth": "50",
            "width":"50"
        },{
            "attributeName": "End",
            "displayName": "End",
            "minWidth": "50",
            "width":"50"
        }]

        getProjectDateBounds ( project.taskData )

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
            let ganttTableWrapper = document.getElementById ( 'gantt-table-wrapper' )
            let dataTableWidth    = document.getElementById ( 'data-table-wrapper' ).offsetWidth
            let delta             = ev.pageX - dataTableWidth
            let absDaysToMousePos = Math.ceil ( (ganttTableWrapper.scrollLeft + delta) / GANTT_CELL_WIDTH )

            if ( ev.deltaY < 0 )
                GANTT_CELL_WIDTH += GANTT_CELL_WIDTH
            else
                GANTT_CELL_WIDTH /= 2 
            
            if ( GANTT_CELL_WIDTH <  2 ) GANTT_CELL_WIDTH = 2
            if ( GANTT_CELL_WIDTH > 64 ) GANTT_CELL_WIDTH = 64

            updateGanttTable ( ev.pageX, absDaysToMousePos )
        }
    })
})

// Keep outside of DOMContentLoaded listener to avoid issues with function being undefined!
var scrollSvgHeader = function (ev) {
    let val = (document.getElementById ( 'data-table-wrapper' ).offsetWidth - this.scrollLeft) + 'px'
    document.getElementById ( 'gantt-table-header-svg').style.left = val
}