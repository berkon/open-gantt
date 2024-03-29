"use strict";

const ONE_DAY_IN_MS  = 86400000
const ONE_HOUR_IN_MS = 3600000

function createAndAppendElement ( parentElement, tagName, options ) {
    let newElem = document.createElement ( tagName )

    if ( options && options.class ) {
        if ( typeof options.class === 'string' )
            newElem.classList.add ( options.class )
        else {
            for ( let cl of options.class )
                if ( cl )
                    newElem.classList.add ( cl )
        }
    }

    if ( options && options.style ) {
        if ( typeof options.style === 'string' ) {
            let style = options.style.split(':')[0]
            let value = options.style.split(':')[1]
            newElem.style[style] = value
        } else {
            for ( let st of options.style ) {
                let style = st.split(':')[0]
                let value = st.split(':')[1]
                newElem.style[style] = value
            }
        }
    }

    if ( options && options.title && typeof options.title === 'string' )
        newElem.title = options.title

    if ( options && options.content && typeof options.content === 'string' )
        newElem.innerHTML = options.content

    if ( options && options.id && typeof options.id === 'string' )
        newElem.id = options.id

    if ( parentElement ) {
        if ( typeof parentElement === 'string' )
            parentElement = document.getElementById ( parentElement )

        if ( parentElement )
            parentElement.appendChild ( newElem )
    }

    return newElem
}

function compareDate ( date, dateToCompareTo ) {
    if ( date === undefined || dateToCompareTo === undefined )
        return null

    date            = convertDate ( date           , 'number' )
    dateToCompareTo = convertDate ( dateToCompareTo, 'number' )

    // Add time zone and/or DST offset to make dates comparable
    date            = new Date ( date            ).getTimeAbsUTC()
    dateToCompareTo = new Date ( dateToCompareTo ).getTimeAbsUTC()

    if ( date <   dateToCompareTo ) return -1
    if ( date === dateToCompareTo ) return  0
    if ( date >   dateToCompareTo ) return  1
}

function _ganttDateStr_To_dateObj ( dateStr ) {
    let dateArr = dateStr.split('-')
    let d = parseInt ( dateArr[0] )        // Day of month
    let m = parseInt ( dateArr[1] ) - 1    // Month is zero - based in Date() object, thus need to subtract 1
    let y = parseInt ( dateArr[2] ) + 2000 // Year (last two digits, thus need to add 2000)
    return new Date ( y, m, d, 0, 0 ,0, 0 )
}

function _ms1970_To_ganttDateStr ( val ) {
    let newDate = new Date ( val )
    let y = newDate.getFullYear() - 2000
    let m = newDate.getMonth() + 1
    let d = newDate.getDate()
    if ( y.toString().length === 1 ) y = '0' + y
    if ( m.toString().length === 1 ) m = '0' + m
    if ( d.toString().length === 1 ) d = '0' + d
    return d + "-" + m + "-" + y
}

function createResizeHandle ( parentElem, id, size ) {
    let div = document.createElement ( 'div' )
    div.id               = id
    div.style.top        = 0
    div.style.position   = 'absolute'
    div.style.userSelect = 'none'
    div.style.height     = parentElem.offsetHeight + 'px'

    switch ( size ) {
        case 'LEFT':
            div.style.left  = 0
            div.style.width = '3px'
            div.className   = 'column-selector'
            div.style.cursor= 'col-resize'
            break
        case 'RIGHT':
            div.style.right = 0
            div.style.width = '3px'
            div.className   = 'column-selector'
            div.style.cursor= 'col-resize'
            break
        case 'BODY':
        default: // Fully fill parent element
            div.style.left  = 0
            div.style.width = parentElem.offsetWidth + 'px'
            div.style.cursor= 'w-resize'
    }

    parentElem.appendChild ( div )
    return div
}

function isWeekend ( date ) {
    if ( date.getDay() === 0 || date.getDay() === 6 )
        return true
    else
        return false
}

function isToday ( date ) {
    let today = new Date ()

    if ( today.getFullYear() === date.getFullYear() &&
         today.getMonth   () === date.getMonth   () &&
         today.getDate    () === date.getDate    ()
    )
        return true
    else
        return false
}

function increaseDate ( date, numOfDays ) {
    date = convertDate ( date, 'object')

    if ( numOfDays === undefined )
        date.setDate ( date.getDate() + 1 )
    else
        date.setDate ( date.getDate() + numOfDays )

    return date
}

function decreaseDate ( date, numOfDays ) {
    date = convertDate ( date, 'object')

    if ( numOfDays === undefined )
        date.setDate ( date.getDate() - 1 )
    else
        date.setDate ( date.getDate() - numOfDays )

    return date
}

function getLengthInDays ( dateA, dateB ) {
    let delta = getOffsetInDays ( dateA, dateB )
    delta++ // icrease by one because we don't need the zero-based offset but the length 
    return delta
}

function getOffsetInDays ( dateA, dateB ) {
    // Important division by 86400000 must take place on the result value,
    // not on dateA/dateB, because dateA and dataB are NOT divisible by
    // 86400000! This is because of the UTC offset! So dividing dateA or dateB
    // by 86400000 always has a remainder if not located in timezone UTC (0 hours offset)
    dateA = convertDate ( dateA, 'number' )
    dateB = convertDate ( dateB, 'number' )

    // Add time zone and/or DST offset to make dates comparable
    dateA = new Date ( dateA ).getTimeAbsUTC()
    dateB = new Date ( dateB ).getTimeAbsUTC()

    let delta = (dateB - dateA) / ONE_DAY_IN_MS  // One day contains 86400000 milliseconds
    return delta
}

function convertDate ( date, toFormat ) {
    if ( typeof date === toFormat )
        return date

    switch ( typeof date ) {
        case 'object':
            if ( toFormat === 'string' ) return _ms1970_To_ganttDateStr ( date.getTime() )
            if ( toFormat === 'number' ) return date.getTime ()
            break
        case 'number':
            if ( toFormat === 'string' ) return _ms1970_To_ganttDateStr ( date )
            if ( toFormat === 'object' ) return new Date ( date )
            break
        case 'string': 
            if ( toFormat === 'object' ) return _ganttDateStr_To_dateObj ( date )
            if ( toFormat === 'number' ) return _ganttDateStr_To_dateObj ( date ).getTime()
            break
        default:
            log ( "ERROR: Invalid format for date conversion. Cannot convert: " + typeof date + " to " + toFormat, ERROR )
    }
}

// <type>_<lineIdx>_<colIdentifier>
String.prototype.type = function () {
    return this.split('_')[0]
}
// <type>_<lineIdx>_<colIdentifier>
String.prototype.lineIndex = function () {
    let idx = this.split('_')[1]

    if ( /^\d+$/.test ( idx ) )
        return parseInt ( idx ) // If string only contains digits => convert to number value
    else
        return idx
}
// <type>_<lineIdx>_<colIdentifier>
String.prototype.colIdentifier = function () {
    let colId = this.split('_')[2]

    if ( /^\d+$/.test ( colId ) )
        return parseInt ( colId ) // If string only contains digits => convert to number value
    else
        return colId
}

function createSVGRect ( svg, x, y, width, height, classes, id ) {
    let svgRect = document.createElementNS ( "http://www.w3.org/2000/svg", 'rect' ) // Use SVG name space
    svgRect.setAttribute ( "x", x.toString() )
    svgRect.setAttribute ( "y", y.toString() )
    svgRect.setAttribute ( "width" , width.toString() )
    svgRect.setAttribute ( "height", height.toString() )
    svgRect.setAttribute ( "class", classes )
    if ( id ) svgRect.setAttribute ( "id", id )
    svg.appendChild ( svgRect )
    return svgRect
}

function createSVGLine ( svg, x1, y1, x2, y2, classes, id ) {
    let svgLine = document.createElementNS ( "http://www.w3.org/2000/svg", 'line' ) // Use SVG name space
    svgLine.setAttribute ( "x1", x1.toString() )
    svgLine.setAttribute ( "y1", y1.toString() )
    svgLine.setAttribute ( "x2", x2.toString() )
    svgLine.setAttribute ( "y2", y2.toString() )
    svgLine.setAttribute ( "class", classes )
    if ( id ) svgLine.setAttribute ( "id", id )
    svg.appendChild ( svgLine )
    return svgLine
}

function createSVGText ( svg, str, x, y, id )  {
    let svgText = document.createElementNS ( "http://www.w3.org/2000/svg", 'text' ) // Use SVG name space
    svgText.setAttribute ( "x", x.toString() )
    svgText.setAttribute ( "y", y.toString() )
    svgText.classList.add ( "gantt-calendar-cell" )
    if ( typeof str === 'string')
        svgText.innerHTML = str
    else {
        svgText.innerHTML = str[0]
/*
        let tspan = document.createElementNS ( "http://www.w3.org/2000/svg", 'tspan' ) // Use SVG name space
        tspan.setAttribute ( "dx", -10 )
        tspan.setAttribute ( "dy", -12 )
        tspan.innerHTML = str[1]
        svgText.appendChild(tspan)
*/
    }
    if ( id ) svgText.setAttribute ( "id", id )
    svg.appendChild ( svgText )
//console.log ("***",str, svgText.getBBox())
    return svgText
}

function getProjectDateBounds ( tasks ) {
    // Find earliest and latest date among all tasks to define calendar range which should be displayed
    if ( tasks.length ) {
        START_DATE_OBJ = undefined
        END_DATE_OBJ   = undefined

        for ( let task of tasks ) {
            if ( task.Start && (START_DATE_OBJ === undefined || compareDate (task.Start, START_DATE_OBJ) < 0) )
                START_DATE_OBJ = new Date ( convertDate(task.Start, 'number' ) )

            if ( task.End && (END_DATE_OBJ === undefined || compareDate (task.End, END_DATE_OBJ) > 0) )
                END_DATE_OBJ = new Date ( convertDate(task.End, 'number' ) )
        }
    } else {
        START_DATE_OBJ = new Date ()
        END_DATE_OBJ   = new Date ()
    }

    decreaseDate ( START_DATE_OBJ, 10 )
    increaseDate ( END_DATE_OBJ  , 10 )

    // If no end date present set end date 30 days after start date
    if  ( START_DATE_OBJ && END_DATE_OBJ === undefined ) {
        END_DATE_OBJ = new Date ( convertDate(START_DATE_OBJ, 'number') )
        increaseDate ( END_DATE_OBJ, 20 )
    }

    // If no start date present set start date 30 days ahead of end date
    if  ( END_DATE_OBJ && START_DATE_OBJ === undefined ) {
        START_DATE_OBJ = new Date ( convertDate(END_DATE_OBJ, 'number') )
        decreaseDate ( START_DATE_OBJ, 20 )
    }
}

function centerSvgObj ( svgObj, width, height ) {
    if ( width !== undefined ) {
        let svgObjWidth   = Math.round ( svgObj.getBBox().width  )
        let svgObjXOffset = ( width  - svgObjWidth  ) / 2
        svgObj.setAttribute ( "x", Math.round ( svgObj.getBBox().x + svgObjXOffset ) )
    }

    if ( height !== undefined ) {
        let svgObjHeight  = Math.round ( svgObj.getBBox().height )
        let svgObjYOffset = ( height - svgObjHeight ) / 2
        svgObj.setAttribute ( "y", Math.round ( svgObj.getBBox().y + svgObjYOffset ) )
    }

    //console.log ( "**", svgObj.getBBox())
    //console.log ( "*", svgObj.getBBox().y, height, svgObj.getBBox().height, svgObjYOffset )
}

Number.prototype.toPx = function () {
    if ( typeof this === 'number' )
        return this + 'px'
    else
        return this
}

String.prototype.toVal = function () {
    if ( typeof this === 'string' )
        return parseInt ( this.substr(0, this.length-2) )
    else
        return this
}

function addToRecentProjects ( recentProjects, path ) {
    recentProjects.unshift ({ path: path })

    // Exclude index 0. Start with 1!
    for ( let idx = 1 ; idx < recentProjects.length ; idx++ ) {
        if ( recentProjects[idx].path === path )
            recentProjects.splice ( idx, 1 )
    }

    recentProjects = recentProjects.slice ( 0, 10 ) // Only keep max 10 entries
    config.set ( 'recentProjects', recentProjects )
}

// Adds the correct amount of time according to time zone and DST,
// to the getTime() result, to get comparable dates
Date.prototype.getTimeAbsUTC = function () {
    let tzAndDstOffsetWinter = Math.abs ( new Date ( 2021,1,1 ).getTimezoneOffset())*60*1000
    let tzAndDstOffsetSummer = Math.abs ( new Date ( 2021,6,1 ).getTimezoneOffset())*60*1000
    let curOffset = Math.abs ( this.getTimezoneOffset())*60*1000

    if ( curOffset === tzAndDstOffsetSummer )
        return this.getTime() + tzAndDstOffsetSummer
    else
        return this.getTime() + tzAndDstOffsetWinter
}

function clone ( obj ) {
    return JSON.parse ( JSON.stringify(obj) )
}

function checkIsHidden ( idx ) {
    let td = project.taskData

    if ( td[idx].groupLevel === 0 ) // On uppermost group level elments are never hidden (they can only be hidden if inside a collapsed group)
        return false

    let idxCnt = idx - 1

    while ( idxCnt >= 0 ) { // search tree upwards
        if ( td[idxCnt].isGroup && td[idxCnt].groupCollapsed && td[idxCnt].groupLevel < td[idx].groupLevel )
            return true

        if ( td[idxCnt].groupLevel === 0 ) // stop search if we reached group level 0
            return false

        idxCnt--
    }

    return true
}

function getNumOfGroupChilds ( idx ) {
    let td = project.taskData
    let numOfGroupChilds = 0

    if ( !td[idx].isGroup )
        return false

    let idxCnt = idx + 1

    while ( idxCnt < td.length && td[idxCnt].groupLevel > td[idx].groupLevel ) {
        numOfGroupChilds++
        idxCnt++
    }

    return numOfGroupChilds
}

function isGroupMember ( idx ) {
    if ( project.taskData[idx].groupLevel )
        return true
    else
        return false
}

function getGroup ( idx ) {
    let td = project.taskData
    let idxCnt = idx - 1

    while ( idxCnt >= 0 && td[idxCnt].groupLevel >= td[idx].groupLevel )
        idxCnt--

    return idxCnt
}

function getGroupBoundsOfChildren ( idx ) {
    let td = project.taskData
    let Start = undefined
    let End   = undefined
    let idxCnt = idx + 1

    while ( idxCnt < td.length && td[idxCnt].groupLevel > td[idx].groupLevel ) {
        if ( Start === undefined ) {
            Start = td[idxCnt].Start
            End   = td[idxCnt].End
            idxCnt++
            continue
        }

        if ( compareDate ( td[idxCnt].Start, Start ) < 0 )
            Start = td[idxCnt].Start

        if ( compareDate ( td[idxCnt].End, End ) > 0 )
            End = td[idxCnt].End

        idxCnt++
    }

    return { Start, End }
}

function getIdxAfterGroupEnd ( idx ) {
    let idxGroupEnd = idx

    do {
        idxGroupEnd++
    } while ( idxGroupEnd < project.taskData.length && project.taskData[idxGroupEnd].groupLevel >= project.taskData[idx].groupLevel )
    
    return idxGroupEnd
}