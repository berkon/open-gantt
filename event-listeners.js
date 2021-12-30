const addMouseDownListener = ( elem ) => {
    elem.addEventListener ( 'mousedown', (ev) => {
        ev.target.parentElement.setAttribute ( "draggable", "false" )
        let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
        let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
        mouseDownData  = {
            action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
            x     : ev.pageX,
            elem  : ev.target.parentElement,
            width : ev.target.parentElement.offsetWidth,
            ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
        }

        // Select all elements starting with 'data-cell_' and ending with the corresponding attribute
        // This results in something like: 'data-cell_'*'_'<Attribute>
        queriedElements = document.querySelectorAll ( "[id^='data-cell_'][id$="+mouseDownData.elem.innerText+"]" )
    })
}

const addMouseUpListener = ( elem ) => {
    elem.addEventListener ( 'mouseup', (ev) => {
        ev.target.parentElement.setAttribute ( "draggable", "true" )
    })
}

const addDblClickListener = ( elem, callback ) => {
    elem.addEventListener ( 'dblclick', (ev) => {
        log ( "Event 'dblclick' target: " + ev.target.id )
        let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
        let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
        mouseDownData = {
            action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
            x     : ev.offsetX,
            elem  : ev.target.parentElement,
            ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
        }

        let foundColumn = project.columnData.find ( col => col.attributeName === mouseDownData.elem.id.split('_')[1] )
        callback ( ev, foundColumn.minWidth )
        mouseDownData = undefined
    })
}

exports.addMouseDownListener = addMouseDownListener
exports.addMouseUpListener   = addMouseUpListener
exports.addDblClickListener  = addDblClickListener