const addMouseDownListener = ( elem ) => {
    elem.addEventListener ( 'mousedown', (ev) => {
        let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
        let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
        mouseDownData  = {
            action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
            x     : ev.pageX,
            elem  : ev.target.parentElement,
            width : ev.target.parentElement.offsetWidth,
            ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
        }

        queriedElements = document.querySelectorAll ( "[id^='data-cell_'][id$="+mouseDownData.elem.innerText+"]" )
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
exports.addDblClickListener = addDblClickListener