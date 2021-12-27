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
    })
}

const addDblClickListener = ( elem, callback ) => {
    elem.addEventListener ( 'dblclick', (ev) => {
        let dataTableWidth   = document.getElementById ( 'data-table' ).offsetWidth
        let svgHeaderleftPos = document.getElementById ( 'gantt-table-header-svg' ).style.left.toVal() 
        mouseDownData = {
            action: MOUSE_ACTION_DATA_RESIZE_COLUMN,
            x     : ev.offsetX,
            elem  : ev.target.parentElement,
            ganttHeaderOffset: dataTableWidth - svgHeaderleftPos
        }

        callback ( ev, 50 )
        document.dispatchEvent ( new Event ('mousemove') ) // Strange! Need to trigger mousemove manually, otherwise SVG header is not redrawn
        mouseDownData = undefined
    })
}

exports.addMouseDownListener = addMouseDownListener
exports.addDblClickListener = addDblClickListener