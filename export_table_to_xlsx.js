/*  this function is used to export any given table to xlsx file.
    tableDom - Required. The selector of the table DOM element (you can get it with: $("table:has(tbody)"))
	skipColumnsIndexes - Optional. In case some colums are not needed, they can be skipped by the yindex (starting from 0)
    exportName - Optional. The name of the exported file (suffix .xlsx will be added)
    important! to use this function, include the following code in your HTML header:
	<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.15.6/xlsx.core.min.js"></script> */
function downloadTableAsXLSX(tableDom, skipColumnsIndexes = [], exportName = 'export') { 
	// check that the library was added to the HTML header
    if (typeof XLSX === "undefined") {
        alert("Need to include the xlsx library in the HTML header");
        return;
    }
    // check that the table is not empty. another case could be an incorrect table selector ("tableDom")
    if (!tableDom) {
        alert('There is no data to export. The table is empty1');
        return;
    }
	var rowsLen = tableDom.rows.length;
    if (rowsLen < 2) {
        alert('There is no data to export. The table is empty2');
        return;
    }    

	// check that skipColumnsIndexes contains only numbers
	if (skipColumnsIndexes.some(element => isNaN(element))) {
		alert('The array "skipColumnsIndexes" contains non-number elements');
		return;
	}

	// preapre an array of column indexes (from 0...columns_length), and remove column indexes that are marked to be skipped
	const columnIndexes = Array.from(tableDom.rows[0].cells).map((_, index) => index)
        .filter(columnIndex => !skipColumnsIndexes.includes(columnIndex));

	// populate a 2-dimensional array with the cell values
    const matrix = [];
    for (let rowIndex = 0; rowIndex < rowsLen; rowIndex++) {
        matrix[rowIndex] = [];
        let matrixColumnIndex = 0;
        for (const columnIndex of columnIndexes) {
            const cellDom = tableDom.rows[rowIndex].cells[columnIndex];
            // if the cell consists of inner DOM element(s), check the child(ren) element(s) with recursion fucntion
            const cellValue = getInnerCellValue(cellDom);
            matrix[rowIndex][matrixColumnIndex] = cellValue;
            matrixColumnIndex++;			
        }
    }
    // export the created 2D array to an xlsx file
    const worksheet = XLSX.utils.aoa_to_sheet(matrix);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, exportName + ".xlsx");
}

/* recursive function that gets text of all elements in the hierarchy of a cell */
function getInnerCellValue(dom) {
    let text = '';
    const domTag = dom.tagName.toLowerCase();
    
    // ignore the values of hidden items
    if (dom.type === 'hidden' || dom.style.display === 'none') {
        return '';
    }
    
    // <select> can't have children so return its value
    if (domTag === 'select') {
        text = dom.options[dom.selectedIndex].text + ' ';
        return text;
    }
    
    // for <span>/<td> elements, get their text content and continue
    if (domTag === 'span' || domTag === 'td') {
		const ownNode = dom.childNodes;
		if (ownNode.length > 0) {
			var ownText = dom.childNodes[0].nodeValue;
			if (ownText) {
				text += ownText + ' ';
			}
		}   
    }

    // then concatenate the values of children elements
    
    // the "if" section uses recursion for elements with children, to get text of all the descendents 
    const cellChildren = dom.children;
	if (cellChildren.length > 0) { 
        for (let i = 0; i < cellChildren.length; i++) {
            text += getInnerCellValue(cellChildren[i]);
        }
    } else {
		// if this element does not have children then return its text when available
        if (domTag === 'th') {
            text = dom.textContent + ' ';
        }
        else if (domTag === 'input') {
            text = dom.value + ' ';
        }
		else if (domTag === 'a') {
            text = dom.textContent + ' ';
        }
        // in case of any another unique tag
        else {
            text = dom.textContent;
        }    
    }

    return text.trim();
}
