/*  this function is used to export any given table to xlsx file.
    table_dom - Required. The selector of the table DOM element (you can get it with: $("table:has(tbody)"))
	skip_columns_indexes - Optional. In case some colums are not needed, they can be skipped by the yindex (starting from 0)
    export_name - Optional. The name of the exported file (suffix .xlsx will be added)
    important! to use this function, include the following code in your HTML header:
	<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.15.6/xlsx.core.min.js"></script> */
function download_table_as_xlsx(table_dom, skip_columns_indexes = [], export_name = 'export') { 
	// check that the library was added to the HTML header
    if (typeof XLSX === "undefined") {
        alert("Need to include the xlsx library in the HTML header");
        return;
    }
    // check that the table is not empty. another case could be an incorrect table selector ("table_dom")
    if (table_dom === undefined) {
        alert('There is no data to export. The table is empty');
        return;
    }
	var rows_len = table_dom.rows.length;
    if (rows_len < 2) {
        alert('There is no data to export. The table is empty');
        return;
    }    

	// check that skip_columns_indexes contains only numbers
	if (skip_columns_indexes.some(element => isNaN(element))) {
		alert('The array ""skip_columns_indexes" comprised of non-number elements');
		return;
	}

	// preapre array of columns indexes (from 0...columns_length), and remove columns indexes that marked to be skipped
	const column_indexes = Array.from(table_dom.rows[0].cells).map((_, index) => index)
        .filter(column_index => !skip_columns_indexes.includes(column_index));

	// populate a 2-dimensional array with the cell values
    const matrix = [];
    for (let row_i = 0; row_i < rows_len; row_i++) {
        matrix[row_i] = [];
        let matrix_column_i = 0;
        for (const column_i of column_indexes) {
            const cell_dom = table_dom.rows[row_i].cells[column_i];
            // if the cell consists of inner DOM element(s), check the child(ren) element(s) with recursion fucntion
            const cell_value = get_inner_cell_value(cell_dom);
            matrix[row_i][matrix_column_i] = cell_value;
            matrix_column_i++;			
        }
    }
    // export the created 2D array to an xlsx file
    const worksheet = XLSX.utils.aoa_to_sheet(matrix);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, export_name + ".xlsx");
}

/* recursive function that gets text of all elements in the hierarchy of a cell */
function get_inner_cell_value(dom) {
    let text = '';
    const dom_tag = dom.tagName.toLowerCase();
    
    // ignore the values of hidden items
    if (dom.type === 'hidden' || dom.style.display === 'none') {
        return '';
    }
    
    // select can't have children so return its value
    if (dom_tag === 'select') {
        text = dom.options[dom.selectedIndex].text + ' ';
        return text;
    }
    
    // for span/td/a elements, get their text content and continue
    if (dom_tag === 'span' || dom_tag === 'td') {
		const own_node = dom.childNodes;
		if (own_node.length > 0) {
			var own_text = dom.childNodes[0].nodeValue;
			if (own_text) {
				text += own_text + ' ';
			}
		}   
    }

    // then concatenate the values of children elements
    
    // the "if" section uses recursion for elements with children, to get text of all the descendents 
    var cell_children = dom.children;
	if (cell_children.length > 0) { 
        for (let i = 0; i < cell_children.length; i++) {
            text += get_inner_cell_value(cell_children[i]);
        }
    } else {
		// if this element does not have children then return its text when available
        if (dom_tag === 'th') {
            text = dom.textContent + ' ';
        }
        else if (dom_tag === 'input') {
            text = dom.value + ' ';
        }
		else if (dom_tag === 'a') {
            text = dom.textContent + ' ';
        }
        // in case of any another unique tag
        else {
            text = dom.textContent;
        }    
    }

    return text.trim();
}
