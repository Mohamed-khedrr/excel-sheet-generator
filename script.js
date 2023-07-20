
let table = document.getElementsByClassName("sheet-body")[0],
    rows = document.getElementsByClassName("rows")[0],
    columns = document.getElementsByClassName("columns")[0]
tableExists = false

// =================

// =================

const generateError = (errorText) => {
    Swal.fire({
        title: '😢 Error!',
        text: errorText,
        icon: 'error',
        confirmButtonText: 'Ok'
    })
};

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    // Validate Inputs
    if (columnsNumber > 0 && rowsNumber > 0) {

        // Create Table
        table.innerHTML = ""
        for (let i = 0; i < rowsNumber; i++) {
            var tableRow = ""
            for (let j = 0; j < columnsNumber; j++) {
                tableRow += `<td contenteditable></td>`
            }
            table.innerHTML += tableRow
        }

        // Generate Flag
        tableExists = true

    } else {
        generateError('Please enter number for the number of rows and columns')
    }


}

const ExportToExcel = (type, fn, dl) => {
    if (!tableExists) {
        generateError('Please generate a table first')
        return
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}


