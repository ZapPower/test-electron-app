const XLSX = require('xlsx')

// const submitButton = document.getElementById('excelsubmit')
// submitButton.addEventListener('click', () => {
//     loadexcel()
// })

// loads and parses excel file using xlsx
function loadexcel(form) {
    const f = form.excelfile.files
    var workbook = XLSX.read(f);
    var sheet = workbook.SheetNames[0];
    document.getElementById('sheetName').innerText = sheet;
}

// function functionthingy() {
//     alert("alert")
//     const thingy = document.getElementById('thingy')
//     thingy.innerHTML = "New Thingy!!!"
// }