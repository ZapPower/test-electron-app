const XLSX = require('xlsx');
const {ipcRenderer} = require('electron')
const fs = require('fs');

// const submitButton = document.getElementById('excelsubmit')
// submitButton.addEventListener('click', () => {
//     loadexcel()
// })

// loads and parses excel file using xlsx
function loadexcel() {

    ipcRenderer.send('requestExcelFile');

    ipcRenderer.on('excelFile', (event, filepath) => {
        console.log(filepath.toString())
        fs.readFile(filepath.toString(), 'utf-8', (err, data) => {
            if (err) {
                alert("An error occurred when reading the file : " + err.message);
                return;
            }
            var workbook = XLSX.read(data);
            var sheet = workbook.SheetNames[0];
            document.getElementById('sheetname').innerText = sheet;
        });
        
    })

    // const f = form.excelfile.files;
    // var workbook = XLSX.read(f);
    // var sheet = workbook.SheetNames[0];
    // document.getElementById('sheetName').innerText = sheet;
}



// function functionthingy() {
//     alert("alert")
//     const thingy = document.getElementById('thingy')
//     thingy.innerHTML = "New Thingy!!!"
// }