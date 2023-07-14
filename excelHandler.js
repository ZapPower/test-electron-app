const XLSX = require('xlsx')
const ipc = require('electron').ipcRenderer

// // initializes event listener
// function init() {
//     console.log('initialized listener')
//     document.getElementById('excelFile').addEventListener('change', loadexcel, false);
// }

const submitButton = document.getElementById('excelsubmit')
submitButton.addEventListener('click', () => {
    document.getElementById('sheetname').innerHTML = "hello"
})

// // loads and parses excel file using xlsx
// function loadexcel() {
//     console.log('began load');
//     const f = document.getElementById('excelFile');
//     var workbook = XLSX.read(f);
//     var sheet = workbook.getActiveSheet();
//     document.getElementById('sheetName').innerText = sheet.getName();
// }

// function functionthingy() {
//     alert("alert")
//     const thingy = document.getElementById('thingy')
//     thingy.innerHTML = "New Thingy!!!"
// }