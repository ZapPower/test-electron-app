const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('path')
// const XLSX = require('xlsx')


// creates a new window of the application
const createWindow = () => {

    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js')
        }
    })

    win.loadFile('index.html')
}

// listen for the ready signal to create the window (createWindow can only be called after the ready signal)
app.whenReady().then(() => {
    createWindow()

    // MacOS only -- creates a new window if all previous windows have been closed
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
})

// exit process if the window is closed
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
})

// load excel file when submitted and display the name
// document.getElementById('excelSubmit').addEventListener('click', function() {
//     console.log('began load');
//     const f = document.getElementById('excelFile');
//     var workbook = XLSX.read(f);
//     var sheet = workbook.getActiveSheet();
//     document.getElementById('sheetName').innerText = sheet.getName();
// });

