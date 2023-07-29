const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');

// creates a new window of the application
const createWindow = () => {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: true,
            contextIsolation: false
        },
        resizable: false
    });

    win.loadFile('index.html');
}

// listen for the ready signal to create the window (createWindow can only be called after the ready signal)
app.whenReady().then(() => {
    createWindow();

    // MacOS only -- creates a new window if all previous windows have been closed
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
    });
})

// exit process if the window is closed
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
});

// when received request for getting excel file, open system dialog to allow user to select the excel file
// and send the file path back to the renderer process
ipcMain.on('requestExcelFile', (event) => {
    dialog.showOpenDialog({
        title: "Select the excel file (.XLSX) to be used",
        buttonLabel: "Submit"
    }).then(file => {
        console.log("File canceled: " + file.canceled);
        if (!file.canceled) {
            const filepath = file.filePaths[0].toString();
            console.log("Filepath: " + filepath);
            event.reply('excelFile', filepath);
        }
    }).catch(err => {
        console.log(err);
    });
});