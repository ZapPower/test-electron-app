const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');

// creates a new window of the application
const createWindow = () => {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        },
        resizable: false,
        icon: path.join(__dirname, 'MSD_logo.jpg')
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
ipcMain.on('requestExcelFile', (event, filetype) => {
    extension = ".XLSX";
    if (filetype == "ebt") {
        extension = ".CSV";
    }
    dialog.showOpenDialog({
        title: `Select the excel file (${extension}) to be used`,
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

// when recieved request for getting download path, open save dialog and allow user to select path
// then send the download path back to the renderer process
ipcMain.on('requestDownloadPath', (event, pdfName) => {
    dialog.showSaveDialog({
        title: "Choose Path to Download File",
        defaultPath: path.join(__dirname, pdfName),
        buttonLabel: 'Save'
    }).then(path => {
        console.log("Download Canceled: " + path.canceled);
        if (!path.canceled) {
            const downloadPath = path.filePath.toString()
            console.log("Download Path: " + downloadPath);
            event.reply('downloadPath', downloadPath);
        }
    }).catch(err => {
        console.log(err);
    });
});