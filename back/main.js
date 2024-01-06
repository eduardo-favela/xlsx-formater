const { app, BrowserWindow, ipcMain, Menu } = require('electron');
const path = require('path');
const convertXlsxToTxt = require('./convertXlsxToTxt');

let mainWindow;

function createWindow() {
    Menu.setApplicationMenu(null);
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        icon: __dirname + "/icons/favicon.png",
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        },
    });

    mainWindow.loadFile(path.join(__dirname, '../front/index.html'));

    mainWindow.on('closed', function () {
        mainWindow = null;
    });

    mainWindow.setResizable(false);
    mainWindow.setMaximizable(false);

    mainWindow.once('ready-to-show', () => { mainWindow.show(); })

    /* mainWindow.webContents.openDevTools(); */
}

app.on('ready', createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', function () {
    if (mainWindow === null) createWindow();
});

ipcMain.on('convert', (event, filePath) => {
    const convertion = convertXlsxToTxt(filePath);

    event.reply('convertResponse', convertion);
});
