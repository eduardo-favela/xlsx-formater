const { ipcRenderer } = require('electron');

function convert() {
    const xlsxInput = document.getElementById('xlsxInput');
    /* console.log(xlsxInput) */
    ipcRenderer.send('convert', xlsxInput.files[0].path);
}
