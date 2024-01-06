const { ipcRenderer } = require('electron');

function convert() {
    document.getElementById('convertionBtn').setAttribute('disabled', true)
    const xlsxInput = document.getElementById('xlsxInput');
    ipcRenderer.send('convert', xlsxInput.files[0].path);
}

ipcRenderer.on('convertResponse', (event, data) => {
    console.log(data);
    document.getElementById('convertionBtn').removeAttribute('disabled')
    let toast = document.getElementById('liveToast');
    let toastTitle = document.getElementById('toastTitle');
    let toastBody = document.getElementById('toastText');
    if (data.convertion) {
        toastTitle.innerHTML = 'Operación exitosa'
        toastBody.innerHTML = 'El archivo fue convertido correctamente'
    }
    else {
        toastTitle.innerHTML = 'Ocurrió un error'
        toastBody.innerHTML = data.error
        console.log(data.error)
    }
    const toastBootstrap = bootstrap.Toast.getOrCreateInstance(toast)
    toastBootstrap.show();
});
