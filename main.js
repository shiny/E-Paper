
const {app, BrowserWindow, ipcMain } = require('electron')

let mainWindow

function createWindow () {
  mainWindow = new BrowserWindow({width: 920, height: 600})
  mainWindow.loadFile('index.html')
  mainWindow.on('closed', function () {
    mainWindow = null
  })
}
app.on('ready', createWindow)

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  if (mainWindow === null) {
    createWindow()
  }
})

ipcMain.on('get-app-path', event => {
  event.sender.send('got-app-path', app.getAppPath())
})
