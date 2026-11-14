const { app, BrowserWindow } = require("electron");
const path = require("path");

function createWindow() {
  const win = new BrowserWindow({
    width: 560,
    height: 880,
    minWidth: 380,
    autoHideMenuBar: true,
    title: "معلومات شخصية",
    webPreferences: { contextIsolation: true }
  });
  win.loadFile(path.join(__dirname, "app", "index.html"));
}

app.whenReady().then(createWindow);
app.on("window-all-closed", () => app.quit());
