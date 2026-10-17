// Electron main process: one window that hosts the dashboard in app/index.html.
const { app, BrowserWindow, Menu, nativeTheme, shell } = require("electron");
const path = require("path");

function createWindow() {
  const win = new BrowserWindow({
    width: 1360,
    height: 880,
    minWidth: 980,
    minHeight: 640,
    title: "Inventory Tracker",
    titleBarStyle: "hiddenInset",
    trafficLightPosition: { x: 18, y: 20 },
    backgroundColor: nativeTheme.shouldUseDarkColors ? "#11181c" : "#f3f5f6",
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  // Links to websites open in the default browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: "deny" };
  });
  win.webContents.on("will-navigate", (e, url) => {
    if (!url.startsWith("file:")) {
      e.preventDefault();
      shell.openExternal(url);
    }
  });

  win.loadFile(path.join(__dirname, "app", "index.html"));
}

function buildMenu() {
  const template = [
    { role: "appMenu" },
    {
      label: "File",
      submenu: [
        {
          label: "Upload Excel Files…",
          accelerator: "CmdOrCtrl+O",
          click: (_, win) => win && win.webContents.executeJavaScript("window.openFilePicker && window.openFilePicker()"),
        },
        { type: "separator" },
        { role: "close" },
      ],
    },
    { role: "editMenu" },
    { role: "viewMenu" },
    { role: "windowMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

app.whenReady().then(() => {
  buildMenu();
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
