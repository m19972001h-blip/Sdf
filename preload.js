const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("api", {
  inv: (c, a) => ipcRenderer.invoke(c, a),
  on: f => ipcRenderer.on("prog", (_, d) => f(d))
});
