const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  loadData:    ()        => ipcRenderer.invoke("data:load"),
  saveData:    (entries) => ipcRenderer.invoke("data:save", entries),
  loadSettings:()        => ipcRenderer.invoke("settings:load"),
  saveSettings:(s)       => ipcRenderer.invoke("settings:save", s),
  openFolder:  ()        => ipcRenderer.invoke("data:folder"),
  saveFile:    (opts)    => ipcRenderer.invoke("data:saveFile", opts),
  scanInvoice: (opts)    => ipcRenderer.invoke("scan:invoice", opts)
});
