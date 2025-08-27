const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  createProject: (data) => ipcRenderer.invoke('project:create', data),
  getProjects: () => ipcRenderer.invoke('project:list')
});
