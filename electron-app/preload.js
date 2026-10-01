const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getAppInfo: () => ipcRenderer.invoke('app:info'),
  pickDirectory: () => ipcRenderer.invoke('dialog:pick-directory'),

  // task: monthly-calendar / list-calendar / leave-matrix / hours-summary / submit-files
  runTask: (payload) => ipcRenderer.invoke('task:run', payload),
  syncSchedules: () => ipcRenderer.invoke('storage:sync'),
  downloadRemoteFile: (payload) => ipcRenderer.invoke('storage:download', payload),
  openRemoteInExcel: (payload) => ipcRenderer.invoke('sharepoint:open-in-excel', payload),
  onTaskLog: (callback) => {
    const listener = (_event, chunk) => callback(String(chunk || ''));
    ipcRenderer.on('task:log', listener);
    return () => ipcRenderer.removeListener('task:log', listener);
  },

  getSharePointSessionStatus: () => ipcRenderer.invoke('sharepoint:session-status'),
  loginSharePoint: () => ipcRenderer.invoke('sharepoint:login'),
  deleteSharePointSession: () => ipcRenderer.invoke('sharepoint:session-delete'),

  queryDb: (payload) => ipcRenderer.invoke('db:query', payload),
});
