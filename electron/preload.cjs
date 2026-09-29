const { contextBridge } = require('electron');

contextBridge.exposeInMainWorld('aetherVizDesktop', {
  isDesktop: true
});