const { app, BrowserWindow } = require('electron');
const path = require('path');
const express = require('express');
const serverApp = require('./server'); // Importa o arquivo server.js
const { spawn } = require('child_process'); // Importa o módulo child_process

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false, // Para permitir o uso de require no frontend
    },
  });

  mainWindow.loadURL('http://localhost:9784'); // Carrega o backend
  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

// Função para abrir o terminal e executar o processo
function openCommandPrompt(process, seguenciaNF) {
  // Caminho para o script que você deseja executar
  const scriptPath = path.join(app.getAppPath(), 'server.js'); // Alterado para garantir o caminho correto no ambiente do Electron

  // Spawn a command prompt window
  const cmd = spawn('cmd.exe', ['/c', 'node', scriptPath, process, seguenciaNF], {
    shell: true,
    stdio: 'inherit', // Isso garante que os logs serão exibidos no terminal
  });

  // Exibir logs de erro (caso haja algum)
  cmd.on('error', (error) => {
    console.error(`Erro ao abrir o terminal: ${error.message}`);
  });

  // Finaliza o processo quando o terminal é fechado
  cmd.on('exit', (code) => {
    console.log(`Processo finalizado com código: ${code}`);
  });
}

// Inicializa o servidor Express quando o Electron estiver pronto
app.on('ready', () => {
  createWindow();
  serverApp.listen(9784, () => {
    console.log('Servidor rodando no Electron na porta 9784');
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});
