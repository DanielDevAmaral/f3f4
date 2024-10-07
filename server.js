const express = require('express');
const puppeteer = require('puppeteer');
const app = express();
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const http = require('http');
const socketIO = require('socket.io');

// Inicializa os recursos no servidor
app.use(cors());
app.use(express.json());
app.use(fileUpload());

// Cria um servidor HTTP para usar com o socket.io
const server = http.createServer(app);
const io = socketIO(server);  // Associando ao servidor HTTP

const connectedClients = [];

io.on('connection', (socket) => {
  console.log(`Cliente conectado: ${socket.id}`);
  connectedClients.push(socket);

  // Remove o socket da lista quando o cliente desconectar
  socket.on('disconnect', () => {
    const index = connectedClients.indexOf(socket);
    if (index !== -1) {
      connectedClients.splice(index, 1);
    }
    console.log('Cliente desconectado');
  });
});

async function startProcess(process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, linha, socket) {
  const browser = await puppeteer.launch({ 
    headless: true, 
    defaultViewport: null, 
    args: ['--start-maximized'],
    executablePath: puppeteer.executablePath()
  });
  const page = await browser.newPage();

  try {
    if (socket) {
      socket.emit('message', `Iniciando o processo na linha ${linha}...`);
    }

    if (process === 'F3') {
      await require(path.join(__dirname, 'F3'))(page, convenio, nrTitulo, seguenciaNF, estabelecimento, linha, socket);
    } else if (process === 'F4') {
      await require(path.join(__dirname, 'F4'))(page, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, linha, socket);
    }
    await browser.close();
    return { success: true };

  } catch (error) {
    console.error(`Erro no processo na linha ${linha}:`, error);
    if (socket) {
      socket.emit('message', `Erro no processo na linha ${linha}: ${error.message}`);
    }
    await browser.close();
    return { success: false, error: error.message };
  }
}

// Rota iniciada quando o documento é postado
app.post('/upload', async (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('Nenhum arquivo foi enviado.');
  }

  const process = req.body.process;
  const fileData = req.files.file.data;
  const seguenciaNF = req.body.seguenciaNF;

  const workbook = xlsx.read(fileData, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  if (connectedClients.length === 0) {
    console.error('Nenhum cliente conectado.');
    return res.status(500).json({ error: 'Nenhuma conexão de socket ativa. Conecte-se via WebSocket.' });
  }

  const socket = connectedClients[0];

  for (const [rowIndex, row] of sheet.entries()) {
    const convenio = row["Convênio"];
    const nrTitulo = row["Nr Retorno"];
    const retorno = String(row["Tipo "]);
    const glosa = row["Glosa Total"];
    const estabelecimento = row["Estabelecimento"];

    if (socket) {
      socket.emit('message', `Processando linha ${rowIndex + 1}: Estabelecimento: ${estabelecimento}, Convênio = ${convenio}, Nr Título = ${nrTitulo}${glosa == "0" ? ", sem GRG" : ", com GRG"}`);
    }

    const result = await startProcess(process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, rowIndex + 1, socket);

    if (!result.success) {
      console.error(`Erro no processo para Convênio: ${convenio}, Nr Título: ${nrTitulo}, result.error`);
      return res.status(500).json(result);
    }
  }

  res.send('Arquivo processado com sucesso.');
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
