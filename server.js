const express = require('express');
const puppeteer = require('puppeteer');
const app = express();
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const cors = require('cors');
const ejs = require('ejs');
const { exec } = require('child_process');
const path = require('path');

app.use(cors());
app.use(express.json());
app.use(fileUpload());

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.get('/', (req, res) => {
  res.render('index');
});

// Função para iniciar o processo F3 ou F4
async function startProcess(process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, linha) {
  const browser = await puppeteer.launch({ 
    headless: true, 
    defaultViewport: null, 
    args: ['--start-maximized'],
    executablePath: puppeteer.executablePath() // Garante que o Chromium baixado seja usado
  });
  const page = await browser.newPage();

  try {
    if (process === 'F3') {
      await require(path.join(__dirname, 'F3'))(page, convenio, nrTitulo, seguenciaNF, estabelecimento, linha);
    } else if (process === 'F4') {
      await require(path.join(__dirname, 'F4'))(page, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, linha);
    }
    await browser.close();
    return { success: true };
  } catch (error) {
    console.error(`Erro no processo na linha ${linha}:`, error);
    await browser.close();
    return { success: false, error: error.message };
  }
}

app.post('/start-process', async (req, res) => {
  const { process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento } = req.body;

  try {
    const result = await startProcess(process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento);
    console.log('Resultado enviado ao front-end:', result);  // Log da resposta
    res.json(result);
  } catch (error) {
    console.error('Erro ao iniciar processo:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/upload', async (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('Nenhum arquivo foi enviado.');
  }

  const process = req.body.process;  // Recebe o processo selecionado no front-end
  const fileData = req.files.file.data;
  const fileType = req.files.file.mimetype;
  const workbook = xlsx.read(fileData, { type: 'buffer' });

  const sheetName = workbook.SheetNames[0];
  const sheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
  const seguenciaNF = req.body.seguenciaNF;

  for (const [rowIndex, row] of sheet.entries()) {
    const convenio = row["Convênio"];
    const nrTitulo = row["Nr Retorno"];
    const retorno = String(row["Tipo "]);
    const glosa = row["Glosa Total"];
    const estabelecimento = row["Estabelecimento"];

    console.log(`Processando linha ${rowIndex + 1}: Estabelecimento: ${estabelecimento}, Convênio = ${convenio}, Nr Título = ${nrTitulo}${glosa == "0" ? ", sem GRG" : ", com GRG"}`);

    const result = await startProcess(process, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, rowIndex + 1);

    if (!result.success) {
      console.error(`Erro no processo para Convênio: ${convenio}, Nr Título: ${nrTitulo}, result.error`);
      return res.status(500).json(result);
    }
  }

  res.send('Arquivo processado com sucesso.');
});

module.exports = app;
