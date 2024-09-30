const puppeteer = require('puppeteer');
const fs = require('fs');
const xlsx = require('xlsx');

module.exports = async (page, convenio, nrTitulo, seguenciaNF, estabelecimento, linha) => {

    async function elementoExiste(selector) {
        try {
            await page.waitForSelector(selector);
            return true;
        } catch (error) {
            return false;
        }
    }

    function delay(time) {
        return new Promise(function(resolve) { 
            setTimeout(resolve, time);
        });
    }

    let statusprocess;
    try {
        console.log(`Processando convênio: ${convenio}, Título: ${nrTitulo}, na modalidade F3...`);
    
        // Acessar a URL
        console.log('Acessando a página de login...');
        await page.goto('http://tasy.redeprimavera.com.br/#/login', {
            waitUntil: 'networkidle2',
            timeout: 60000
        });
        await delay(3000);
    
        // Login
        console.log('Realizando login...');
        await page.waitForSelector('#loginUsername', { timeout: 10000 });
        await page.type('#loginUsername', 'TBCOSTA');
        await delay(3000);
    
        await page.type('#loginPassword', 'ty246810');
        await delay(3000);
    
        await page.click('input.btn-green.w-login-button.w-login-button--green');
        await delay(3000);
        console.log('Login realizado com sucesso...');
    
        // Fechar popup se existir
        try {
            console.log('Verificando popup...');
            await page.waitForSelector('button.w-header-tab__close.ng-scope', { timeout: 5000 });
            await page.click('button.w-header-tab__close.ng-scope');
            await delay(3000);
        } catch (e) {
            console.log('Popup não encontrado ou não existia...');
        }
    
        if (estabelecimento.trim() !== "Hospital Primavera Assist. Medica Hospitalar Ltda") {
            console.log('Alterando o estabelecimento...');
            await page.waitForSelector('div.w-header-avatar');
            await page.click('div.w-header-avatar');
            await delay(3000);
    
            await page.waitForSelector('#userData-dropdown-options > ul > li:nth-child(1)');
            await page.click('#userData-dropdown-options > ul > li:nth-child(1)');
            await delay(3000);
    
            if (elementoExiste('#detail_3_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div')) {
                const listbox = await page.$$('xpath/.//tasy-listbox[contains(@class, "ng-scope ng-isolate-scope")]');
                await listbox[1].click();
            } else {
                await delay(3000);
                await page.click('#detail_1_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div');
            }
    
            await delay(3000);
            const items = await page.$$('xpath/.//span[contains(@class, "ng-binding ng-scope")]');
    
            if (estabelecimento == 'Rede Primavera - Diagnose Barão de Maruim') {
                estabelecimento = 'Rede Primavera - Diagnose Barão de Maruim';
            }
    
            for (const item of items) {
                const text = await item.evaluate(node => node.textContent.trim());
    
                if (text.includes(estabelecimento)) {
                    await item.focus();
                    await item.click();
                    break;
                }
            }
            console.log(`Estabelecimento alterado para ${estabelecimento} com sucesso.`);
            await delay(3000);
            await page.waitForSelector('.establishment-change-footer .btn-blue.establishment-change-buttons');
            await page.click('.establishment-change-footer .btn-blue.establishment-change-buttons');
            await delay(8000);
        }
    
        // Navegar para "Recurso de Glosas"
        console.log('Navegando para "Recurso de Glosas"...');
        await page.waitForSelector('div.w-header-avatar');
        await page.click('div.w-header-avatar');
        await delay(3000);
    
        const perfil = await page.$$('xpath/.//li[contains(@class, "wpopupmenu__item ng-scope")]');
        await perfil[1].click();
        await delay(3000);
    
        const items = await page.$$('xpath/.//div[contains(@class, "wpopupmenu__label truncate ng-binding dropdown-toggle")]');
        for (const item of items) {
            const text = await item.evaluate(node => node.textContent.trim());
            if (text.includes("Recurso de Glosas")) {
                await item.focus();
                await item.click();
                break;
            }
        }
        console.log('Entrando na seção "Recurso de Glosas"...');
        await delay(3000);
    
        // Manipular pop-up
        if (elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')) {
            const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
            await button.click();
        } else {
            await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
        }
        await delay(3000);
    
        console.log('Navegando para "Retorno Convênio"...');
        await page.waitForSelector('button.w-apps__next');
        await page.click('button.w-apps__next');
        await delay(3000);
    
        await page.waitForSelector('#app-view > tasy-corsisf1 > div > w-mainlayout > div > div > w-launcher > div > div > div:nth-child(1) > w-apps > div > div.w-apps__carousel > ul > li:nth-child(2) > w-feature-app > a');
        await page.click('#app-view > tasy-corsisf1 > div > w-mainlayout > div > div > w-launcher > div > div > div:nth-child(1) > w-apps > div > div.w-apps__carousel > ul > li:nth-child(2) > w-feature-app > a');
        await delay(3000);
    
        // Selecionar convênio
        console.log('Selecionando o convênio...');
        await page.waitForSelector('form > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > tasy-listbox');
        await page.click('form > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > tasy-listbox');
        await delay(3000);
    
        const dropdownOptions = await page.$$('xpath/.//span[contains(@class, "ng-binding ng-scope")]');
        await delay(3000);
    
        for (let option of dropdownOptions) {
            const text = await page.evaluate(el => el.textContent, option);
            if (text.includes(convenio)) {
                await option.click();
                break;
            }
        }
        await delay(3000);
    
        console.log(`Convênio ${convenio} selecionado com sucesso.`);
        
        // Preencher número de título
        if (await elementoExiste("#detail_1_container > div:nth-child(4) > div.w-attr-container__content > tasy-wtextbox > div > div > input")) {
            console.log(`Preenchendo número de título: ${nrTitulo}...`);
            const inputElement = await page.$('xpath/.//input[contains(@class, "gwt-TextBox ng-valid ng-valid-required ng-pristine ng-untouched")]');
            await inputElement.click();
            await inputElement.type(String(nrTitulo));
        } else {
            console.log(`Tentando preencher número de título: ${nrTitulo} na segunda tentativa...`);
            const inputElement = await page.$('xpath/.//input[contains(@class, "gwt-TextBox ng-valid ng-valid-required ng-pristine ng-untouched")]');
            await inputElement.click();
            await inputElement.type(String(nrTitulo));
        }
        await delay(3000);
    
        // Filtrar e interagir com tabela
        console.log('Filtrando dados...');
        await page.click('button.btn-green.wfilter-button.ng-binding');
        await delay(3000);
    
        // Clicar na primeira linha da tabela
        await page.click('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div');
        await delay(3000);
    
        // Clicar Alt + V
        console.log('Executando comando Alt + V...');
        await page.keyboard.down('Alt');
        await page.keyboard.press('V');
        await page.keyboard.up('Alt');
        await delay(3000);
    
        // Iterar para encontrar a sequência NF desejada
        console.log(`Procurando sequência NF: ${seguenciaNF}...`);
        const elements = await page.$$('xpath/.//div[contains(@class, "datagrid-cell-content-wrapper") and @style="line-height: 28px; "]/span');
    
        for (let i = 0; i < elements.length; i++) {
            const cellText = await elements[i].evaluate(el => el.textContent.trim());
    
            if (cellText.includes(seguenciaNF)) {
                console.log(`Encontrada sequência NF: ${cellText}`);
                await elements[i].click();
                break;
            }
        }
    
        // Conclusão
        statusprocess = 'Finalizado com sucesso';
        console.log('Processo concluído com sucesso.');
    } catch (e) {
        statusprocess = 'Erro';
        console.error(`Erro ao performar F3, verificar a linha: ${linha} no arquivo excel`);
    } finally {
        // Função para salvar dados em uma planilha
        salvarDadosPlanilha(convenio, nrTitulo, statusprocess);
    }
};

function salvarDadosPlanilha(convenio, nrTitulo, statusprocess) {
    const filePath = './resultado_processamentoF3.xlsx';
    let workbook;
    let worksheet;

    // Verifica se o arquivo já existe
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        workbook = xlsx.utils.book_new();
        const headers = [{ convenio: "Convênio", nrTitulo: "Número do Título", statusprocess: "Status" }];
        worksheet = xlsx.utils.json_to_sheet(headers);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Resultados');
    }

    // Adiciona os novos dados
    const newData = [{ convenio, nrTitulo, statusprocess }];
    xlsx.utils.sheet_add_json(worksheet, newData, { origin: -1, skipHeader: true });

    // Salva a planilha
    xlsx.writeFile(workbook, filePath);
    console.log('Dados salvos na planilha.');
}
