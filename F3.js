const puppeteer = require('puppeteer');

// Obtém os argumentos passados para o script Node.js
const args = process.argv.slice(2);

// Certifique-se de que os argumentos estão sendo recebidos corretamente
console.log("Argumentos recebidos: ", args);

// Espera-se que os argumentos estejam na seguinte ordem:
// [convenio, nrTitulo, sequenciaNF, estabelecimento, index, retorno, glosa]
const [convenio, nrTitulo, sequenciaNF, estabelecimento, index, retorno, glosa] = args;

(async () => {
    // Iniciar o navegador do Puppeteer
    const browser = await puppeteer.launch({ headless: false, defaultViewport: null }); // Não rodar em modo headless para visualização
    const page = await browser.newPage();
    
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
    
        // Alterar o estabelecimento, se necessário
        if (estabelecimento.trim() !== "Hospital Primavera Assist. Medica Hospitalar Ltda") {
            console.log('Alterando o estabelecimento...');
            await page.waitForSelector('div.w-header-avatar');
            await page.click('div.w-header-avatar');
            await delay(3000);
    
            await page.waitForSelector('#userData-dropdown-options > ul > li:nth-child(1)');
            await page.click('#userData-dropdown-options > ul > li:nth-child(1)');
            await delay(3000);
    
            if (await elementoExiste('#detail_3_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div')) {
                const listbox = await page.$$('xpath/.//tasy-listbox[contains(@class, "ng-scope ng-isolate-scope")]');
                await listbox[1].click();
            } else {
                await delay(3000);
                await page.click('#detail_1_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div');
            }
    
            await delay(3000);
            const items = await page.$$('xpath/.//span[contains(@class, "ng-binding ng-scope")]');
    
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
        if (await elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')) {
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
        }
        await delay(3000);
    
        // Filtrar e interagir com tabela
        console.log('Filtrando dados...');
        await page.click('button.btn-green.wfilter-button.ng-binding');
        await delay(3000);
    
        // Clicar na primeira linha da tabela com a sequencia correta
        const elements = await page.$$('xpath/.//div[contains(@class, "slick-cell l0 r0 right-aligned active selected")]/div/span');
        for (const el of elements) {
            const text = await el.evaluate(node => node.textContent);
            if (text.includes(sequenciaNF)) {
                await el.click();
                break;
            }
        }
    
        console.log('Processo finalizado com sucesso.');
        statusprocess = "Finalizado";
    } catch (err) {
        console.log('Erro durante o processo:', err);
        statusprocess = "Erro";
    } finally {
        await browser.close();
    }
})();
