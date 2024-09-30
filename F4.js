const puppeteer = require('puppeteer');
const fs = require('fs');
const xlsx = require('xlsx');

module.exports = async (page, convenio, nrTitulo, retorno, glosa, seguenciaNF, estabelecimento, linha) => {

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
    let valor;
    let valorGRG;

    try {
        // Acessar a URL
        await page.goto('http://tasy.redeprimavera.com.br/#/login', {
            waitUntil: 'networkidle2',
            timeout: 60000  // Aumenta o tempo limite para 60 segundos
        });
        await delay(3000);
        // Login
        await page.waitForSelector('#loginUsername', { timeout: 10000 });
        await page.type('#loginUsername', 'TBCOSTA');
        await delay(3000);
        await page.type('#loginPassword', 'ty246810');
        await delay(3000);
        await page.click('input.btn-green.w-login-button.w-login-button--green');
        await delay(3000);
        // Fechar popup se existir
        try {
            await page.waitForSelector('button.w-header-tab__close.ng-scope', { timeout: 5000 });
            await page.click('button.w-header-tab__close.ng-scope');
            await delay(3000);
        } catch (e) {
            console.log('Popup não encontrado ou não existia.');
        }

        if (estabelecimento.trim() !== "Hospital Primavera Assist. Medica Hospitalar Ltda") {
            await page.waitForSelector('div.w-header-avatar');
            await page.click('div.w-header-avatar');
            await delay(3000);
        
            await page.waitForSelector('#userData-dropdown-options > ul > li:nth-child(1)');
            await page.click('#userData-dropdown-options > ul > li:nth-child(1)');
            await delay(3000);
            
            if(elementoExiste('#detail_3_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div')){
                const listbox = await page.$$('xpath/.//tasy-listbox[contains(@class, "ng-scope ng-isolate-scope")]');
                await listbox[1].click()
                //await page.click('#detail_3_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div');

            } else {
                await delay(3000);
                await page.click('#detail_1_container > div:nth-child(2) > div > div.w-attr-container__content > tasy-listbox > div');
            } 
        
            // 3. Encontrar e clicar no item que corresponde ao texto
            await delay(3000);
            const items = await page.$$('xpath/.//span[contains(@class, "ng-binding ng-scope")]');
            
            if(estabelecimento == 'Rede Primavera - Diagnose Barăo de Maruim'){
                estabelecimento = 'Rede Primavera - Diagnose Barão de Maruim'
            }
            
            for (const item of items) {
                const text = await item.evaluate(node => node.textContent.trim());

                if (text.includes(estabelecimento)) {
                    await item.focus(); // Foca no item antes de clicar
                    await item.click();  // Clica no item correto
                    break; // Sai do loop após clicar no item correto
                }
            }
            await delay(3000);
            await page.waitForSelector('.establishment-change-footer .btn-blue.establishment-change-buttons');
            await page.click('.establishment-change-footer .btn-blue.establishment-change-buttons');
            await delay(8000);
        }
        
        // Navegar para o perfil e para "Recurso de Glosas"
        await page.waitForSelector('div.w-header-avatar');
        await page.click('div.w-header-avatar');
        await delay(3000);

        const perfil = await page.$$('xpath/.//li[contains(@class, "wpopupmenu__item ng-scope")]');
        await perfil[1].click()
        await delay(3000);
        
        // Seleciona "Recurso de Glosa"
        const items = await page.$$('xpath/.//div[contains(@class, "wpopupmenu__label truncate ng-binding dropdown-toggle")]');
        for (const item of items) {
            const text = await item.evaluate(node => node.textContent.trim());
            if (text.includes("Recurso de Glosas")) {
                await item.focus(); // Foca no item antes de clicar
                await item.click();  // Clica no item correto
                break; // Sai do loop após clicar no item correto
            }
        }
        await delay(3000);
        if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
            const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
            await button.click()
        } else {
            await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
        }
        await delay(8000);
        
        // Navegar para "Retorno Convênio"
        await page.waitForSelector('button.w-apps__next');
        await page.click('button.w-apps__next');
        await delay(3000);

        await page.waitForSelector('#app-view > tasy-corsisf1 > div > w-mainlayout > div > div > w-launcher > div > div > div:nth-child(1) > w-apps > div > div.w-apps__carousel > ul > li:nth-child(2) > w-feature-app > a');
        await page.click('#app-view > tasy-corsisf1 > div > w-mainlayout > div > div > w-launcher > div > div > div:nth-child(1) > w-apps > div > div.w-apps__carousel > ul > li:nth-child(2) > w-feature-app > a');
        await delay(3000);

        // Selecionar convênio
        await page.waitForSelector('form > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > tasy-listbox');
        await page.click('form > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > tasy-listbox');
        await delay(3000);

        const dropdownOptions = await page.$$(
            'xpath/.//span[contains(@class, "ng-binding ng-scope")]'
        );
        await delay(3000);

        for (let option of dropdownOptions) {
            const text = await page.evaluate(el => el.textContent, option);
            if (text.includes(convenio)) {
                await option.click();
                break;
            }
        }
        await delay(3000);

        if (await elementoExiste("#detail_1_container > div:nth-child(4) > div.w-attr-container__content > tasy-wtextbox > div > div > input")){
            console.log("passou na primeira tentativa");
            const inputElement = await page.$('xpath/.//input[contains(@class, "gwt-TextBox  ng-valid ng-valid-required ng-pristine ng-untouched")]');
            await inputElement.click();
            await inputElement.type(String(nrTitulo));
        } else{
            console.log("passou na segunda tentativa");
            const inputElement = await page.$('xpath/.//input[contains(@class, "gwt-TextBox  ng-valid ng-valid-required ng-pristine ng-untouched")]');
            await inputElement.click();
            await inputElement.type(String(nrTitulo));
        }
        await delay(3000);

        await page.click('button.btn-green.wfilter-button.ng-binding');
        await delay(3000);

        if (retorno == "Retorno" || glosa != "0"){
            console.log("GRG identificado");

            // Verifique a cor de fundo do elemento
            const backgroundColor = await page.evaluate(() => {
                const element = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-left > div.slick-viewport.slick-viewport-top.slick-viewport-left.no-vertical-scrollbar > div > div > div.slick-cell.l1.r1.frozen.selected > div > span > div > div > div');
                return window.getComputedStyle(element).backgroundColor; // Obtém a cor de fundo
            });

            // Se a cor de fundo não for "VERDE", pula para o bloco else
            if (backgroundColor === 'rgb(0, 166, 88)') { // Converte a cor hex para RGB
                await delay(3000);
                try {
                    // Extraia o valor do span
                    valor = await page.evaluate(() => {
                        const span = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div > div.slick-cell.l7.r7.selected > div > span:nth-child(2)');
                        console.log(span.textContent);
                        return span ? span.textContent.trim() : null;
                    });
                
                    if (valor) {
                        console.log('Valor extraído:', valor);
                    } else {
                        console.log('Elemento não encontrado ou sem valor.');
                    }
                } catch (e) {
                    console.log('Erro ao extrair o valor:', e);
                }
                await delay(3000);
                await page.click('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div', { button: 'right' });
                await delay(3000);
                await page.click('#popupViewPort > li:nth-child(4)');
                await delay(3000); //document.querySelector("#popupViewPort > li:nth-child(4)")
                await page.click('#wpopupmenu_759884 > div > li:nth-child(3) > div.wpopupmenu__label.truncate.ng-binding');
                await delay(5000); //document.querySelector("#wpopupmenu_759884 > div > li:nth-child(3) > div.wpopupmenu__label.truncate.ng-binding")
                const check = await page.$('xpath/.//input[contains(@name, "IE_GERAR_NOVO_LOTE")]');
                await check.click()
                await delay(3000);
                try {
                    // Aguarde até o botão com o span "Gerar" estar visível
                    await page.waitForSelector("span.ng-binding.ng-scope", { visible: true });
                
                    // Use evaluate para selecionar o span pelo texto "Gerar" e clicar no botão pai
                    await page.evaluate(() => {
                        const span = Array.from(document.querySelectorAll('span.ng-binding.ng-scope'))
                            .find(el => el.textContent.includes('Gerar'));
                        if (span) {
                            const button = span.closest('button');
                            if (button) {
                                button.click();
                            }
                        }
                    });
                } catch (e) {
                    console.log('Erro ao tentar clicar no botão "Gerar":', e);
                }
                await delay(3000); //document.querySelector("#ngdialog13 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button")
                const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                await button.click();
                await delay(3000);
                try {
                    // Aguarde até o botão com o texto "Não" estar visível
                    await page.waitForSelector("button", { visible: true });
                
                    // Use evaluate para selecionar o botão pelo texto "Não" e clicar nele
                    await page.evaluate(() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        const button = buttons.find(el => el.textContent.trim() === 'Não');
                        if (button) {
                            button.click();
                        } else {
                            console.log('Botão "Não" não encontrado.');
                        }
                    });
                } catch (e) {
                    console.log('Erro ao tentar clicar no botão "Não":', e);
                }
                
                await delay(6000); 
                valorGRG = await page.evaluate(() => {
                    const span = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div > div.slick-cell.l9.r9.selected > div > span:nth-child(2)');
                    return span ? span.textContent.trim() : null;
                });
                if (valorGRG) {
                    console.log('Valor extraído:', valor);
                } else {
                    console.log('Elemento não encontrado ou sem valor.');
                }
                console.log("Finalizado")
                statusprocess = 'Finalizado com sucesso';
                await page.close();
            } else {
                await delay(3000);
                console.log("Executando o fechamento")
                try {
                    // Extraia o valor do span
                    valor = await page.evaluate(() => {
                        const span = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div > div.slick-cell.l7.r7.selected > div > span:nth-child(2)');
                        console.log(`O valor encontrado foi: ${span}`)
                        return span ? span.textContent.trim() : null;
                    });
                
                    if (valor) {
                        console.log('Valor extraído:', valor);
                    } else {
                        console.log('Elemento não encontrado ou sem valor.');
                    }
                } catch (e) {
                    console.log('Erro ao extrair o valor:', e);
                }
                await delay(3000);
                await page.click('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div');
                await delay(3000);
                await page.keyboard.down('Shift');
                await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada
    
                // Pressiona e mantém pressionada a tecla 'Alt'
                await page.keyboard.down('Alt');
                await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada
    
                // Pressiona a tecla '4'
                await page.keyboard.press('4');
                await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada
    
                // Libera a tecla 'Alt'
                await page.keyboard.up('Alt');
                await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada
    
                // Libera a tecla 'Shift'
                await page.keyboard.up('Shift');
                await delay(5000);
                try {
                    console.log("Verificando se a aba Guias sem repasse foi aberta");
                    //tentativa de clicar "OK"
                    await page.evaluate(() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        const button = buttons.find(el => el.textContent.trim() === 'OK');
                        if (button) {
                            button.click();
                        } else {
                            console.log('Botão "Ok" não encontrado.');
                        }
                    });
                    /*
                    await page.waitForSelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button", { visible: true });
                    await page.evaluate(() => {
                        const button = document.querySelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button");
                        if (button) {
                            button.click();
                        }
                    });
                    await page.evaluate(() => {
                        const button = document.querySelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button");
                        if (button) {
                            button.click();
                        }
                    });
                    */
                } catch (e) {
                    console.log('Popup não encontrado ou não existia.');
                }
                await delay(5000);
                if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                    const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                    await button.click()
                } else {
                    await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
                }        
                await delay(3000);
                if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                    const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                    await button.click()
                } else {
                    await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
                }   
                await delay(3000);
                if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                    const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                    await button.click()
                } else {
                    await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
                }
                await delay(3000);
                await page.click('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div', { button: 'right' });
                await delay(3000);
                await page.click('#popupViewPort > li:nth-child(4)');
                await delay(3000); //document.querySelector("#popupViewPort > li:nth-child(4)")
                await page.click('#wpopupmenu_759884 > div > li:nth-child(3) > div.wpopupmenu__label.truncate.ng-binding');
                await delay(5000); //document.querySelector("#wpopupmenu_759884 > div > li:nth-child(3) > div.wpopupmenu__label.truncate.ng-binding")
                const check = await page.$('xpath/.//input[contains(@name, "IE_GERAR_NOVO_LOTE")]');
                await check.click()
                await delay(3000);
                try {
                    // Aguarde até o botão com o span "Gerar" estar visível
                    await page.waitForSelector("span.ng-binding.ng-scope", { visible: true });
                
                    // Use evaluate para selecionar o span pelo texto "Gerar" e clicar no botão pai
                    await page.evaluate(() => {
                        const span = Array.from(document.querySelectorAll('span.ng-binding.ng-scope'))
                            .find(el => el.textContent.includes('Gerar'));
                        if (span) {
                            const button = span.closest('button');
                            if (button) {
                                button.click();
                            }
                        }
                    });
                } catch (e) {
                    console.log('Erro ao tentar clicar no botão "Gerar":', e);
                }
                await delay(3000); //document.querySelector("#ngdialog13 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button")
                const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                await button.click();
                await delay(3000); //document.querySelector("#ngdialog14 > div.ngdialog-content > div > div:nth-child(3) > div > button.dialog_ok_button.btn-blue")
                try {
                    // Aguarde até o botão com o texto "Não" estar visível
                    await page.waitForSelector("button", { visible: true });
                
                    // Use evaluate para selecionar o botão pelo texto "Não" e clicar nele
                    await page.evaluate(() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        const button = buttons.find(el => el.textContent.trim() === 'Não');
                        if (button) {
                            button.click();
                        } else {
                            console.log('Botão "Não" não encontrado.');
                        }
                    });
                } catch (e) {
                    console.log('Erro ao tentar clicar no botão "Não":', e);
                }
                
                await delay(6000); 
                valorGRG = await page.evaluate(() => {
                    const span = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div > div.slick-cell.l9.r9.selected > div > span:nth-child(2)');
                    return span ? span.textContent.trim() : null;
                });
                if (valorGRG) {
                    console.log('Valor extraído:', valor);
                } else {
                    console.log('Elemento não encontrado ou sem valor.');
                }
                console.log("Finalizado")
                statusprocess = 'Finalizado com sucesso';
                await page.close();
            }
        // caso não seja GRG
        } else{
            // somente realiza o fechamento
            console.log("Realizando Fechamento")
            await delay(3000);
            try {
                // Extraia o valor do span
                valor = await page.evaluate(() => {
                    const span = document.querySelector('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div > div.slick-cell.l7.r7.active.selected > div > span:nth-child(2)');
                    return span ? span.textContent.trim() : null;
                });
            
                if (valor) {
                    console.log('Valor extraído:', valor);
                } else {
                    console.log('Elemento não encontrado ou sem valor.');
                }
            } catch (e) {
                console.log('Erro ao extrair o valor:', e);
            }
            await delay(3000);
            await page.click('#datagrid > div.slick-pane.slick-pane-top.slick-pane-right > div.slick-viewport.slick-viewport-top.slick-viewport-right > div > div');
            await delay(3000);
            // Pressiona e mantém pressionada a tecla 'Shift'
            await page.keyboard.down('Shift');
            await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada

            // Pressiona e mantém pressionada a tecla 'Alt'
            await page.keyboard.down('Alt');
            await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada

            // Pressiona a tecla '4'
            await page.keyboard.press('4');
            await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada

            // Libera a tecla 'Alt'
            await page.keyboard.up('Alt');
            await delay(100); // Adiciona um pequeno atraso para garantir que a tecla foi registrada

            // Libera a tecla 'Shift'
            await page.keyboard.up('Shift');
            await delay(5000);
            try {
                console.log("Verificando se a aba Guias sem repasse foi aberta");
                //tentativa de clicar "OK"
                await page.evaluate(() => {
                    const buttons = Array.from(document.querySelectorAll('button'));
                    const button = buttons.find(el => el.textContent.trim() === 'OK');
                    if (button) {
                        button.click();
                    } else {
                        console.log('Botão "Ok" não encontrado.');
                    }
                });
                /*
                await page.waitForSelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button", { visible: true });
                await page.evaluate(() => {
                    const button = document.querySelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button");
                    if (button) {
                        button.click();
                    }
                });
                await page.evaluate(() => {
                    const button = document.querySelector("#ngdialog6 > div.ngdialog-content > div.dialog-box.dialog-default > div.dialog-footer.ng-scope > div:nth-child(3) > tasy-wdlgpanel-button > button");
                    if (button) {
                        button.click();
                    }
                });
                */
            } catch (e) {
                console.log('Popup não encontrado ou não existia.');
            }
            await delay(5000);
            if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                await button.click()
            } else {
                await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
            }        
            await delay(3000);
            if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                await button.click()
            } else {
                await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
            }   
            await delay(3000);
            if(elementoExiste('#ngdialog4 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue')){
                const button = await page.$('xpath/.//button[contains(@class, "dialog_ok_button btn-blue")]');
                await button.click()
            } else {
                await page.click('#ngdialog2 > div.ngdialog-content > div.dialog-box.wdialogbox-container.dialog-warning > div:nth-child(3) > div > button.dialog_ok_button.btn-blue');
            }
            await delay(3000);
            console.log("Finalizado")
            statusprocess = 'Finalizado com sucesso';
            await page.close();
        }

    } catch (e) {
        console.error('Ocorreu um erro:', e);
        statusprocess = 'Erro';
    } finally {
        // Função para salvar dados em uma planilha
        salvarDadosPlanilha(convenio, nrTitulo, valor, valorGRG, statusprocess);
    }
};

function salvarDadosPlanilha(convenio, nrTitulo, valor, valorGRG, statusprocess) {
    const filePath = './resultado_processamentoF4.xlsx';
    let workbook;
    let worksheet;

    // Verifica se o arquivo já existe
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        workbook = xlsx.utils.book_new();
        const headers = [{ convenio: "Convênio", nrTitulo: "Número do Título", valor: "Valor Minimo", valorGRG: "Valor GRG", statusprocess: "Status" }];
        worksheet = xlsx.utils.json_to_sheet(headers);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Resultados');
    }

    // Adiciona os novos dados
    const newData = [{ convenio, nrTitulo, valor, valorGRG, statusprocess }];
    xlsx.utils.sheet_add_json(worksheet, newData, { origin: -1, skipHeader: true });

    // Salva a planilha
    xlsx.writeFile(workbook, filePath);
    console.log('Dados salvos na planilha.');
}
