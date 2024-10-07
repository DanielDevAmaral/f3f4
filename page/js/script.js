const socket = io("http://localhost:3000")

document.getElementById('operationForm').addEventListener('submit', async function(event) {
    event.preventDefault(); // Previne o envio padrão do formulário

    // Captura os dados do formulário
    const formData = new FormData();
    const operation = document.querySelector('input[name="operation"]:checked').value;
    const fileUpload = document.getElementById('fileUpload').files[0];
    const sequenciaNR = document.getElementById('sequenciaNR').value;

    formData.append('process', operation);
    formData.append('file', fileUpload);
    formData.append('sequenciaNF', sequenciaNR);

    try {
        // Envia os dados para o servidor
        const response = await fetch('http://localhost:3000/upload', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            const result = await response.text();
            sucessoNotification('Processo iniciado com sucesso: ' + result);
        } else {
            throw new Error('Erro ao processar o arquivo');
        }
    } catch (error) {
        erroNotification('Erro: ' + error.message);
    }
});

function sucessoNotification(message) {
    Toastify({
        text: message,
        duration: 5000,
        gravity: 'top',
        position: 'right',
        backgroundColor: '#4CAF50',
        close: true
    }).showToast();
}

function erroNotification(message) {
    Toastify({
        text: message,
        duration: 5000,
        gravity: 'top',
        position: 'right',
        backgroundColor: '#f44336',
        close: true
    }).showToast();
}
