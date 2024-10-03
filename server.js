const express = require('express');
const cors = require('cors'); // Importa o módulo cors
const app = express();
const printer = require('printer'); // Módulo printer (garanta que esteja instalado via npm)
const bodyParser = require('body-parser');

// Middleware para permitir CORS
app.use(cors()); // Permite CORS para todas as origens

// Middleware para lidar com requisições JSON
app.use(bodyParser.json());

// Rota para impressão
app.post('/imprimir', (req, res) => {
    const { printer: printerName, command } = req.body;

    if (!printerName || !command) {
        return res.status(400).json({ error: 'Parâmetros de impressão inválidos.' });
    }

    try {
        // Envia o comando de impressão para a impressora especificada
        printer.printDirect({
            data: command,
            printer: printerName,
            type: 'RAW',
            success: (jobID) => {
                console.log(`Impressão enviada com sucesso. ID do Job: ${jobID}`);
                res.status(200).json({ message: 'Impressão enviada com sucesso.' });
            },
            error: (err) => {
                console.error(`Erro ao enviar impressão: ${err}`);
                res.status(500).json({ error: 'Erro ao enviar impressão.' });
            }
        });
    } catch (error) {
        console.error('Erro ao processar a impressão:', error);
        res.status(500).json({ error: 'Erro no servidor de impressão.' });
    }
});

// Iniciar o servidor
const PORT = 3000; // Defina a porta
app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
    console.log(`Acesse o servidor em: http://localhost:${PORT}`);
});
