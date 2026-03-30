import express, { Request, Response } from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';

const app = express();

// Usamos a memória para não precisar salvar o Excel no HD do servidor
const upload = multer({ storage: multer.memoryStorage() });

app.post('/upload', upload.single('planilha'), async (req: Request, res: Response): Promise<void> => {
    try {
        if (!req.file) {
            res.status(400).send('Nenhum arquivo enviado.');
            return;
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        
        // Pega a primeira aba da planilha
        const worksheet = workbook.worksheets[0];
        let textoParaImpressora = '';

        // Itera sobre cada linha da planilha
        worksheet.eachRow((row, rowNumber) => {
            // Pula a linha 1 se for o cabeçalho
            if (rowNumber === 1) return; 

            // Extrai os dados das colunas (Ex: Coluna 1 = Nome, Coluna 2 = Valor)
            // O padEnd(30, ' ') garante que o Nome SEMPRE ocupe 30 caracteres.
            const nome = row.getCell(1).text.padEnd(30, ' '); 
            
            // O padStart(10, ' ') alinha números à direita, ocupando 10 caracteres.
            const valor = row.getCell(2).text.padStart(10, ' '); 

            // Monta a linha do formulário e adiciona uma quebra de linha (\n)
            textoParaImpressora += `${nome}${valor}\n`;
        });

        // Configura a resposta para forçar o download de um arquivo .txt
        res.setHeader('Content-disposition', 'attachment; filename=formulario_matricial.txt');
        res.setHeader('Content-type', 'text/plain');
        res.send(textoParaImpressora);

    } catch (error) {
        console.error('Erro ao processar planilha:', error);
        res.status(500).send('Erro interno ao processar o arquivo.');
    }
});

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Servidor TypeScript rodando na porta ${PORT}`);
});
