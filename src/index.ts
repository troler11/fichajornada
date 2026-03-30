import express, { Request, Response } from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Adicionando suporte para o frontend ler JSON
app.use(express.json());
// Aqui você configuraria para servir sua página HTML (frontend)
app.use(express.static('public')); 

app.post('/processar-excel', upload.single('planilha'), async (req: Request, res: Response): Promise<void> => {
    if (!req.file) {
        res.status(400).send('Nenhum arquivo enviado.');
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];
        
        const motoristas = [];
        let motoristaAtual: any = null;

        worksheet.eachRow((row, rowNumber) => {
            const colG = row.getCell(7).text; // Coluna G (Nome ou ID)
            const colK = row.getCell(11).text; // Coluna K (Data)
            const colB = row.getCell(2).text; // Coluna B (Linha)

            // REGRA 1: Detectar se é uma linha de Cabeçalho de Motorista
            if (colG.includes(' - ') && colK !== '') {
                // Se já tínhamos um motorista sendo processado, salva ele na lista
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                const [id, nome] = colG.split(' - ');
                motoristaAtual = {
                    id: id.trim(),
                    nome: nome.trim(),
                    data: colK.trim(),
                    viagens: []
                };
            }
            // REGRA 2: Detectar se é uma linha de Viagem Programada (tem a Linha na Col B)
            else if (motoristaAtual && colB !== '' && rowNumber > 3) {
                const viagem = {
                    linha: colB,
                    veiculo: row.getCell(3).text, // Col C
                    checkList1: row.getCell(4).text, // Col D (Ajuste conforme sua planilha)
                    deslocamento1: row.getCell(5).text, // Col E
                    pontoInicial: row.getCell(6).text, // Col F
                    // ... mapear o restante das colunas ...
                };
                motoristaAtual.viagens.push(viagem);
            }
        });

        // Adiciona o último motorista lido à lista
        if (motoristaAtual) motoristas.push(motoristaAtual);

        // Devolve os dados em JSON para a tela do site!
        res.json(motoristas);

    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao ler a planilha.');
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
