import express, { Request, Response } from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';

const app = express();
app.use(express.json());
app.use(express.static('public')); 

const upload = multer({ storage: multer.memoryStorage() });

app.post('/processar-excel', upload.single('planilha'), async (req: Request, res: Response): Promise<void> => {
    if (!req.file) {
        res.status(400).send('Nenhum arquivo enviado.');
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        // O "as any" evita aquele erro do TypeScript com o Buffer
        await workbook.xlsx.load(req.file.buffer as any);
        const worksheet = workbook.worksheets[0];
        
        const motoristas: any[] = [];
        let motoristaAtual: any = null;

        worksheet.eachRow((row, rowNumber) => {
            // Função BLINDADA para ler qualquer formato de célula do Excel
            const lerCelula = (num: number) => {
                const celula = row.getCell(num);
                if (!celula || celula.value === null || celula.value === undefined) return '';
                
                if (typeof celula.value === 'object') {
                    if ('richText' in celula.value) {
                        return celula.value.richText.map((rt: any) => rt.text).join('').trim();
                    }
                    if ('result' in celula.value) {
                        return String(celula.value.result).trim();
                    }
                    if (celula.value instanceof Date) {
                        return celula.value.toLocaleDateString('pt-BR');
                    }
                }
                return String(celula.value).trim();
            };

            const colG = lerCelula(7);  // Coluna G
            const colK = lerCelula(11); // Coluna K
            const colB = lerCelula(2);  // Coluna B

            // REGRA 1: Detectar Cabeçalho do Motorista
            // Se a Coluna G tem nome, a Coluna K tem data, e a Coluna B está vazia
            if (colG.length > 5 && colK.length > 5 && colB === '') {
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                let id = "N/D";
                let nome = colG;
                
                // Separa o ID do Nome se tiver o traço
                if (colG.includes('-')) {
                    const partes = colG.split('-');
                    id = partes[0].trim();
                    nome = partes.slice(1).join('-').trim();
                }

                motoristaAtual = {
                    id: id,
                    nome: nome,
                    data: colK,
                    viagens: []
                };
            }
            // REGRA 2: Detectar Viagem (Tem dados na Coluna B e já achou um motorista)
            else if (motoristaAtual && colB !== '' && rowNumber > 2) {
                motoristaAtual.viagens.push({
                    linha: colB,
                    veiculo: lerCelula(3),       // Col C
                    checkList1: lerCelula(4),    // Col D
                    deslocamento1: lerCelula(5), // Col E
                    pontoInicial: lerCelula(6)   // Col F
                });
            }
        });

        if (motoristaAtual) motoristas.push(motoristaAtual);

        res.json(motoristas);

    } catch (error) {
        console.error("Erro ao processar planilha:", error);
        res.status(500).send('Erro interno ao ler a planilha.');
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
