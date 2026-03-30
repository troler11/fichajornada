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
        await workbook.xlsx.load(req.file.buffer as any);
        const worksheet = workbook.worksheets[0];
        
        const motoristas = [];
        let motoristaAtual: any = null;

        worksheet.eachRow((row, rowNumber) => {
            // Função mais robusta para ler o valor da célula, não importa a formatação
            const lerCelula = (num: number) => {
                const celula = row.getCell(num);
                return celula.value ? celula.value.toString().trim() : '';
            };

            const colG = lerCelula(7);  // Coluna G (Nome ou ID)
            const colK = lerCelula(11); // Coluna K (Data)
            const colB = lerCelula(2);  // Coluna B (Linha/Código)

            // REGRA 1: Detectar se é uma linha de Cabeçalho de Motorista
            // Agora procura pelo hífen, mesmo se os espaços ao redor estiverem diferentes
            if (colG.includes('-') && colK !== '') {
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                // Divide "000204 - FRANCLIN GAMA" separando o ID do Nome
                const partes = colG.split('-');
                motoristaAtual = {
                    id: partes[0].trim(),
                    nome: partes.slice(1).join('-').trim(), // Pega tudo depois do primeiro hífen
                    data: colK,
                    viagens: []
                };
            }
            // REGRA 2: Detectar se é uma linha de Viagem Programada (tem dados na Coluna B)
            else if (motoristaAtual && colB !== '' && rowNumber > 1) {
                // Para não pegar as linhas de "Realizado" (que não tem a linha preenchida igual)
                // Se a sua linha "Realizado" também tiver a coluna B preenchida, me avise!
                const viagem = {
                    linha: colB,
                    veiculo: lerCelula(3),       // Col C
                    checkList1: lerCelula(4),    // Col D
                    deslocamento1: lerCelula(5), // Col E
                    pontoInicial: lerCelula(6)   // Col F
                };
                
                motoristaAtual.viagens.push(viagem);
            }
        });

        if (motoristaAtual) motoristas.push(motoristaAtual);

        // Devolve os dados em JSON para a tela do site!
        res.json(motoristas);

    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao ler a planilha.');
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
