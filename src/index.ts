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
        await workbook.xlsx.load(req.file.buffer as any);
        const worksheet = workbook.worksheets[0];
        
        const motoristas: any[] = [];
        let motoristaAtual: any = null;

        worksheet.eachRow((row, rowNumber) => {
            const lerCelula = (num: number) => {
                const celula = row.getCell(num);
                if (!celula || celula.value === null || celula.value === undefined) return '';
                if (typeof celula.value === 'object') {
                    if ('richText' in celula.value) return celula.value.richText.map((rt: any) => rt.text).join('').trim();
                    if ('result' in celula.value) return String(celula.value.result).trim();
                    if (celula.value instanceof Date) return celula.value.toLocaleDateString('pt-BR');
                }
                return String(celula.value).trim();
            };

            const colB = lerCelula(2);  // Linha
            const colG = lerCelula(7);  // Nome do Motorista OU Ponto Final
            const colL = lerCelula(12); // Data (Ajustado para a Coluna L da foto)

            // REGRA 1: É o cabeçalho do motorista?
            // Tem traço no nome (ex: 000204 - FRANCLIN) e a coluna da Linha (B) está vazia
            if (colG.includes('-') && colB === '') {
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                const partes = colG.split('-');
                motoristaAtual = {
                    id: partes[0].trim(),
                    nome: partes.slice(1).join('-').trim(),
                    data: colL !== '' ? colL : 'Data não encontrada',
                    viagens: []
                };
            }
            // REGRA 2: É uma viagem válida (Programado)?
            // Tem a linha preenchida (B) e o horário de CheckList preenchido (D)
            // Isso ignora a linha vazia do "Realizado"
            else if (motoristaAtual && colB !== '' && lerCelula(4) !== '') {
                motoristaAtual.viagens.push({
                    linha: colB,
                    veiculo: lerCelula(3),       // C (ex: 23910)
                    checkList1: lerCelula(4),    // D (ex: 00:20)
                    deslocamento1: lerCelula(5), // E (ex: 00:30)
                    pontoInicial: lerCelula(6),  // F (ex: 01:00)
                    pontoFinal: colG,            // G (ex: 06:00) - Na viagem, G é horário!
                    deslocamento2: lerCelula(9), // I (ex: 06:30)
                    checkList2: lerCelula(10)    // J (ex: 06:40)
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
