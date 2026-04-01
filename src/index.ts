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
        
        // Variáveis temporárias para dados que vêm ANTES da linha do nome
        let ultimaOSEncontrada = '';
        let ultimaEmpresaEncontrada = '';

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

            // --------------------------------------------------------
            // REGRA 0: Scanner de dados flutuantes (OS, Empresa, Filial, Função)
            // --------------------------------------------------------
            for (let i = 1; i <= 20; i++) {
                const celulaBruta = lerCelula(i);
                if (!celulaBruta) continue;
                
                const valorUpper = celulaBruta.toUpperCase();

                // 1. Captura OS (Ex: 01042026011001-156)
                if (/\d{10,}-\d{2,}/.test(valorUpper)) {
                    ultimaOSEncontrada = celulaBruta;
                }
                
                // 2. Captura Empresa (Ex: VIAÇÃO MIMO LTDA, EMPRESA 811)
                // Procura por "VIAÇÃO", "LTDA" ou o padrão "EMPRESA + Número"
                if (valorUpper.includes('VIAÇÃO') || valorUpper.includes('LTDA') || /^EMPRESA\s+\d+/.test(valorUpper)) {
                    ultimaEmpresaEncontrada = celulaBruta;
                }

                // 3. Captura Filial (Ex: FILIAL 8)
                // Se o motorista já foi criado na linha anterior, injetamos a filial nele
                if (valorUpper.startsWith('FILIAL')) {
                    if (motoristaAtual && motoristaAtual.viagens.length === 0) {
                        motoristaAtual.filial = celulaBruta;
                    }
                }

                // 4. Captura Função (Ex: MOTORISTA ONIBUS)
                if (valorUpper.includes('MOTORISTA')) {
                    if (motoristaAtual && motoristaAtual.viagens.length === 0) {
                        motoristaAtual.funcao = celulaBruta;
                    }
                }
            }

            const colB = lerCelula(2);  // Linha
            const colG = lerCelula(7);  // Nome do Motorista OU Ponto Final
            const colL = lerCelula(12); // Data

            // REGRA 1: É o cabeçalho do motorista? (Acha o Nome)
            if (colG.includes('-') && colB === '') {
                // Se já tínhamos um motorista sendo montado, salva ele
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                const partes = colG.split('-');
                motoristaAtual = {
                    os: ultimaOSEncontrada,
                    empresa: ultimaEmpresaEncontrada,
                    filial: '', // Será preenchido pelo Scanner na próxima linha
                    funcao: '', // Será preenchido pelo Scanner na próxima linha
                    id: partes[0].trim(),
                    nome: partes.slice(1).join('-').trim(),
                    data: colL !== '' ? colL : 'Data não encontrada',
                    viagens: []
                };
                
                // Limpa as variáveis temporárias para não repetirem se faltar no próximo
                ultimaOSEncontrada = '';
                ultimaEmpresaEncontrada = '';
            }
            // REGRA 2: É uma viagem válida (Programado)?
            else if (motoristaAtual && colB !== '' && lerCelula(4) !== '') {
                motoristaAtual.viagens.push({
                    linha: colB,
                    veiculo: lerCelula(3),
                    checkList1: lerCelula(4),
                    deslocamento1: lerCelula(5),
                    pontoInicial: lerCelula(6),
                    pontoFinal: colG,
                    deslocamento2: lerCelula(9),
                    checkList2: lerCelula(10)
                });
            }
        });

        // Salva o último motorista quando acabar o loop
        if (motoristaAtual) motoristas.push(motoristaAtual);
        res.json(motoristas);

    } catch (error) {
        console.error("Erro ao processar planilha:", error);
        res.status(500).send('Erro interno ao ler a planilha.');
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
