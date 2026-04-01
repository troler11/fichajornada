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
        
        // Guardam os dados soltos até acharmos o dono (o motorista)
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

            let celulaNomeEncontrada = '';
            let dataEncontrada = '';

            // --------------------------------------------------------
            // REGRA 0: O SCANNER (Procura os dados soltos em qualquer coluna de 1 a 20)
            // --------------------------------------------------------
            for (let i = 1; i <= 20; i++) {
                const celulaBruta = lerCelula(i);
                if (!celulaBruta) continue;
                
                const valorUpper = celulaBruta.toUpperCase();

                // 1. Acha a OS (Ex: 01042026011001-156)
                if (/\d{10,}-\d{2,}/.test(valorUpper)) {
                    ultimaOSEncontrada = celulaBruta;
                }
                
                // 2. Acha a Empresa (Ex: VIAÇÃO MIMO LTDA, EMPRESA 811)
                else if (valorUpper.includes('VIAÇÃO') || valorUpper.includes('LTDA') || /^EMPRESA\s+\d+/.test(valorUpper)) {
                    ultimaEmpresaEncontrada = celulaBruta;
                }

                // 3. Acha a Filial (só preenche se o motorista já foi criado)
                else if (valorUpper.startsWith('FILIAL')) {
                    if (motoristaAtual && motoristaAtual.viagens.length === 0) {
                        motoristaAtual.filial = celulaBruta;
                    }
                }

                // 4. Acha a Função (Ex: MOTORISTA ONIBUS)
                else if (valorUpper.includes('MOTORISTA')) {
                    if (motoristaAtual && motoristaAtual.viagens.length === 0) {
                        motoristaAtual.funcao = celulaBruta;
                    }
                }

                // 5. Acha o Nome do Motorista (Padrão ID - Nome, ex: 000204 - FRANCLIN GAMA)
                else if (/^\d{4,8}\s*-\s*[A-ZÀ-Ÿ]/.test(valorUpper)) {
                    celulaNomeEncontrada = celulaBruta;
                }

                // 6. Acha a Data (Ex: 01/04/2026)
                else if (/^\d{2}\/\d{2}\/\d{4}$/.test(valorUpper)) {
                    dataEncontrada = celulaBruta;
                }
            }

            const colB = lerCelula(2);  // Continua sendo a Linha da viagem (Ex: B8170853EP)

            // --------------------------------------------------------
            // REGRA 1: CRIANDO O MOTORISTA
            // --------------------------------------------------------
            // Se o Scanner achou o nome e a coluna B tá vazia (não é viagem)
            if (celulaNomeEncontrada !== '' && colB === '') {
                // Salva o motorista anterior antes de criar o novo
                if (motoristaAtual) motoristas.push(motoristaAtual);
                
                const partes = celulaNomeEncontrada.split('-');
                motoristaAtual = {
                    os: ultimaOSEncontrada,
                    empresa: ultimaEmpresaEncontrada,
                    filial: '', // Será preenchido na próxima linha pelo Scanner
                    funcao: '', // Será preenchido na próxima linha pelo Scanner
                    id: partes[0].trim(),
                    nome: partes.slice(1).join('-').trim(),
                    data: dataEncontrada !== '' ? dataEncontrada : 'Data não encontrada',
                    viagens: []
                };
                
                // Limpa as variáveis para não misturar dados
                ultimaOSEncontrada = '';
                ultimaEmpresaEncontrada = '';
            }

            // --------------------------------------------------------
            // REGRA 2: ADICIONANDO AS VIAGENS
            // --------------------------------------------------------
            // Mantive os números exatos do seu código original (lerCelula(3), lerCelula(4), etc) 
            // pois a tabela de viagens parece estar estática.
            else if (motoristaAtual && colB !== '' && lerCelula(4) !== '') {
                motoristaAtual.viagens.push({
                    linha: colB,
                    veiculo: lerCelula(3),
                    checkList1: lerCelula(4),
                    deslocamento1: lerCelula(5),
                    pontoInicial: lerCelula(6),
                    pontoFinal: lerCelula(7), 
                    deslocamento2: lerCelula(9),
                    checkList2: lerCelula(10)
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
