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
        
        // MUDANÇA 1: Usar um objeto para agrupar por ID
        const mapaMotoristas: Record<string, any> = {};
        let motoristaAtual: any = null;
        
        let ultimaOSEncontrada = '';
        let ultimaEmpresaEncontrada = '';

        worksheet.eachRow((row, rowNumber) => {
            const lerCelula = (num: number) => {
                const celula = row.getCell(num);
                if (!celula || celula.value === null || celula.value === undefined) return '';
                if (typeof celula.value === 'object') {
                    if ('richText' in celula.value) return celula.value.richText.map((rt: any) => rt.text).join('').trim();
                    if ('result' in celula.value) return String(celula.value.result).trim();
                    if (celula.value instanceof Date) {
                        // Converte a data do Excel preservando a hora exata se existir
                        const dia = String(celula.value.getDate()).padStart(2, '0');
                        const mes = String(celula.value.getMonth() + 1).padStart(2, '0');
                        const ano = celula.value.getFullYear();
                        const hora = String(celula.value.getHours()).padStart(2, '0');
                        const min = String(celula.value.getMinutes()).padStart(2, '0');
                        const sec = String(celula.value.getSeconds()).padStart(2, '0');
                        if (hora === '00' && min === '00' && sec === '00') return `${dia}/${mes}/${ano}`;
                        return `${dia}/${mes}/${ano} ${hora}:${min}:${sec}`;
                    }
                }
                return String(celula.value).trim();
            };

            let celulaNomeEncontrada = '';
            let dataEncontrada = '';

            // REGRA 0: O SCANNER
            for (let i = 1; i <= 20; i++) {
                const celulaBruta = lerCelula(i);
                if (!celulaBruta) continue;
                
                const valorUpper = celulaBruta.toUpperCase();

                if (/\d{10,}-\d{2,}/.test(valorUpper)) ultimaOSEncontrada = celulaBruta;
                else if (valorUpper.includes('VIAÇÃO') || valorUpper.includes('LTDA') || /^EMPRESA\s+\d+/.test(valorUpper)) ultimaEmpresaEncontrada = celulaBruta;
                else if (valorUpper.startsWith('FILIAL') && motoristaAtual && motoristaAtual.viagens.length === 0) motoristaAtual.filial = celulaBruta;
                else if (valorUpper.includes('MOTORISTA') && motoristaAtual && motoristaAtual.viagens.length === 0) motoristaAtual.funcao = celulaBruta;
                else if (/^\d{4,8}\s*-\s*[A-ZÀ-Ÿ]/.test(valorUpper)) celulaNomeEncontrada = celulaBruta;
                else if (/^\d{2}\/\d{2}\/\d{4}$/.test(valorUpper)) dataEncontrada = celulaBruta;
            }

            const colB = lerCelula(2); 

            // REGRA 1: CRIANDO OU RECUPERANDO O MOTORISTA
            if (celulaNomeEncontrada !== '' && colB === '') {
                const partes = celulaNomeEncontrada.split('-');
                const idMotorista = partes[0].trim();
                const nomeMotorista = partes.slice(1).join('-').trim();

                // Se o motorista ainda não existe no nosso "mapa", criamos ele.
                // Se ele já existir, nós apenas ignoramos a criação e "puxamos" a ficha dele de volta
                if (!mapaMotoristas[idMotorista]) {
                    mapaMotoristas[idMotorista] = {
                        os: ultimaOSEncontrada,
                        empresa: ultimaEmpresaEncontrada,
                        filial: '', 
                        funcao: '', 
                        id: idMotorista,
                        nome: nomeMotorista,
                        data: dataEncontrada !== '' ? dataEncontrada : 'Data não encontrada',
                        viagens: []
                    };
                }
                
                // O motorista atual passa a ser o do mapa (seja ele novo ou recuperado lá de cima)
                motoristaAtual = mapaMotoristas[idMotorista];
                
                ultimaOSEncontrada = '';
                ultimaEmpresaEncontrada = '';
            }

            // REGRA 2: ADICIONANDO AS VIAGENS
            else if (motoristaAtual && colB !== '' && lerCelula(4) !== '') {
                // MUDANÇA 2: Capturar a data e hora completas da linha da viagem para ordenação
                let timestampViagem = '';
                for (let i = 1; i <= 20; i++) {
                    const cel = lerCelula(i);
                    // Procura o padrão: DD/MM/YYYY HH:MM:SS
                    if (/\d{2}\/\d{2}\/\d{4}\s\d{2}:\d{2}:\d{2}/.test(cel)) {
                        timestampViagem = cel;
                        break;
                    }
                }

                motoristaAtual.viagens.push({
                    linha: colB,
                    veiculo: lerCelula(3),
                    checkList1: lerCelula(4),
                    deslocamento1: lerCelula(5),
                    pontoInicial: lerCelula(6),
                    pontoFinal: lerCelula(7), 
                    deslocamento2: lerCelula(9),
                    checkList2: lerCelula(10),
                    // Salvamos esse dado escondido só para o backend conseguir ordenar
                    _timestampOrdenacao: timestampViagem || lerCelula(4) 
                });
            }
        });

        // MUDANÇA 3: Transforma o mapa num Array normal e Ordena as viagens por horário
        const listaFinalMotoristas = Object.values(mapaMotoristas);

        listaFinalMotoristas.forEach(mot => {
            mot.viagens.sort((a: any, b: any) => {
                // Função para converter "DD/MM/YYYY HH:MM:SS" em um número gigantesco para o Javascript comparar
                // Ex: "01/04/2026 05:10:10" vira 20260401051010
                const converterParaNumero = (str: string) => {
                    const matchCompleto = str.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
                    if (matchCompleto) {
                        return Number(`${matchCompleto[3]}${matchCompleto[2]}${matchCompleto[1]}${matchCompleto[4]}${matchCompleto[5]}${matchCompleto[6]}`);
                    }
                    
                    // Fallback: Se não achar a data completa, converte o horário do Checklist1 (Ex: "04:00" vira 400)
                    const matchHora = str.match(/(\d{2}):(\d{2})/);
                    if (matchHora) {
                        return Number(`${matchHora[1]}${matchHora[2]}00`);
                    }
                    return 0;
                };

                const valorA = converterParaNumero(a._timestampOrdenacao);
                const valorB = converterParaNumero(b._timestampOrdenacao);

                return valorA - valorB;
            });
        });

        // Envia para o frontend a lista perfeitamente agrupada e ordenada
        res.json(listaFinalMotoristas);

    } catch (error) {
        console.error("Erro ao processar planilha:", error);
        res.status(500).send('Erro interno ao ler a planilha.');
    }
});

app.listen(3000, () => console.log('Servidor rodando na porta 3000'));
