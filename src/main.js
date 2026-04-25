const { chromium } = require("playwright");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

// Você pode mudar o número da página abaixo
const START_PAGE = 4000
const TOTAL_PAGES = 2000; // A quantidade TOTAL de páginas a processar (não a última página, mas quantas ler)
const LOGIN_URL = "https://central156.fortaleza.ce.gov.br/coepa/admin/formulario_login.php#paralogin";
const LOGIN_USERNAME = "Eduarda Cândido";
const LOGIN_PASSWORD = "010104";

const SELECTOR_USERNAME = 'input[name="usuario"]';
const SELECTOR_PASSWORD = 'input[name="senha"]';
const SELECTOR_SUBMIT = 'input[type="submit"]';

const BASE_URL = "https://central156.fortaleza.ce.gov.br/coepa/admin/agendamento/pagination.php?id=1&ucad=10";

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function fazerLogin(page) {
    console.log("🔐 Fazendo login...");

    try {
        await page.goto(LOGIN_URL, { waitUntil: "networkidle" });
        console.log("✓ Página de login carregada");

        await page.fill(SELECTOR_USERNAME, LOGIN_USERNAME);
        await page.fill(SELECTOR_PASSWORD, LOGIN_PASSWORD);
        console.log("✓ Credenciais preenchidas");

        await page.click(SELECTOR_SUBMIT);

        await page.waitForLoadState("networkidle");
        console.log("✓ Login realizado com sucesso");

        await sleep(1000 + Math.floor(Math.random() * 1000));

        return true;
    } catch (err) {
        console.error("❌ Erro ao fazer login:", err.message);
        return false;
    }
}

async function extrairDadosDaPagina(page, pageNumber) {
    const url = `${BASE_URL}&page=${pageNumber}`;
    console.log(`[${pageNumber}] Navegando para: ${url}`);

    try {
        await page.goto(url, { waitUntil: "networkidle" });

        const html = await page.content();
        console.log(`[${pageNumber}] Página carregada (${html.length} caracteres)`);

        const $ = cheerio.load(html);
        const rows = $("#example1 tbody tr");
        console.log(`[${pageNumber}] ${rows.length} linhas encontradas na tabela`);

        const dados = [];

        rows.each((i, el) => {
            const colunas = $(el).find("td");
            if (colunas.length < 6) {
                console.log(`[${pageNumber}] Linha ${i + 1}: menos de 6 colunas, ignorando...`);
                return;
            }

            const nome = $(colunas[0]).find("strong").text().trim();
            const endereco = $(colunas[0]).find(".endereco").text().trim();
            const telefone = $(colunas[1]).text().trim();
            const animal = $(colunas[2]).text().trim();
            const qtd = $(colunas[3]).text().trim();
            const data = $(colunas[4]).text().trim();
            const status = $(colunas[5]).text().trim();

            let botao = null;
            if (colunas.length > 7) {
                botao = $(colunas[7]).find("button").attr("onClick");
            }

            let id = null;
            if (botao) {
                const match = botao.match(/id=(\d+)/);
                if (match) id = match[1];
            }

            console.log(`[${pageNumber}] Linha ${i + 1}:`, {
                nome,
                endereco,
                telefone,
                animal,
                qtd,
                data,
                status,
                id
            });

            dados.push({
                nome,
                endereco,
                telefone,
                animal,
                qtd,
                data,
                status,
                id
            });
        });

        return dados;
    } catch (err) {
        console.error(`[${pageNumber}] Erro ao extrair dados:`, err.message);
        return [];
    }
}

async function main() {
    console.log("🚀 Iniciando extração de dados...");
    console.log(`🔢 Página inicial: ${START_PAGE}`);
    console.log(`🔢 Total de páginas a processar: ${TOTAL_PAGES}\n`);

    // Usaremos um Set para manter o registro do par (nome, telefone) já extraído
    const registrosUnicos = new Set();

    const browser = await chromium.launch({
        headless: false,
        slowMo: 100
    });

    const context = await browser.newContext({
        viewport: { width: 1920, height: 1080 },
        userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    });

    const page = await context.newPage();

    try {
        // Faz login
        const loginSucesso = await fazerLogin(page);
        if (!loginSucesso) {
            console.error("❌ Falha no login. Encerrando...");
            return;
        }

        // Prepara o Excel
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Agendamentos");

        sheet.columns = [
            { header: "Nome", key: "nome", width: 30 },
            { header: "Endereço", key: "endereco", width: 40 },
            { header: "Telefone", key: "telefone", width: 20 },
            { header: "Animal", key: "animal", width: 15 },
            { header: "Qtd", key: "qtd", width: 10 },
            { header: "Solicitação", key: "data", width: 20 },
            { header: "Status", key: "status", width: 20 },
            { header: "ID", key: "id", width: 10 }
        ];

        let totalLinhas = 0;

        // Variável para saber se ao menos uma linha foi inserida (para tratar mensagem de "nenhuma linha" ao final)
        let algumaLinhaInserida = false;
        // Controla se o arquivo já foi salvo alguma vez
        let arquivoSalvoAoMenosUmaVez = false;

        // Extrai dados de cada página
        for (let pageNum = START_PAGE; pageNum < START_PAGE + TOTAL_PAGES; pageNum++) {
            console.log(`\n📄 Processando página ${pageNum} (de ${START_PAGE} até ${START_PAGE + TOTAL_PAGES - 1})`);

            const dados = await extrairDadosDaPagina(page, pageNum);

            let linhasInseridas = 0;

            dados.forEach(dado => {
                // Normalização simples para evitar problemas de espaços e caixa
                const chaveUnica = `${dado.nome.trim().toLowerCase()}|${dado.telefone.trim()}`;
                if (!registrosUnicos.has(chaveUnica)) {
                    sheet.addRow(dado);
                    registrosUnicos.add(chaveUnica);
                    totalLinhas++;
                    linhasInseridas++;
                    algumaLinhaInserida = true;
                } else {
                    // Uncomment if you want to log duplicates found
                    // console.log(`🔄 Duplicado encontrado para: ${dado.nome} - ${dado.telefone}, ignorando.`);
                }
            });

            console.log(`[${pageNum}] ${linhasInseridas} linhas inseridas (descartando duplicados)`);

            // Salva no Excel a cada 25 páginas OU na última página
            const paginasProcessadas = pageNum - START_PAGE + 1;
            const ultimaPagina = pageNum === (START_PAGE + TOTAL_PAGES - 1);
            if (paginasProcessadas % 25 === 0 || ultimaPagina) {
                try {
                    await workbook.xlsx.writeFile("agendamentos.xlsx");
                    arquivoSalvoAoMenosUmaVez = true;
                    console.log(`💾 Progresso salvo após página ${pageNum}!`);
                } catch (e) {
                    console.error(`❌ Erro ao salvar o Excel na página ${pageNum}:`, e && e.message);
                }
            }

            if (!ultimaPagina) {
                await sleep(1000 + Math.floor(Math.random() * 1000));
            }
        }

        if (!algumaLinhaInserida) {
            console.warn("\n⚠️ Nenhuma linha foi inserida no Excel. Verifique os dados extraídos e os seletores usados.");
        } else if (!arquivoSalvoAoMenosUmaVez) {
            // Se não salvou por algum motivo, salva ao menos agora
            try {
                await workbook.xlsx.writeFile("agendamentos.xlsx");
                console.log(`\n✔ Excel gerado com sucesso ao final! Total de ${totalLinhas} linhas (após remoção de duplicados)`);
            } catch (e) {
                console.error("❌ Erro ao salvar Excel ao final:", e && e.message);
            }
        } else {
            console.log(`\n✔ Excel gerado com sucesso! Total de ${totalLinhas} linhas (após remoção de duplicados)`);
        }

    } catch (err) {
        console.error("❌ Erro durante a execução:", err);
    } finally {
        await browser.close();
        console.log("\n🔒 Navegador fechado");
    }
}

main().catch(console.error);