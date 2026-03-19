/**
 * ============================================================
 * AJ MODAS — ROBÔ DE ESTOQUE + DASHBOARD WEB + EXPORTAÇÃO UPSELLER
 * + MÓDULO DE CONFERÊNCIA (HÍBRIDO: BLING + PLANILHAS UPSELLER/BAGY + PDF)
 * ============================================================
 */

const archiver              = require('archiver'); // 👈 Adicionado para gerar o ZIP
const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode                = require("qrcode-terminal");
const cron                  = require("node-cron");
const axios                 = require("axios");
const fs                    = require("fs");
const path                  = require("path");
const express               = require("express");
const cors                  = require("cors");
const xlsx                  = require("xlsx"); 
const multer                = require("multer"); 
const pdfParse              = require("pdf-parse"); // Adicionado suporte a PDF

const delay = (ms) => new Promise(res => setTimeout(res, ms));
const upload = multer({ storage: multer.memoryStorage() }); 

const CONFIG = {
  bling: {
    clientId     : "0e1b32755e0a5c0d22bfaf0c6e58d7de886e45ee",
    clientSecret : "50bdda91881b0b1981c7529677f5b99e825ffe9f7d7bfcffa422bde94f26",
    tokenFile    : path.join(__dirname, "tokens.json"),
  },
  whatsapp: {
    nomeDoGrupo: "Controle de Estoque"
  },
  limiteMin: 0,
  limiteMax: 10,
  ignorarRefs: [
    "151", "06", "6", "181", "193", "03", "3", "22", "88", "75", "180", 
    "86", "04", "4", "08", "8", "29", "02", "2", "05", "5", "01", "1", 
    "18", "13", "12", "19", "11", "102", "91", "26", "20", "23", "31", 
    "70", "76"
  ]
};

// 📦 MEMÓRIAS VOLÁTEIS DO SERVIDOR
let bancoDadosPlanilha = []; 
let cacheProdutos = null;    
let ultimoCacheHora = 0;     
let cachePedidosRecentes = []; 

function lerTokens() {
  if (!fs.existsSync(CONFIG.bling.tokenFile)) return null;
  return JSON.parse(fs.readFileSync(CONFIG.bling.tokenFile, "utf8"));
}

function salvarTokens(dados) {
  fs.writeFileSync(CONFIG.bling.tokenFile, JSON.stringify(dados, null, 2));
}

async function renovarToken(refreshToken) {
  const credenciais = Buffer.from(`${CONFIG.bling.clientId}:${CONFIG.bling.clientSecret}`).toString("base64");
  const resp = await axios.post(
    "https://www.bling.com.br/Api/v3/oauth/token",
    new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken }),
    { headers: { Authorization: `Basic ${credenciais}`, "Content-Type": "application/x-www-form-urlencoded" } }
  );
  const tokens = {
    access_token : resp.data.access_token,
    refresh_token: resp.data.refresh_token,
    expires_at   : Date.now() + resp.data.expires_in * 1000,
  };
  salvarTokens(tokens);
  return tokens.access_token;
}

async function obterAccessToken() {
  const tokens = lerTokens();
  if (!tokens) throw new Error("⛔ Arquivo tokens.json não encontrado.");
  if (Date.now() >= tokens.expires_at - 120_000) {
    return await renovarToken(tokens.refresh_token);
  }
  return tokens.access_token;
}

async function buscarEstoque(accessToken) {
  const todosProdutos = [];
  let pagina = 1;
  let temMais = true;

  console.log("   [Bling] A iniciar varredura completa do catálogo...");
  do {
    const resp = await axios.get("https://www.bling.com.br/Api/v3/produtos", {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: { pagina, limite: 100, tipo: 'T' } 
    });
    
    const data = resp.data?.data ?? [];
    todosProdutos.push(...data);
    
    if (data.length < 100) temMais = false;
    else pagina++;
    await delay(500); 
  } while (temMais);
  
  const ids = todosProdutos.map(p => p.id).filter(id => id);
  const lotes = [];
  for (let i = 0; i < ids.length; i += 50) lotes.push(ids.slice(i, i + 50));

  const saldos = [];
  for (const lote of lotes) {
    const params = new URLSearchParams();
    for (const id of lote) params.append("idsProdutos[]", id);
    const resp = await axios.get("https://www.bling.com.br/Api/v3/estoques/saldos", {
      headers: { Authorization: `Bearer ${accessToken}` },
      params: params
    });
    saldos.push(...(resp.data?.data ?? []));
    await delay(500); 
  }

  const mapaSaldos = new Map();
  for (const s of saldos) {
    const idProduto = s.produto?.id || s.id;
    if (idProduto) mapaSaldos.set(idProduto, s);
  }

  return todosProdutos.map(p => ({
    id: p.id,
    codigo: p.codigo,
    descricao: p.nome,
    saldoFisicoTotal: mapaSaldos.get(p.id)?.saldoFisicoTotal ?? 0
  }));
}

const app = express();
app.use(cors());
app.use(express.static(__dirname)); 
app.use(express.json());

// ──────────────────────────────────────────────
// 🟢 UPLOAD INTELIGENTE (EXCEL UPSELLER, EXCEL BAGY e PDF)
// ──────────────────────────────────────────────

app.post('/api/checkout/upload-csv', upload.single('arquivo'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ erro: "Nenhum arquivo recebido pelo servidor." });

        const nomeArquivo = req.file.originalname.toLowerCase();
        let novosItens = [];

        // 📝 SE FOR ARQUIVO PDF (Romaneio / UpSeller)
        if (nomeArquivo.endsWith('.pdf')) {
            console.log("📄 Processando PDF do UpSeller...");
            const data = await pdfParse(req.file.buffer);
            
            const linhas = data.text.split('\n').map(l => l.trim()).filter(l => l !== '');
            let ultimoPedido = "";

            for (let i = 0; i < linhas.length; i++) {
                const linha = linhas[i];

                const matchPedido = linha.match(/(?:#)?\b([1-9]\d{13,16}|GSH[A-Z0-9]{10,15})\b/);
                if (matchPedido) {
                    ultimoPedido = matchPedido[1];
                }

                const matchSKU = linha.match(/\b(\d{3,4}-[A-Za-zÀ-ÿ\s]+-[A-Z0-9]{1,4})\b/);
                
                if (matchSKU && ultimoPedido !== "") {
                    let sku = matchSKU[1].trim();
                    let qtd = 1; 

                    let lSeguinte = linhas[i+1] || "";
                    let lAnterior = i > 0 ? linhas[i-1] : "";
                    let lAcima = i > 1 ? linhas[i-2] : "";

                    if (lSeguinte.match(/^\d+$/)) {
                        qtd = parseInt(lSeguinte); 
                    } else if (lAnterior.match(/x\s*(\d+)/)) {
                        qtd = parseInt(lAnterior.match(/x\s*(\d+)/)[1]); 
                    } else if (lAcima.match(/x\s*(\d+)/)) {
                        qtd = parseInt(lAcima.match(/x\s*(\d+)/)[1]); 
                    }

                    novosItens.push({
                        pedido: ultimoPedido,
                        sku: sku,
                        nome: `Peça: ${sku}`,
                        qtd: qtd
                    });
                }
            }
        } 
        // 📊 SE FOR PLANILHA (O Recomendado - Bagy/UpSeller CSV/Excel)
        else {
            console.log("📊 Processando Planilha Excel/CSV com nova inteligência...");
            const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
            const sheetName = workbook.SheetNames[0];
            const rawData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

            if (rawData.length === 0) return res.status(400).json({ erro: "A planilha está vazia." });

            const colunas = Object.keys(rawData[0]);
            
            // 👇 INTELIGÊNCIA MÁXIMA PARA ENCONTRAR COLUNAS 👇
            const normalizar = (texto) => String(texto).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "").toLowerCase();
            
            const acharColunaExata = (nomesPossiveis) => {
                return colunas.find(colReal => {
                    const colLimpa = normalizar(colReal);
                    return nomesPossiveis.some(nome => colLimpa === normalizar(nome) || colLimpa.includes(normalizar(nome)));
                });
            };

            const colPedido = acharColunaExata(['Nº de Pedido da Plataforma', 'code', 'pedido']) || colunas[0];
            const colSku = acharColunaExata(['SKU (Armazém)', 'Order Items__reference', 'sku']) || 'sku';
            const colNome = acharColunaExata(['Nome do Produto', 'Nome do Anúncio', 'Order Items__name', 'nome']) || 'name';
            
            // 🚨 SOLUÇÃO SNIPER PARA A QUANTIDADE
            const colQtd = acharColunaExata(['Qtd. do Produto', 'Order Items__quantity']) || 
                           colunas.find(c => normalizar(c) === 'qtd') || 
                           colunas.find(c => normalizar(c) === 'quantidade') || 
                           'quantity';
            
            const colVar = acharColunaExata(['Variação', 'Order Items__variation']);
            // 👆 FIM DA NOVA LÓGICA 👆

            novosItens = rawData.map(row => {
                let pedido = String(row[colPedido] || "").trim();
                let sku = String(row[colSku] || "").trim();
                let nome = String(row[colNome] || "").trim();
                let variacao = colVar ? String(row[colVar] || "").trim() : "";
                
                let qtd = parseInt(row[colQtd]);

                if (variacao && variacao !== "undefined") nome = `${nome} (${variacao})`;
                
                if (isNaN(qtd) || qtd <= 0) qtd = 1; 

                return { pedido, sku, nome, qtd };
            }).filter(item => item.pedido !== "" && item.sku !== "" && item.sku !== "undefined"); 
        }

        bancoDadosPlanilha = [...bancoDadosPlanilha, ...novosItens];

        console.log(`✅ [Checkout] Arquivo LIDO! A memória agora tem ${bancoDadosPlanilha.length} itens prontos.`);
        res.json({ sucesso: true, total: novosItens.length, memoriaTotal: bancoDadosPlanilha.length });
        
    } catch (e) {
        console.error("❌ Erro fatal ao processar arquivo:", e);
        res.status(500).json({ erro: "O servidor não conseguiu ler este arquivo. Verifique o formato." });
    }
});


// 👇 ROTA: SINCRONIZAR PEDIDOS RECENTES (Últimos 60 dias) 👇
app.get('/api/checkout/sincronizar', async (req, res) => {
    try {
        const token = await obterAccessToken();
        console.log("\n🔄 [Checkout] Sincronizando pedidos RECENTES com o Bling...");
        let tempPedidos = [];
        
        const dataAtras = new Date();
        dataAtras.setDate(dataAtras.getDate() - 60);
        const ano = dataAtras.getFullYear();
        const mes = String(dataAtras.getMonth() + 1).padStart(2, '0');
        const dia = String(dataAtras.getDate()).padStart(2, '0');
        const dataInicial = `${ano}-${mes}-${dia}`;
        
        for (let pagina = 1; pagina <= 10; pagina++) {
            try {
                const url = `https://www.bling.com.br/Api/v3/pedidos/vendas?dataInicial=${dataInicial}&pagina=${pagina}&limite=100`;
                const resp = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
                
                if (resp.data && resp.data.data && resp.data.data.length > 0) {
                    tempPedidos.push(...resp.data.data);
                } else {
                    break;
                }
                await delay(350);
            } catch (e) { break; }
        }
        
        cachePedidosRecentes = tempPedidos;
        console.log(`✅ [Checkout] ${cachePedidosRecentes.length} pedidos em memória prontos para busca instantânea!`);
        res.json({ sucesso: true, total: cachePedidosRecentes.length });
    } catch (e) {
        res.status(500).json({ erro: "Falha ao sincronizar pedidos" });
    }
});

// 🟢 BUSCAR PEDIDO 
app.get('/api/checkout/pedido/:numero', async (req, res) => {
    const numero = req.params.numero.trim();
    console.log(`\n🔍 [Checkout] Solicitada busca pelo pedido: ${numero}`);
    
    // 1. Tenta achar na Planilha 
    const itensCSV = bancoDadosPlanilha.filter(i => i.pedido === numero);
    if (itensCSV.length > 0) {
        console.log(`✅ [Checkout] Pedido ${numero} encontrado na Planilha!`);
        return res.json({ origem: 'PLANILHA', numero: numero, itens: itensCSV.map(i => ({ sku: i.sku, nome: i.nome, esperado: i.qtd, conferido: 0 })) });
    }

    try {
        const token = await obterAccessToken();
        let pedidoId = null;
        let numeroBling = null;
        
        // 2. Busca na Memória que foi Sincronizada
        const pedidoEmMemoria = cachePedidosRecentes.find(p => String(p.numero) === numero || String(p.numeroLoja) === numero);
        
        if (pedidoEmMemoria) {
            pedidoId = pedidoEmMemoria.id;
            numeroBling = pedidoEmMemoria.numero;
            console.log(`⚡ [Checkout] Encontrado instantaneamente na Memória! ID: ${pedidoId}`);
        } else {
            console.log(`📡 [Checkout] Não estava na memória. Tentativa de emergência na API...`);
            try {
                const resp = await axios.get(`https://www.bling.com.br/Api/v3/pedidos/vendas?numeroLoja=${numero}`, { headers: { Authorization: `Bearer ${token}` } });
                if(resp.data && resp.data.data && resp.data.data.length > 0) {
                    const found = resp.data.data.find(p => String(p.numeroLoja) === numero);
                    if (found) {
                        pedidoId = found.id;
                        numeroBling = found.numero;
                        console.log(`🎯 Encontrado via emergência! ID: ${pedidoId}`);
                    }
                }
            } catch(e) {}
        }

        if (!pedidoId) {
            console.log(`❌ [Checkout] Pedido ${numero} não existe.`);
            return res.status(404).json({ erro: `Pedido ${numero} não localizado nas vendas recentes.` });
        }

        console.log(`📦 Baixando peças do pedido...`);
        const respDetalhes = await axios.get(`https://www.bling.com.br/Api/v3/pedidos/vendas/${pedidoId}`, { headers: { Authorization: `Bearer ${token}` } });
        
        const itensBling = respDetalhes.data.data.itens.map(i => ({
            sku: i.codigo || i.produto?.codigo || "S/COD",
            nome: i.descricao || "Produto Sem Nome",
            esperado: Math.round(i.quantidade),
            conferido: 0
        }));

        console.log(`✅ [Checkout] Sucesso!`);
        res.json({ origem: 'BLING', id: pedidoId, numero: numeroBling, numeroLoja: numero, itens: itensBling });

    } catch (e) {
        res.status(500).json({ erro: "Erro de comunicação ao buscar pedido no Bling." });
    }
});

// 👇 ROTA DE FINALIZAR PEDIDO 👇
app.post('/api/checkout/finalizar', async (req, res) => {
    const { origem, id, itens, numero } = req.body;
    try {
        const token = await obterAccessToken();

        if (origem === 'BLING') {
            console.log(`\n⏳ Injetando Vendedor (SITE) e Loja (SEDE) no pedido ${numero}...`);
            
            const respPedido = await axios.get(`https://www.bling.com.br/Api/v3/pedidos/vendas/${id}`, {
                headers: { Authorization: `Bearer ${token}` }
            });
            
            let dadosPedido = respPedido.data.data;
            
            dadosPedido.loja = { id: 205344151 };       // ID da Loja SEDE
            dadosPedido.vendedor = { id: 15596386514 }; // ID do Vendedor SITE
            
            try {
                await axios.put(`https://www.bling.com.br/Api/v3/pedidos/vendas/${id}`, dadosPedido, {
                    headers: { Authorization: `Bearer ${token}` }
                });
                console.log(`✅ Loja e Vendedor atualizados com sucesso no Bling!`);
            } catch (errPut) {
                console.error(`⚠️ Erro ao injetar Loja/Vendedor (mas o pacote será finalizado mesmo assim).`);
            }

            await axios.patch(`https://www.bling.com.br/Api/v3/pedidos/vendas/${id}/situacoes/9`, {}, {
                headers: { Authorization: `Bearer ${token}` }
            });
            console.log(`✅ [Checkout] Pedido Bling ${numero} marcado como Atendido!`);
            
        } else {
            const depositoId = 14887498122; 
            for (const item of itens) {
                try {
                    const respProd = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${item.sku}`, { headers: { Authorization: `Bearer ${token}` }});
                    if (respProd.data.data.length > 0) {
                        const prodId = respProd.data.data[0].id;
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: prodId },
                            deposito: { id: depositoId },
                            operacao: "S",
                            quantidade: item.esperado,
                            observacoes: `Baixa via Checkout de Expedição. Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` } });
                    }
                } catch (errItem) { console.error(`⚠️ Erro ao dar baixa no SKU ${item.sku}`); }
            }
            console.log(`✅ [Checkout] Pedido Planilha/UpSeller ${numero} finalizado. Baixas efetuadas.`);
        }
        res.json({ sucesso: true });
    } catch (e) {
        console.error(e);
        res.status(500).json({ erro: "Erro ao finalizar pedido no servidor." });
    }
});


// ──────────────────────────────────────────────
// 🟢 EXPORTAÇÃO UPSELLER (MÁSCARA DE ESTOQUE + ZIP)
// ──────────────────────────────────────────────
app.get('/api/exportar-upseller', async (req, res) => {
    try {
        const baseFake = parseInt(req.query.base) || 2000;
        console.log(`\n📦 [UpSeller] Gerando ZIP com Base Fake: +${baseFake}...`);
        
        const token = await obterAccessToken();
        const produtos = await buscarEstoque(token);

        if (!produtos || produtos.length === 0) {
            return res.status(404).json({ erro: "Nenhum produto encontrado no Bling para exportar." });
        }

        const linhasPlanilha = [];
        let logConteudo = `====================================================\n`;
        logConteudo += `📊 RELATÓRIO DE EXPORTAÇÃO UPSELLER (ESTOQUE ESPELHO)\n`;
        logConteudo += `Data: ${new Date().toLocaleString('pt-BR')}\n`;
        logConteudo += `Base Adicional: +${baseFake} unidades\n`;
        logConteudo += `====================================================\n\n`;

        let qtdSucesso = 0;
        let qtdZerado = 0;
        let ignorados = 0;

        // Nomes exatos das colunas da UpSeller
        const colSKU = "SKU*";
        const colEstoqueBaixo = "Estoque Baixo\n(Não será atualizado se não for preenchido)";
        const colQtdTotal = "Qtd. Total Atualizado\n(Não será atualizado se não for preenchido)";
        const colCusto = "Custo Médio Atualizado\n(Não será atualizado se não for preenchido)";

        produtos.forEach(p => {
            const sku = p.codigo;
            const nome = p.descricao || "Sem Nome";
            const quantidadeReal = parseInt(p.saldoFisicoTotal) || 0;

            if (!sku || sku === "S/COD") {
                ignorados++;
                logConteudo += `[IGNORADO] Produto sem SKU: ${nome}\n`;
                return;
            }

            let quantidadeUpSeller = 0;
            
            if (quantidadeReal > 0) {
                quantidadeUpSeller = baseFake + quantidadeReal;
                qtdSucesso++;
            } else {
                qtdZerado++;
                logConteudo += `[ALERTA] Estoque ZERO para o SKU: ${sku} (${nome}) - Enviado como 0.\n`;
            }

            linhasPlanilha.push({
                [colSKU]: sku,
                [colEstoqueBaixo]: "",
                [colQtdTotal]: quantidadeUpSeller,
                [colCusto]: ""
            });
        });

        logConteudo += `\n====================================================\n`;
        logConteudo += `RESUMO DA EXPORTAÇÃO:\n`;
        logConteudo += `- SKUs com Estoque Mascarado (+${baseFake}): ${qtdSucesso}\n`;
        logConteudo += `- SKUs Zerados (Enviados como 0): ${qtdZerado}\n`;
        logConteudo += `- SKUs Ignorados (Sem Código): ${ignorados}\n`;
        logConteudo += `====================================================\n`;

        const worksheet = xlsx.utils.json_to_sheet(linhasPlanilha);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
        const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        const dataAtual = new Date().toISOString().slice(0,10);
        res.setHeader('Content-Disposition', `attachment; filename="UpSeller_Exportacao_${dataAtual}.zip"`);
        res.setHeader('Content-Type', 'application/zip');

        const archive = archiver('zip', { zlib: { level: 9 } });
        
        archive.on('error', function(err) { throw err; });
        archive.pipe(res);

        archive.append(excelBuffer, { name: `Update_warehouse_SKU_${dataAtual}.xlsx` });
        archive.append(logConteudo, { name: `Relatorio_Seguranca.txt` });

        await archive.finalize();
        console.log(`✅ [UpSeller] ZIP gerado e enviado com sucesso!`);

    } catch (e) {
        console.error("❌ Erro ao exportar ZIP UpSeller:", e.message);
        res.status(500).json({ erro: "Erro interno ao gerar o pacote ZIP." });
    }
});


// ──────────────────────────────────────────────
// RESTANTE DO CÓDIGO (DASHBOARD / WMS / WHATSAPP)
// ──────────────────────────────────────────────

app.get('/api/produtos', async (req, res) => {
  try {
    const trintaMinutos = 1800000;
    if (cacheProdutos && (Date.now() - ultimoCacheHora < trintaMinutos)) {
        return res.json(cacheProdutos); 
    }

    const token = await obterAccessToken();
    const produtos = await buscarEstoque(token);
    
    cacheProdutos = produtos;
    ultimoCacheHora = Date.now();

    res.json(produtos);
  } catch (error) {
    if (cacheProdutos) return res.json(cacheProdutos);
    res.status(500).json({ error: "Erro ao buscar produtos" });
  }
});

// 👇 ROTA ADICIONADA: DISPARO MANUAL DE ALERTA NO WHATSAPP
app.get('/api/disparar-alerta', async (req, res) => {
    try {
        console.log("\n🚨 [WhatsApp] Disparo manual de alerta solicitado via Dashboard!");
        executarVerificacao(); // Dispara o processo em background
        res.json({ sucesso: true, mensagem: "Varredura iniciada no fundo." });
    } catch (error) {
        res.status(500).json({ erro: "Erro ao disparar alerta" });
    }
});

app.get('/api/wms/produto/:codigo', async (req, res) => {
  try {
    const codigoBipado = req.params.codigo;
    const token = await obterAccessToken();
    const urlBusca = `https://www.bling.com.br/Api/v3/produtos?codigo=${codigoBipado}`;
    const respBusca = await axios.get(urlBusca, { headers: { Authorization: `Bearer ${token}` } });
    
    if (!respBusca.data || !respBusca.data.data || respBusca.data.data.length === 0) {
      return res.status(404).json({ erro: 'Produto não encontrado' });
    }
    const produto = respBusca.data.data[0];
    const urlEstoque = `https://www.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=${produto.id}`;
    const respEstoque = await axios.get(urlEstoque, { headers: { Authorization: `Bearer ${token}` } });
    let estoqueAtual = respEstoque.data?.data?.[0]?.saldoFisicoTotal || 0;
    res.json({ id: produto.id, codigo: codigoBipado, nome: produto.nome, estoqueAtual: estoqueAtual, fotoUrl: produto.imagemURL || "" });
  } catch (error) {
    res.status(500).json({ erro: 'Erro interno' });
  }
});

app.post('/api/wms/entrada', async (req, res) => {
  try {
    const { idProduto, quantidade, operacao } = req.body;
    const token = await obterAccessToken();
    const depositoId = 14887498122; 
    const tipoOperacao = operacao === 'S' ? 'S' : 'E';
    await axios.post("https://www.bling.com.br/Api/v3/estoques", {
      produto: { id: parseInt(idProduto) },
      deposito: { id: depositoId },
      operacao: tipoOperacao, 
      quantidade: parseFloat(quantidade),
      observacoes: tipoOperacao === 'E' ? "Entrada via WMS Local" : "Saída/Correção via WMS Local"
    }, { headers: { Authorization: `Bearer ${token}` } });
    res.json({ sucesso: true });
  } catch (error) {
    res.status(500).json({ erro: 'Erro ao salvar no Bling' });
  }
});

// Descobre o IP local da máquina na rede
const os = require("os");
function obterIPLocal() {
  const interfaces = os.networkInterfaces();
  for (const nome of Object.keys(interfaces)) {
    for (const iface of interfaces[nome]) {
      if (iface.family === "IPv4" && !iface.internal) {
        return iface.address;
      }
    }
  }
  return "localhost";
}

app.listen(3000, "0.0.0.0", () => {
  const ip = obterIPLocal();
  console.log("======================================================");
  console.log("   🏢 SISTEMA AJ MODAS - SERVIDOR INICIADO!");
  console.log("======================================================");
  console.log("");
  console.log("   📌 ACESSO LOCAL (neste PC):");
  console.log("   🌐 Dashboard: http://localhost:3000/dashboard.html");
  console.log("   📦 WMS:       http://localhost:3000/wms.html");
  console.log("   🛒 Checkout:  http://localhost:3000/checkout.html");
  console.log("");
  console.log("   📌 ACESSO PELA REDE (outros PCs da empresa):");
  console.log(`   🌐 Dashboard: http://${ip}:3000/dashboard.html`);
  console.log(`   📦 WMS:       http://${ip}:3000/wms.html`);
  console.log(`   🛒 Checkout:  http://${ip}:3000/checkout.html`);
  console.log("");
  console.log("   ⚠️  Passe esses links acima para seus colaboradores!");
  console.log("======================================================");
});

// WHATSAPP
const wppClient = new Client({ authStrategy: new LocalAuth(), puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] } });
wppClient.on("qr", (qr) => { qrcode.generate(qr, { small: true }); });
wppClient.on("ready", () => { console.log("✅ [Robô] WhatsApp conectado!"); });
wppClient.initialize();

function filtrarEmRisco(produtos) {
  return produtos.filter(p => {
    const emRisco = p.saldoFisicoTotal >= CONFIG.limiteMin && p.saldoFisicoTotal <= CONFIG.limiteMax;
    if (!emRisco) return false;
    let ref = "";
    let nomeLimpo = p.descricao || "";
    const regexRef = /^(\d+)\s*[-_]\s*/;
    const matchRef = nomeLimpo.match(regexRef);
    if (matchRef) ref = matchRef[1];
    else ref = p.codigo ? p.codigo.split(/[-_]/)[0] : "";
    if (CONFIG.ignorarRefs.includes(ref)) return false;
    return true;
  });
}

function formatarProdutoIndividual(p) {
  let cor = "-", tam = "-";
  let nomeLimpo = p.descricao || "Sem descrição";
  const regexCor = /COR:\s*([^;]+)/i;
  const matchCor = nomeLimpo.match(regexCor);
  if (matchCor) { cor = matchCor[1].trim(); nomeLimpo = nomeLimpo.replace(regexCor, ""); }
  const regexTam = /TAM(?:ANHO)?:\s*([^;]+)/i;
  const matchTam = nomeLimpo.match(regexTam);
  if (matchTam) { tam = matchTam[1].trim(); nomeLimpo = nomeLimpo.replace(regexTam, ""); }
  nomeLimpo = nomeLimpo.replace(/;/g, "").replace(/\s{2,}/g, " ").trim();
  let linhaEstoque = p.saldoFisicoTotal === 0 ? `🛑 *ESTOQUE ZERADO* ➔ Prioridade!` : `🟡 *Estoque Baixo:* Restam ${p.saldoFisicoTotal} un.`;
  return `📦 *SKU:* ${p.codigo || "S/COD"}\n🏷️ *Produto:* ${nomeLimpo}\n🎨 *Cor:* ${cor}  |  📏 *Tam:* ${tam}\n${linhaEstoque}`;
}

async function executarVerificacao() {
  try {
    const token = await obterAccessToken();
    const produtos = await buscarEstoque(token);
    let emRisco = filtrarEmRisco(produtos);

    if (emRisco.length === 0) return;

    emRisco.sort((a, b) => a.saldoFisicoTotal - b.saldoFisicoTotal);

    const qtdZerados = emRisco.filter(p => p.saldoFisicoTotal === 0).length;
    const qtdBaixos = emRisco.length - qtdZerados;

    const conversas = await wppClient.getChats();
    const grupoEncontrado = conversas.find(chat => chat.isGroup && chat.name === CONFIG.whatsapp.nomeDoGrupo);
    if (!grupoEncontrado) return;

    await wppClient.sendMessage(grupoEncontrado.id._serialized, 
        `📊 *RESUMO DE ESTOQUE*\n\n` +
        `Identificamos *${emRisco.length} produtos* que precisam de atenção:\n` +
        `🛑 *${qtdZerados}* totalmente zerados.\n` +
        `🟡 *${qtdBaixos}* com estoque baixo.\n\n` +
        `⏳ _Iniciando o envio fracionado para não sobrecarregar o grupo..._`
    );
    await delay(3000);

    const TAMANHO_LOTE = 20;             
    const TEMPO_PAUSA_MINUTOS = 2;       
    const TEMPO_PAUSA_MS = TEMPO_PAUSA_MINUTOS * 60 * 1000; 

    for (let i = 0; i < emRisco.length; i += TAMANHO_LOTE) {
        const lote = emRisco.slice(i, i + TAMANHO_LOTE);
        const numLote = Math.floor(i / TAMANHO_LOTE) + 1;
        const totalLotes = Math.ceil(emRisco.length / TAMANHO_LOTE);

        await wppClient.sendMessage(grupoEncontrado.id._serialized, `📦 *Enviando Parte ${numLote} de ${totalLotes}* 👇`);
        await delay(2000);

        for (const produto of lote) {
            await wppClient.sendMessage(grupoEncontrado.id._serialized, formatarProdutoIndividual(produto));
            await delay(1500); 
        }

        if (i + TAMANHO_LOTE < emRisco.length) {
            await wppClient.sendMessage(grupoEncontrado.id._serialized, 
                `⏸️ _Pausa estratégica de ${TEMPO_PAUSA_MINUTOS} minutos para leitura. Já volto com os próximos..._`
            );
            await delay(TEMPO_PAUSA_MS); 
        }
    }

    await delay(2000);
    await wppClient.sendMessage(grupoEncontrado.id._serialized, `✅ *Fim da lista de alertas!* Todos os produtos em risco foram informados.`);

  } catch (err) {
    console.error("❌ Erro no WhatsApp:", err.message);
  }
}

cron.schedule("0 * * * *", executarVerificacao, { timezone: "America/Sao_Paulo" });