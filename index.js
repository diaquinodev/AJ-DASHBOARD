/**
 * ============================================================
 * AJ MODAS — ROBÔ DE ESTOQUE + DASHBOARD WEB + EXPORTAÇÃO UPSELLER
 * + MÓDULO DE CONFERÊNCIA (HÍBRIDO: BLING + PLANILHAS UPSELLER/BAGY + PDF)
 * ============================================================
 */

const archiver              = require('archiver'); // 👈 Adicionado para gerar o ZIP
// WhatsApp desativado temporariamente
// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode                = require("qrcode-terminal");
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
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } }); // max 10MB 

// 🔐 Variáveis de ambiente (carrega .env se existir)
try { require('dotenv').config(); } catch (_) { /* dotenv opcional */ }

const CONFIG = {
  bling: {
    clientId     : process.env.BLING_CLIENT_ID     || "0e1b32755e0a5c0d22bfaf0c6e58d7de886e45ee",
    clientSecret : process.env.BLING_CLIENT_SECRET  || "50bdda91881b0b1981c7529677f5b99e825ffe9f7d7bfcffa422bde94f26",
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

// 🔒 PROTEÇÃO CONTRA DUPLICIDADE — Registra operações já processadas
const operacoesFinalizadas = new Map(); // chave: idempotencyKey → { timestamp, resultado }
const EXPIRACAO_OPERACAO_MS = 30 * 60 * 1000; // 30 minutos

function limparOperacoesExpiradas() {
    const agora = Date.now();
    for (const [chave, dados] of operacoesFinalizadas) {
        if (agora - dados.timestamp > EXPIRACAO_OPERACAO_MS) {
            operacoesFinalizadas.delete(chave);
        }
    }
}

// 🏢 DEPÓSITO SEDE — Usado para consulta de saldo, entrada e saída
const DEPOSITO_SEDE_ID = 14887498122;

// Atualiza o saldo no cache local após operação de estoque (E = entrada, S = saída)
function atualizarCacheEstoque(produtoId, quantidade, operacao) {
  if (!cacheProdutos || !produtoId) return;
  const produto = cacheProdutos.find(p => p.id === produtoId);
  if (produto) {
    const delta = operacao === 'E' ? quantidade : -quantidade;
    produto.saldoFisicoTotal = Math.max(0, (produto.saldoFisicoTotal || 0) + delta);
    console.log(`   [Cache] Atualizado estoque local: ID ${produtoId} → ${produto.saldoFisicoTotal} (${operacao === 'E' ? '+' : '-'}${quantidade})`);
  }
}

// 📁 CACHE EM DISCO — Catálogo salvo em arquivo para carregamento instantâneo
const CATALOGO_CACHE_FILE = path.join(__dirname, "catalogo-cache.json");
const UPSELLER_CATALOGO_FILE = path.join(__dirname, "upseller-catalogo.json");
let sincronizandoCatalogo = false;

// 📁 CATÁLOGO UPSELLER — Lista de SKUs reais da UpSeller para matching exato
function lerCatalogoUpSeller() {
  try {
    if (fs.existsSync(UPSELLER_CATALOGO_FILE)) {
      const data = JSON.parse(fs.readFileSync(UPSELLER_CATALOGO_FILE, "utf8"));
      return data;
    }
  } catch (e) {
    console.error("   [UpSeller] Erro ao ler catálogo:", e.message);
  }
  return null;
}

function salvarCatalogoUpSeller(skus) {
  const dados = { skus, atualizadoEm: Date.now(), total: skus.length };
  fs.writeFileSync(UPSELLER_CATALOGO_FILE, JSON.stringify(dados, null, 2));
  console.log(`   [UpSeller] Catálogo salvo: ${skus.length} SKUs`);
}

// Remove acentos para matching (salmão→salmao, lilás→lilas, onça→onca)
function removerAcentos(str) {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Normaliza uma chave para matching (remove zeros à esquerda, acentos, lowercase, trim)
function normalizarChaveMatch(ref, cor, tam) {
  const refNorm = String(parseInt(ref) || ref).trim();
  const corNorm = removerAcentos(cor.toLowerCase())
    .replace(/\s+/g, ' ')
    .replace(/\bbebe\b/g, 'bb') // normaliza "bebê"/"bebe" → "bb"
    .replace(/\brc\b/g, '')   // remove sufixo RC
    .replace(/\brg\b/g, '')   // remove sufixo RG
    .replace(/\bristre?a?\b/g, '')  // remove "risca/ristra"
    .replace(/\bristado\b/g, '')
    .replace(/\blistrado\b/g, '')
    .replace(/\bestampado\b/g, '')
    .replace(/\bliso\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  const tamNorm = (tam || '').toUpperCase().trim();
  return `${refNorm}|${corNorm}|${tamNorm}`;
}

// Parseia um SKU da UpSeller (ex: "0010-Verde Militar-M") → { ref, cor, tam }
function parseUpSellerSku(sku) {
  if (!sku) return null;
  const parts = sku.split('-');
  if (parts.length < 2) return null;

  const ref = parts[0].trim();
  const tamanhos = ['P', 'M', 'G', 'GG', 'PP', 'EG', 'EGG', 'XG', 'XXG', 'U'];

  // Checa se o último segmento é um tamanho
  const ultimaParte = parts[parts.length - 1].trim().toUpperCase();
  if (tamanhos.includes(ultimaParte) && parts.length >= 3) {
    const cor = parts.slice(1, -1).join('-').trim();
    return { ref, cor, tam: ultimaParte };
  }

  // Sem tamanho (ex: "101-Sortido", "0082-ZEBRA")
  const cor = parts.slice(1).join('-').trim();
  return { ref, cor, tam: '' };
}

// Constrói mapa de matching: chave normalizada → produto Bling
// Prioriza variações (tipo V) sobre produtos-pai (tipo P) para evitar estoque agregado
function construirMapaBling(produtos) {
  const mapa = new Map();
  for (const p of produtos) {
    if (!p.descricao) continue;
    const matchRef = p.descricao.match(/^(\d+)/);
    if (!matchRef) continue;
    const ref = matchRef[1];

    const matchCor = p.descricao.match(/\bCOR[:\s]+([^,;]+)/i);
    if (!matchCor) continue;
    const cor = matchCor[1].trim();

    const matchTam = p.descricao.match(/\bTAM(?:ANHO)?[:\s]+([^,;\s]+)/i);
    const tam = matchTam ? matchTam[1].trim() : '';

    const chave = normalizarChaveMatch(ref, cor, tam);
    const existente = mapa.get(chave);
    if (!existente) {
      mapa.set(chave, p);
    } else if (existente.tipo === 'P' && p.tipo !== 'P') {
      // Sobrescreve produto-pai com variação real (estoque individual, não agregado)
      mapa.set(chave, p);
    } else if (existente.tipo === p.tipo && p.codigo && p.codigo.length > (existente.codigo || '').length) {
      // Mesmo tipo: prefere o produto com código mais específico (SKU de variação)
      mapa.set(chave, p);
    }
  }
  return mapa;
}

// Busca um produto Bling no mapa com fallback de matching flexível
function buscarNoMapaBling(mapaBling, ref, cor, tam) {
  // 1. Tentativa exata
  const chaveExata = normalizarChaveMatch(ref, cor, tam);
  if (mapaBling.has(chaveExata)) return mapaBling.get(chaveExata);

  // 2. Tentativa sem tamanho (para produtos sem variação de tamanho)
  if (tam) {
    const chaveSemTam = normalizarChaveMatch(ref, cor, '');
    if (mapaBling.has(chaveSemTam)) return mapaBling.get(chaveSemTam);
  }

  // 3. Busca parcial - cor do UpSeller contida na cor do Bling ou vice-versa
  const refNorm = String(parseInt(ref) || ref).trim();
  const corNormBusca = removerAcentos(cor.toLowerCase()).replace(/\s+/g, ' ').replace(/\bbebe\b/g, 'bb').trim();
  const tamNorm = (tam || '').toUpperCase().trim();

  for (const [chave, produto] of mapaBling) {
    const partes = chave.split('|');
    if (partes[0] !== refNorm) continue;
    if (tamNorm && partes[2] && partes[2] !== tamNorm) continue;

    const corMapa = partes[1];
    // Checa se uma cor contém a outra (para truncamentos como "pistach" vs "pistache")
    if (corMapa.includes(corNormBusca) || corNormBusca.includes(corMapa)) {
      return produto;
    }
  }

  return null;
}

function lerCatalogoDoDisco() {
  try {
    if (fs.existsSync(CATALOGO_CACHE_FILE)) {
      const raw = JSON.parse(fs.readFileSync(CATALOGO_CACHE_FILE, "utf8"));
      console.log(`   [Cache] Catálogo carregado do disco: ${raw.produtos.length} produtos (salvo em ${new Date(raw.atualizadoEm).toLocaleString('pt-BR')})`);
      return raw;
    }
  } catch (e) {
    console.error("   [Cache] Erro ao ler cache do disco:", e.message);
  }
  return null;
}

function salvarCatalogoNoDisco(produtos) {
  try {
    const dados = { produtos, atualizadoEm: Date.now() };
    fs.writeFileSync(CATALOGO_CACHE_FILE, JSON.stringify(dados));
    console.log(`   [Cache] Catálogo salvo no disco: ${produtos.length} produtos`);
  } catch (e) {
    console.error("   [Cache] Erro ao salvar cache no disco:", e.message);
  }
}

async function sincronizarCatalogoEmSegundoPlano() {
  if (sincronizandoCatalogo) {
    console.log("   [Sync] Sincronização já em andamento, pulando...");
    return;
  }
  sincronizandoCatalogo = true;
  try {
    console.log("   [Sync] Iniciando sincronização do catálogo com Bling...");
    const token = await obterAccessToken();
    const produtos = await buscarEstoque(token);
    cacheProdutos = produtos;
    ultimoCacheHora = Date.now();
    salvarCatalogoNoDisco(produtos);
    console.log(`   [Sync] Sincronização concluída! ${produtos.length} produtos atualizados.`);
  } catch (e) {
    console.error("   [Sync] Falha na sincronização:", e.message);
  } finally {
    sincronizandoCatalogo = false;
  }
}

// Carrega cache do disco ao iniciar o servidor (instantâneo)
const cacheDisco = lerCatalogoDoDisco();
if (cacheDisco && cacheDisco.produtos.length > 0) {
  cacheProdutos = cacheDisco.produtos;
  ultimoCacheHora = cacheDisco.atualizadoEm;
} 

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

// Faz chamada ao Bling com retry automático em caso de rate-limit (429) ou timeout
async function blingRequest(url, accessToken, params = {}) {
  const maxTentativas = 3;
  for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
    try {
      const resp = await axios.get(url, {
        headers: { Authorization: `Bearer ${accessToken}` },
        params,
        timeout: 15000
      });
      return resp;
    } catch (err) {
      const status = err.response?.status;
      console.log(`   [Bling] Erro HTTP ${status || err.code} em ${url.split('/').pop()} (tentativa ${tentativa}/${maxTentativas})`);
      if (tentativa === maxTentativas) throw err;
      const espera = tentativa * 2000;
      console.log(`   [Bling] Aguardando ${espera/1000}s antes de tentar novamente...`);
      await delay(espera);
    }
  }
}

async function buscarEstoque(accessToken) {
  const todosProdutos = [];
  let pagina = 1;
  let temMais = true;

  console.log("   [Bling] Iniciando busca de produtos...");
  do {
    try {
      const resp = await blingRequest("https://www.bling.com.br/Api/v3/produtos", accessToken, { pagina, limite: 100, tipo: 'T' });

      const data = resp.data?.data ?? [];
      todosProdutos.push(...data);
      console.log(`   [Bling] Página ${pagina}: +${data.length} produtos (total: ${todosProdutos.length})`);

      if (data.length < 100) temMais = false;
      else pagina++;
    } catch (err) {
      console.error(`   [Bling] FALHA na página ${pagina} após 3 tentativas: ${err.message}`);
      // Se já pegou alguma coisa, usa o que tem em vez de travar
      if (todosProdutos.length > 0) {
        console.log(`   [Bling] Usando ${todosProdutos.length} produtos já carregados.`);
        temMais = false;
      } else {
        throw err;
      }
    }
    await delay(600);
  } while (temMais);

  console.log(`   [Bling] ${todosProdutos.length} produtos. Buscando saldos...`);

  const ids = todosProdutos.map(p => p.id).filter(id => id);
  const lotes = [];
  for (let i = 0; i < ids.length; i += 50) lotes.push(ids.slice(i, i + 50));

  const saldos = [];
  for (let li = 0; li < lotes.length; li++) {
    const lote = lotes[li];
    try {
      const params = new URLSearchParams();
      for (const id of lote) params.append("idsProdutos[]", id);
      const resp = await blingRequest("https://www.bling.com.br/Api/v3/estoques/saldos", accessToken, params);
      saldos.push(...(resp.data?.data ?? []));
      console.log(`   [Bling] Saldos: lote ${li+1}/${lotes.length} OK`);
    } catch (err) {
      // NÃO trava — pula o lote que falhou e continua
      console.error(`   [Bling] Saldos lote ${li+1}/${lotes.length} FALHOU (pulando): ${err.message}`);
    }
    await delay(600);
  }

  console.log(`   [Bling] Concluído! ${saldos.length} saldos carregados.`);

  // Extrai saldo específico do depósito SEDE
  const mapaSaldos = new Map();
  for (const s of saldos) {
    const idProduto = s.produto?.id || s.id;
    if (!idProduto) continue;

    let saldoSede = 0;
    // Tenta pegar saldo específico do depósito SEDE
    if (s.depositos && Array.isArray(s.depositos)) {
      const depSede = s.depositos.find(d => d.id === DEPOSITO_SEDE_ID);
      if (depSede) {
        saldoSede = depSede.saldoFisico ?? depSede.saldoVirtual ?? 0;
      } else {
        // Depósito SEDE não encontrado na lista — usa saldoFisicoTotal como fallback
        saldoSede = s.saldoFisicoTotal ?? 0;
      }
    } else {
      // Sem breakdown por depósito — usa saldoFisicoTotal
      saldoSede = s.saldoFisicoTotal ?? 0;
    }
    mapaSaldos.set(idProduto, saldoSede);
  }

  return todosProdutos.map(p => ({
    id: p.id,
    codigo: p.codigo,
    gtin: p.gtin || '',
    descricao: p.nome || p.descricao || '',
    nome: p.nome || '',
    tipo: p.tipo || '',
    saldoFisicoTotal: mapaSaldos.get(p.id) ?? 0
  }));
}

const app = express();
app.use(cors());

// 🔐 Serve apenas arquivos públicos (HTML) — bloqueia tokens.json, .git, etc.
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/dashboard.html', (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/checkout.html', (req, res) => res.sendFile(path.join(__dirname, 'checkout.html')));
app.get('/wms.html', (req, res) => res.sendFile(path.join(__dirname, 'wms.html')));

app.use(express.json({ limit: '5mb' }));

// 🏥 Health check — monitoramento do servidor
app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        uptime: Math.floor(process.uptime()),
        cache: cacheProdutos ? cacheProdutos.length : 0,
        memoria: Math.round(process.memoryUsage().heapUsed / 1024 / 1024) + 'MB',
        operacoesPendentes: operacoesFinalizadas.size
    });
});

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
                // Prioriza match exato primeiro
                for (const nome of nomesPossiveis) {
                    const exato = colunas.find(c => normalizar(c) === normalizar(nome));
                    if (exato) return exato;
                }
                // Depois tenta substring (apenas para nomes com 6+ caracteres para evitar falsos positivos)
                for (const nome of nomesPossiveis) {
                    if (normalizar(nome).length < 6) continue;
                    const parcial = colunas.find(c => normalizar(c).includes(normalizar(nome)));
                    if (parcial) return parcial;
                }
                return undefined;
            };

            const colPedido = acharColunaExata(['Nº de Pedido da Plataforma', 'code', 'pedido']) || colunas[0];
            const colSku = acharColunaExata(['SKU (Armazém)', 'Order Items__reference', 'sku']) || 'sku';
            const colNome = acharColunaExata(['Nome do Produto', 'Nome do Anúncio', 'Order Items__name']) || 'name';
            
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

        // 🔄 Agrupa itens duplicados (mesmo pedido + sku) somando quantidades
        const mapaAgrupado = new Map();
        for (const item of novosItens) {
            const chave = `${item.pedido}|||${item.sku}`;
            if (mapaAgrupado.has(chave)) {
                mapaAgrupado.get(chave).qtd += item.qtd;
            } else {
                mapaAgrupado.set(chave, { ...item });
            }
        }
        novosItens = Array.from(mapaAgrupado.values());

        // Remove itens já existentes para o mesmo pedido+sku (evita duplicatas ao re-subir)
        const chavesNovas = new Set(novosItens.map(i => `${i.pedido}|||${i.sku}`));
        bancoDadosPlanilha = bancoDadosPlanilha.filter(i => !chavesNovas.has(`${i.pedido}|||${i.sku}`));
        bancoDadosPlanilha.push(...novosItens);

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

        // Garante que o cache de produtos esteja populado (para busca na troca/adição)
        if (!cacheProdutos || cacheProdutos.length === 0) {
            const doDisco = lerCatalogoDoDisco();
            if (doDisco && doDisco.produtos && doDisco.produtos.length > 0) {
                cacheProdutos = doDisco.produtos;
                ultimoCacheHora = doDisco.atualizadoEm || Date.now();
                console.log(`   [Checkout] Cache de produtos carregado do disco: ${cacheProdutos.length} produtos`);
            } else {
                try {
                    cacheProdutos = await buscarEstoque(token);
                    ultimoCacheHora = Date.now();
                    salvarCatalogoNoDisco(cacheProdutos);
                    console.log(`   [Checkout] Cache de produtos carregado da API: ${cacheProdutos.length} produtos`);
                } catch (e) {
                    console.log(`   [Checkout] Aviso: não foi possível carregar produtos para troca: ${e.message}`);
                }
            }
        }

        res.json({ sucesso: true, total: cachePedidosRecentes.length });
    } catch (e) {
        res.status(500).json({ erro: "Falha ao sincronizar pedidos" });
    }
});

// 🟢 BUSCAR PEDIDO 
app.get('/api/checkout/pedido/:numero', async (req, res) => {
    const numero = req.params.numero.trim();
    console.log(`\n🔍 [Checkout] Solicitada busca pelo pedido: ${numero}`);
    
    // 1. Tenta achar na Planilha (agrupa duplicados por SKU)
    const itensCSV = bancoDadosPlanilha.filter(i => i.pedido === numero);
    if (itensCSV.length > 0) {
        console.log(`✅ [Checkout] Pedido ${numero} encontrado na Planilha!`);

        // Constrói mapa de matching Bling para resolver códigos de barras
        let mapaBling = null;
        if (cacheProdutos && cacheProdutos.length > 0) {
            mapaBling = construirMapaBling(cacheProdutos);
            console.log(`   [Checkout-Planilha] Mapa Bling: ${mapaBling.size} variações para matching de barcode`);
        }

        const mapaItens = new Map();
        for (const i of itensCSV) {
            if (mapaItens.has(i.sku)) {
                mapaItens.get(i.sku).esperado += i.qtd;
            } else {
                // Resolve código Bling + GTIN usando matching inteligente (SKU UpSeller → produto Bling)
                let gtin = '';
                let codigoBling = '';
                let produtoId = null;

                // 1º Tenta match direto pelo código no cache
                if (cacheProdutos) {
                    const cached = cacheProdutos.find(p => String(p.codigo).toLowerCase() === String(i.sku).toLowerCase());
                    if (cached) {
                        gtin = cached.gtin || '';
                        codigoBling = cached.codigo || '';
                        produtoId = cached.id || null;
                    }
                }

                // 2º Se não achou, usa matching inteligente (parseUpSellerSku → buscarNoMapaBling)
                if (!codigoBling && mapaBling) {
                    const parsed = parseUpSellerSku(i.sku);
                    if (parsed) {
                        const produtoBling = buscarNoMapaBling(mapaBling, parsed.ref, parsed.cor, parsed.tam);
                        if (produtoBling) {
                            gtin = produtoBling.gtin || '';
                            codigoBling = produtoBling.codigo || '';
                            produtoId = produtoBling.id || null;
                            console.log(`   🔗 "${i.sku}" → Bling: "${produtoBling.descricao}" (cod: ${codigoBling}, gtin: ${gtin || 'vazio'})`);
                        }
                    }
                }

                mapaItens.set(i.sku, { sku: i.sku, gtin, codigoBling, produtoId, nome: i.nome, esperado: i.qtd, conferido: 0 });
            }
        }
        return res.json({ origem: 'PLANILHA', numero: numero, itens: Array.from(mapaItens.values()) });
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
        
        const itensBling = respDetalhes.data.data.itens.map(i => {
            const sku = i.codigo || i.produto?.codigo || "S/COD";
            const produtoId = i.produto?.id || null;
            // Busca o GTIN no cache para permitir conferência por código de barras
            let gtin = '';
            if (cacheProdutos && produtoId) {
                const cached = cacheProdutos.find(p => p.id === produtoId);
                if (cached) gtin = cached.gtin || '';
            }
            // Se não achou por ID, tenta por código
            if (!gtin && cacheProdutos) {
                const cached = cacheProdutos.find(p => String(p.codigo).toLowerCase() === String(sku).toLowerCase());
                if (cached) gtin = cached.gtin || '';
            }
            return {
                sku,
                gtin,
                nome: i.descricao || "Produto Sem Nome",
                esperado: Math.round(i.quantidade),
                conferido: 0,
                produtoId
            };
        });

        console.log(`✅ [Checkout] Sucesso!`);
        res.json({ origem: 'BLING', id: pedidoId, numero: numeroBling, numeroLoja: numero, itens: itensBling });

    } catch (e) {
        res.status(500).json({ erro: "Erro de comunicação ao buscar pedido no Bling." });
    }
});

// 👇 ROTA DE FINALIZAR PEDIDO 👇
app.post('/api/checkout/finalizar', async (req, res) => {
    const { origem, id, itens, numero, trocas, adicionados, removidos, idempotencyKey } = req.body;

    // 🔒 PROTEÇÃO CONTRA DUPLICIDADE
    limparOperacoesExpiradas();
    if (idempotencyKey && operacoesFinalizadas.has(idempotencyKey)) {
        const operacaoAnterior = operacoesFinalizadas.get(idempotencyKey);
        console.log(`🔒 [Checkout] Operação duplicada bloqueada! Pedido ${numero} (key: ${idempotencyKey}) já foi processado em ${new Date(operacaoAnterior.timestamp).toLocaleTimeString()}`);
        return res.json(operacaoAnterior.resultado);
    }

    try {
        const token = await obterAccessToken();
        const depositoId = DEPOSITO_SEDE_ID;

        // ── Processar trocas de itens (funciona para BLING e PLANILHA) ──
        if (trocas && trocas.length > 0) {
            console.log(`\n🔄 [Checkout] Processando ${trocas.length} troca(s) no pedido ${numero}...`);
            for (const troca of trocas) {
                // Resolver ID do produto original se não veio do frontend (PLANILHA)
                if (!troca.originalProdutoId && troca.originalSku) {
                    try {
                        const respBusca = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(troca.originalSku)}`, { headers: { Authorization: `Bearer ${token}` }});
                        if (respBusca.data?.data?.length > 0) {
                            troca.originalProdutoId = respBusca.data.data[0].id;
                            console.log(`   🔍 Resolvido ID do item original "${troca.originalSku}" → ${troca.originalProdutoId}`);
                        }
                    } catch (e) {}
                }
                // ENTRADA do item removido (devolver ao estoque)
                if (troca.originalProdutoId) {
                    try {
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: troca.originalProdutoId },
                            deposito: { id: depositoId },
                            operacao: "E",
                            quantidade: troca.quantidade || 1,
                            observacoes: `Troca no Checkout - ENTRADA (item devolvido). Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
                        console.log(`   ✅ ENTRADA: ${troca.quantidade || 1}x "${troca.originalNome}" devolvido ao estoque`);
                        atualizarCacheEstoque(troca.originalProdutoId, troca.quantidade || 1, 'E');
                    } catch (errEntrada) {
                        console.error(`   ⚠️ Erro na ENTRADA do item trocado "${troca.originalNome}":`, errEntrada.response?.data || errEntrada.message);
                    }
                }
                // SAÍDA do novo item (retirar do estoque)
                if (troca.novoProdutoId) {
                    try {
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: troca.novoProdutoId },
                            deposito: { id: depositoId },
                            operacao: "S",
                            quantidade: troca.quantidade || 1,
                            observacoes: `Troca no Checkout - SAÍDA (item substituto). Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
                        console.log(`   ✅ SAÍDA: ${troca.quantidade || 1}x "${troca.novoNome}" retirado do estoque`);
                        atualizarCacheEstoque(troca.novoProdutoId, troca.quantidade || 1, 'S');
                    } catch (errSaida) {
                        console.error(`   ⚠️ Erro na SAÍDA do novo item "${troca.novoNome}":`, errSaida.response?.data || errSaida.message);
                    }
                }
            }
        }

        // ── Processar itens removidos (ENTRADA no estoque para compensar) ──
        if (removidos && removidos.length > 0) {
            console.log(`\n➖ [Checkout] Processando ${removidos.length} item(ns) removido(s) no pedido ${numero}...`);
            for (const item of removidos) {
                let prodId = item.produtoId;
                if (!prodId && item.sku) {
                    try {
                        const resp = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(item.sku)}`, { headers: { Authorization: `Bearer ${token}` }});
                        if (resp.data?.data?.length > 0) prodId = resp.data.data[0].id;
                    } catch (e) {}
                }
                if (prodId) {
                    try {
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: prodId },
                            deposito: { id: depositoId },
                            operacao: "E",
                            quantidade: item.quantidade || 1,
                            observacoes: `Item removido no Checkout - ENTRADA (devolvido ao estoque). Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
                        console.log(`   ✅ ENTRADA: ${item.quantidade || 1}x "${item.nome}" devolvido ao estoque (item removido)`);
                        atualizarCacheEstoque(prodId, item.quantidade || 1, 'E');
                    } catch (err) {
                        console.error(`   ⚠️ Erro na ENTRADA do item removido "${item.nome}":`, err.response?.data || err.message);
                    }
                } else {
                    console.error(`   ❌ Não foi possível resolver ID do item removido "${item.sku}" para entrada no estoque`);
                }
            }
        }

        // ── Processar itens adicionados (SAÍDA do estoque) ──
        if (adicionados && adicionados.length > 0) {
            console.log(`\n➕ [Checkout] Processando ${adicionados.length} item(ns) adicionado(s) no pedido ${numero}...`);
            for (const item of adicionados) {
                if (item.produtoId) {
                    try {
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: item.produtoId },
                            deposito: { id: depositoId },
                            operacao: "S",
                            quantidade: item.quantidade || 1,
                            observacoes: `Item adicionado no Checkout - SAÍDA. Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
                        console.log(`   ✅ SAÍDA: ${item.quantidade || 1}x "${item.nome}" retirado do estoque (item adicionado)`);
                        atualizarCacheEstoque(item.produtoId, item.quantidade || 1, 'S');
                    } catch (err) {
                        console.error(`   ⚠️ Erro na SAÍDA do item adicionado "${item.nome}":`, err.response?.data || err.message);
                    }
                }
            }
        }

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

            try {
                await axios.patch(`https://www.bling.com.br/Api/v3/pedidos/vendas/${id}/situacoes/9`, {}, {
                    headers: { Authorization: `Bearer ${token}` }
                });
                console.log(`✅ [Checkout] Pedido Bling ${numero} marcado como Atendido!`);
            } catch (errSituacao) {
                console.error(`⚠️ Erro ao mudar situação do pedido ${numero} para Atendido:`, errSituacao.response?.data || errSituacao.message);
                console.log(`   (As movimentações de estoque foram realizadas com sucesso)`);
            }

        } else {
            let baixasOk = 0;
            let baixasFalha = 0;

            // Constrói mapa de matching Bling usando o cache (mesmo sistema da exportação UpSeller)
            let mapaBling = null;
            if (cacheProdutos && cacheProdutos.length > 0) {
                mapaBling = construirMapaBling(cacheProdutos);
                console.log(`   [Checkout] Mapa Bling construído: ${mapaBling.size} variações para matching`);
            }

            for (const item of itens) {
                try {
                    let prodId = null;

                    // 1️⃣ Tenta buscar direto pelo código na API do Bling
                    try {
                        const respProd = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(item.sku)}`, { headers: { Authorization: `Bearer ${token}` }});
                        if (respProd.data?.data?.length > 0) {
                            prodId = respProd.data.data[0].id;
                            console.log(`   ✅ SKU "${item.sku}" encontrado direto na API Bling (ID: ${prodId})`);
                        }
                    } catch (e) {}

                    // 2️⃣ Se não achou, usa o sistema de matching robusto (mesmo da exportação UpSeller)
                    if (!prodId && mapaBling) {
                        const parsed = parseUpSellerSku(item.sku);
                        if (parsed) {
                            const produtoBling = buscarNoMapaBling(mapaBling, parsed.ref, parsed.cor, parsed.tam);
                            if (produtoBling) {
                                prodId = produtoBling.id;
                                console.log(`   ✅ SKU "${item.sku}" encontrado via matching inteligente (ID: ${prodId}, Bling: "${produtoBling.descricao}")`);
                            } else {
                                console.log(`   ⚠️ SKU "${item.sku}" parseado como ref=${parsed.ref} cor=${parsed.cor} tam=${parsed.tam} — sem match no mapa Bling`);
                            }
                        } else {
                            console.log(`   ⚠️ SKU "${item.sku}" não é formato UpSeller válido (não foi possível parsear)`);
                        }
                    }

                    // 3️⃣ Último recurso: busca pela referência numérica na API Bling
                    if (!prodId) {
                        const parsed = parseUpSellerSku(item.sku);
                        if (parsed && parsed.ref) {
                            try {
                                const respRef = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(parsed.ref)}`, { headers: { Authorization: `Bearer ${token}` }});
                                const produtos = respRef.data?.data || [];
                                for (const p of produtos) {
                                    // Extrai cor e tamanho do nome Bling e compara com normalização
                                    const matchCor = p.nome?.match(/\bCOR[:\s]+([^,;]+)/i);
                                    const matchTam = p.nome?.match(/\bTAM(?:ANHO)?[:\s]+([^,;\s]+)/i);
                                    if (matchCor) {
                                        const corBling = matchCor[1].trim();
                                        const tamBling = matchTam ? matchTam[1].trim() : '';
                                        const chaveBling = normalizarChaveMatch(parsed.ref, corBling, tamBling);
                                        const chaveUpSeller = normalizarChaveMatch(parsed.ref, parsed.cor, parsed.tam);
                                        if (chaveBling === chaveUpSeller) {
                                            prodId = p.id;
                                            console.log(`   ✅ SKU "${item.sku}" encontrado via busca por ref "${parsed.ref}" (ID: ${prodId})`);
                                            break;
                                        }
                                    }
                                }
                            } catch (e) {}
                        }
                    }

                    if (prodId) {
                        await axios.post("https://www.bling.com.br/Api/v3/estoques", {
                            produto: { id: prodId },
                            deposito: { id: depositoId },
                            operacao: "S",
                            quantidade: item.esperado,
                            observacoes: `Baixa via Checkout de Expedição. Pedido: ${numero}`
                        }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
                        baixasOk++;
                        atualizarCacheEstoque(prodId, item.esperado, 'S');
                        console.log(`   📦 Saída registrada: ${item.esperado}x "${item.sku}" no depósito SEDE (Pedido ${numero})`);
                    } else {
                        baixasFalha++;
                        console.error(`   ❌ SKU "${item.sku}" NÃO encontrado na Bling! Saída NÃO registrada.`);
                    }
                } catch (errItem) {
                    baixasFalha++;
                    console.error(`   ⚠️ Erro ao dar baixa no SKU ${item.sku}:`, errItem.response?.data || errItem.message);
                }
            }
            console.log(`✅ [Checkout] Pedido Planilha/UpSeller ${numero} finalizado. Baixas OK: ${baixasOk} | Falhas: ${baixasFalha}`);
            if (baixasFalha > 0) {
                const resultado = { sucesso: true, aviso: `${baixasFalha} SKU(s) não encontrado(s) na Bling. Verifique o log do servidor.` };
                if (idempotencyKey) operacoesFinalizadas.set(idempotencyKey, { timestamp: Date.now(), resultado });
                return res.json(resultado);
            }
        }

        // 🔒 Registra operação como concluída
        const resultado = { sucesso: true };
        if (idempotencyKey) operacoesFinalizadas.set(idempotencyKey, { timestamp: Date.now(), resultado });
        console.log(`🔒 [Checkout] Operação registrada com sucesso (key: ${idempotencyKey || 'sem-key'})`);
        res.json(resultado);
    } catch (e) {
        console.error(e);
        res.status(500).json({ erro: "Erro ao finalizar pedido no servidor." });
    }
});


// ──────────────────────────────────────────────
// 🔧 PARSER: Converte nome do Bling → SKU da UpSeller
// ──────────────────────────────────────────────
// Bling:    "111- CONJUNTO JORDANIA GG COR VERDE MENTA, TAMANHO GG"
// UpSeller: "111-Verde Menta-GG"
// Formato:  REF-Cor-Tamanho (Title Case, abreviações)
function blingParaSkuUpSeller(nomeBling) {
    if (!nomeBling) return null;

    // 1. Extrai o número de referência no início
    const matchRef = nomeBling.match(/^(\d+)/);
    if (!matchRef) return null;
    const ref = matchRef[1];

    // 2. Extrai a COR (aceita "COR:", "COR " — para antes de , ou ;)
    const matchCor = nomeBling.match(/\bCOR[:\s]+([^,;]+)/i);
    if (!matchCor) return null; // Sem cor = produto-pai, pula
    let cor = matchCor[1].trim();

    // 3. Extrai o TAMANHO (aceita "TAMANHO:", "TAMANHO ")
    const matchTam = nomeBling.match(/\bTAM(?:ANHO)?[:\s]+([^,;\s]+)/i);
    let tam = matchTam ? matchTam[1].trim().toUpperCase() : null;

    // 4. Converte cor para Title Case (tratando acentos corretamente)
    cor = cor.toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());

    // 5. Aplica abreviações e correções conhecidas
    cor = cor.replace(/\bBebe\b/gi, 'BB');
    cor = cor.replace(/\bBb\b/g, 'BB');

    // 6. Monta o SKU: REF-Cor ou REF-Cor-TAM
    let sku = `${ref}-${cor}`;
    if (tam) sku += `-${tam}`;

    return sku;
}

// ──────────────────────────────────────────────
// 🔍 DIAGNÓSTICO DE SKUs (BLING → UPSELLER)
// ──────────────────────────────────────────────
// 🔒 Debug endpoints — somente em ambiente local
app.get('/api/debug-skus', async (req, res) => {
    if (process.env.NODE_ENV === 'production') return res.status(404).json({ erro: 'Não encontrado' });
    try {
        const token = await obterAccessToken();
        console.log(`\n🔍 [Debug] Buscando produtos para diagnóstico de SKU...`);

        const produtos = await buscarEstoque(token);

        // Classifica e ordena: variações primeiro, depois pais
        const classificados = produtos.map(p => {
            const skuGerado = blingParaSkuUpSeller(p.descricao);
            return { ...p, skuGerado, ehVariacao: !!skuGerado };
        });
        classificados.sort((a, b) => {
            if (a.ehVariacao && !b.ehVariacao) return -1;
            if (!a.ehVariacao && b.ehVariacao) return 1;
            return (a.descricao || '').localeCompare(b.descricao || '');
        });

        let contVariacoes = classificados.filter(p => p.ehVariacao).length;
        let contPais = classificados.filter(p => !p.ehVariacao).length;

        let html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <title>Raio-X de SKUs (Bling)</title>
        <style>
            * { box-sizing: border-box; margin: 0; padding: 0; }
            body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; color: #333; padding: 20px; }
            h1 { color: #333; font-size: 1.4em; margin-bottom: 4px; }
            .subtitle { color: #666; font-size: 0.9em; margin-bottom: 20px; }
            .resumo { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
            .card { padding: 15px 20px; border-radius: 8px; color: #fff; min-width: 180px; }
            .card-total { background: #3b82f6; }
            .card-ok { background: #22c55e; }
            .card-pai { background: #ef4444; }
            .card h2 { font-size: 2em; margin: 0; }
            .card span { font-size: 0.85em; opacity: 0.9; }
            table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            th { background: #1e293b; color: #fff; padding: 12px 15px; text-align: left; font-size: 0.85em; text-transform: uppercase; }
            td { padding: 10px 15px; border-bottom: 1px solid #e5e7eb; font-size: 0.9em; }
            tr:hover { background: #f0f9ff; }
            tr.variacao { background: #f0fdf4; }
            tr.pai { background: #fef2f2; }
            .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }
            .badge-ok { background: #dcfce7; color: #166534; }
            .badge-pai { background: #fee2e2; color: #991b1b; }
            .sku { font-family: monospace; font-weight: 700; color: #059669; font-size: 0.95em; }
            .filtros { margin-bottom: 15px; display: flex; gap: 8px; align-items: center; }
            .filtros button { padding: 6px 14px; border: 1px solid #d1d5db; border-radius: 6px; background: #fff; cursor: pointer; font-size: 0.85em; }
            .filtros button.ativo { background: #1e293b; color: #fff; border-color: #1e293b; }
            .filtros input { padding: 6px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 0.85em; width: 250px; }
        </style></head><body>
        <h1>Raio-X de SKUs (Bling)</h1>
        <p class="subtitle">Veja a diferenca entre o Produto Pai (sem COR no nome) e a Variacao/Filho (com COR e TAMANHO)</p>

        <div class="resumo">
            <div class="card card-total"><h2>${produtos.length}</h2><span>Total de Produtos</span></div>
            <div class="card card-ok"><h2>${contVariacoes}</h2><span>Variacoes (com SKU)</span></div>
            <div class="card card-pai"><h2>${contPais}</h2><span>Pais (ignorados)</span></div>
        </div>

        <div class="filtros">
            <button class="ativo" onclick="filtrar('todos')">Todos</button>
            <button onclick="filtrar('variacao')">Variacoes</button>
            <button onclick="filtrar('pai')">Pais</button>
            <input type="text" id="busca" placeholder="Buscar por nome ou SKU..." oninput="filtrar()">
        </div>

        <table>
        <thead><tr><th>Nome no Bling</th><th>Codigo (Bling)</th><th>SKU UpSeller</th><th>Estoque</th><th>Tipo</th></tr></thead>
        <tbody id="corpo">`;

        classificados.forEach(p => {
            if (p.ehVariacao) {
                html += `<tr class="variacao" data-tipo="variacao">
                    <td>${p.descricao || '-'}</td>
                    <td>${p.codigo || '-'}</td>
                    <td class="sku">${p.skuGerado}</td>
                    <td>${p.saldoFisicoTotal}</td>
                    <td><span class="badge badge-ok">Variacao (OK)</span></td>
                </tr>`;
            } else {
                html += `<tr class="pai" data-tipo="pai">
                    <td>${p.descricao || '-'}</td>
                    <td>${p.codigo || '-'}</td>
                    <td style="color:#999;">-</td>
                    <td>${p.saldoFisicoTotal}</td>
                    <td><span class="badge badge-pai">Pai (Ignorado)</span></td>
                </tr>`;
            }
        });

        html += `</tbody></table>

        <script>
        let filtroAtual = 'todos';
        function filtrar(tipo) {
            if (tipo) filtroAtual = tipo;
            const busca = document.getElementById('busca').value.toLowerCase();
            const linhas = document.querySelectorAll('#corpo tr');
            document.querySelectorAll('.filtros button').forEach(b => b.classList.remove('ativo'));
            document.querySelector('.filtros button[onclick*="' + filtroAtual + '"]').classList.add('ativo');
            linhas.forEach(tr => {
                const tipoLinha = tr.dataset.tipo;
                const texto = tr.textContent.toLowerCase();
                const matchTipo = filtroAtual === 'todos' || tipoLinha === filtroAtual;
                const matchBusca = !busca || texto.includes(busca);
                tr.style.display = (matchTipo && matchBusca) ? '' : 'none';
            });
        }
        </script>
        </body></html>`;

        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send(html);

    } catch (e) {
        console.error("❌ [Debug]", e.message);
        res.status(500).json({ erro: e.message });
    }
});


// ──────────────────────────────────────────────
// 🟢 EXPORTAÇÃO UPSELLER (ESTOQUE ESPELHO)
// Parseia nome do Bling → SKU UpSeller (REF-Cor-Tamanho)
// ──────────────────────────────────────────────
app.get('/api/exportar-upseller', async (req, res) => {
    req.setTimeout(120000);
    res.setTimeout(120000);
    try {
        console.log(`\n📦 [UpSeller] Gerando ZIP (regra: >10un = 2000+real, ≤10un = estoque real)...`);

        // Sempre busca dados frescos do Bling para garantir estoque atualizado
        console.log(`   [UpSeller] Buscando estoque atualizado do Bling...`);
        const token = await obterAccessToken();
        const produtos = await buscarEstoque(token);
        cacheProdutos = produtos;
        ultimoCacheHora = Date.now();
        console.log(`   [UpSeller] ${produtos.length} produtos do Bling (dados frescos).`);

        if (!produtos || produtos.length === 0) {
            return res.status(404).json({ erro: "Nenhum produto encontrado no Bling." });
        }

        let logConteudo = `====================================================\n`;
        logConteudo += `📊 RELATÓRIO DE EXPORTAÇÃO UPSELLER (ESTOQUE ESPELHO)\n`;
        logConteudo += `Data: ${new Date().toLocaleString('pt-BR')}\n`;
        logConteudo += `Regra: Estoque real > 10 → envia 2000 + real | Estoque ≤ 10 → envia estoque real\n`;
        logConteudo += `====================================================\n\n`;

        let qtdAtivo = 0;
        let qtdZerado = 0;
        let qtdIgnorados = 0;
        let qtdSemMatch = 0;
        let qtdKit = 0;

        // Cabeçalho exato da UpSeller (AOA = matriz)
        const dadosPlanilha = [
            [
                "SKU*",
                "Estoque Baixo\n(Não será atualizado se não for preenchido)",
                "Qtd. Total Atualizado\n(Não será atualizado se não for preenchido)",
                "Custo Médio Atualizado\n(Não será atualizado se não for preenchido)"
            ]
        ];

        // Verifica se existe catálogo UpSeller salvo
        const catalogoUpSeller = lerCatalogoUpSeller();
        const usarCatalogoReal = catalogoUpSeller && catalogoUpSeller.skus && catalogoUpSeller.skus.length > 0;

        if (usarCatalogoReal) {
            // ====== MODO CATÁLOGO REAL ======
            console.log(`   [UpSeller] Modo CATÁLOGO REAL: ${catalogoUpSeller.skus.length} SKUs do armazém`);
            logConteudo += `MODO: Catálogo Real UpSeller (${catalogoUpSeller.skus.length} SKUs)\n\n`;

            const mapaBling = construirMapaBling(produtos);
            console.log(`   [UpSeller] Mapa Bling construído: ${mapaBling.size} variações mapeadas`);

            for (const skuReal of catalogoUpSeller.skus) {
                const parsed = parseUpSellerSku(skuReal);
                if (!parsed) {
                    qtdIgnorados++;
                    logConteudo += `[IGNORADO] SKU não parseável: ${skuReal}\n`;
                    continue;
                }

                const produtoBling = buscarNoMapaBling(mapaBling, parsed.ref, parsed.cor, parsed.tam);

                if (!produtoBling) {
                    // Verifica se é um Kit — envia 100 ao invés de 0
                    if (/kit/i.test(skuReal)) {
                        qtdKit++;
                        logConteudo += `[KIT] ${skuReal} — sem match no Bling, enviando 100 (kit)\n`;
                        dadosPlanilha.push([skuReal, "", 100, ""]);
                    } else {
                        qtdSemMatch++;
                        const chaveDebug = normalizarChaveMatch(parsed.ref, parsed.cor, parsed.tam);
                        logConteudo += `[SEM MATCH] ${skuReal} → chave: ${chaveDebug} (enviando 0 para zerar na UpSeller)\n`;
                        dadosPlanilha.push([skuReal, "", 0, ""]);
                    }
                    continue;
                }

                const quantidadeReal = parseInt(produtoBling.saldoFisicoTotal) || 0;
                let quantidadeUpSeller = 0;

                // Kit com match no Bling — sempre envia 100
                if (/kit/i.test(skuReal)) {
                    quantidadeUpSeller = 100;
                    qtdKit++;
                    logConteudo += `[KIT] ${skuReal} — real: ${quantidadeReal} → 100 (kit)\n`;
                } else if (quantidadeReal > 10) {
                    quantidadeUpSeller = 2000 + quantidadeReal;
                    qtdAtivo++;
                    logConteudo += `[ATIVO] ${skuReal} — real: ${quantidadeReal} → ${quantidadeUpSeller}\n`;
                } else {
                    quantidadeUpSeller = quantidadeReal;
                    qtdZerado++;
                    logConteudo += `[BAIXO] ${skuReal} — real: ${quantidadeReal} → ${quantidadeUpSeller} (≤10, mantém real)\n`;
                }

                dadosPlanilha.push([skuReal, "", quantidadeUpSeller, ""]);
            }
        } else {
            // ====== MODO LEGADO (geração de SKU a partir do Bling) ======
            console.log(`   [UpSeller] Modo LEGADO: gerando SKUs a partir dos nomes Bling`);
            logConteudo += `MODO: Geração automática de SKU (sem catálogo UpSeller)\n\n`;

            for (const p of produtos) {
                const skuUpSeller = blingParaSkuUpSeller(p.descricao);

                if (!skuUpSeller) {
                    qtdIgnorados++;
                    continue;
                }

                const skuLimpo = skuUpSeller
                    .replace(/[\u200B\u200C\u200D\uFEFF\u00A0]/g, '')
                    .trim();

                const quantidadeReal = parseInt(p.saldoFisicoTotal) || 0;
                let quantidadeUpSeller = 0;

                // Kit — sempre envia 100
                if (/kit/i.test(skuLimpo)) {
                    quantidadeUpSeller = 100;
                    qtdKit++;
                } else if (quantidadeReal > 10) {
                    quantidadeUpSeller = 2000 + quantidadeReal;
                    qtdAtivo++;
                } else {
                    quantidadeUpSeller = quantidadeReal;
                    qtdZerado++;
                }

                dadosPlanilha.push([skuLimpo, "", quantidadeUpSeller, ""]);
            }
        }

        const totalExportados = qtdAtivo + qtdZerado;

        logConteudo += `\n====================================================\n`;
        logConteudo += `RESUMO:\n`;
        logConteudo += `- SKUs ATIVOS (estoque >10, enviado 2000+real): ${qtdAtivo}\n`;
        logConteudo += `- SKUs BAIXOS (estoque ≤10, enviado estoque real): ${qtdZerado}\n`;
        if (qtdKit > 0) logConteudo += `- SKUs KIT (enviado 100): ${qtdKit}\n`;
        if (qtdSemMatch > 0) logConteudo += `- SKUs SEM MATCH no Bling (zerados na UpSeller): ${qtdSemMatch}\n`;
        logConteudo += `- Ignorados: ${qtdIgnorados}\n`;
        logConteudo += `- Total de linhas na planilha: ${totalExportados + qtdSemMatch + qtdKit}\n`;
        logConteudo += `====================================================\n`;

        console.log(`   [UpSeller] ${totalExportados} SKUs convertidos. ${qtdIgnorados} produtos-pai ignorados.`);

        // Gera planilha via AOA (garante cabeçalho exato com \n)
        const worksheet = xlsx.utils.aoa_to_sheet(dadosPlanilha);
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

        console.log(`✅ [UpSeller] ZIP gerado com ${totalExportados} SKUs!`);

    } catch (e) {
        console.error("❌ Erro ao exportar ZIP UpSeller:", e.message);
        res.status(500).json({ erro: "Erro interno ao gerar o pacote ZIP." });
    }
});


// Rota de teste — verifica se o token do Bling funciona
app.get('/api/teste-bling', async (req, res) => {
  if (process.env.NODE_ENV === 'production') return res.status(404).json({ erro: 'Não encontrado' });
  try {
    console.log("   [Teste] Verificando conexão com o Bling...");
    const token = await obterAccessToken();
    const resp = await axios.get("https://www.bling.com.br/Api/v3/produtos", {
      headers: { Authorization: `Bearer ${token}` },
      params: { pagina: 1, limite: 1 },
      timeout: 10000
    });
    const qtd = resp.data?.data?.length ?? 0;
    console.log(`   [Teste] OK! Bling respondeu com ${qtd} produto(s).`);
    res.json({ ok: true, produtos: qtd, token: token.substring(0, 10) + "..." });
  } catch (err) {
    console.error(`   [Teste] FALHA: ${err.message}`);
    res.status(500).json({ ok: false, erro: err.message, status: err.response?.status });
  }
});

// ──────────────────────────────────────────────
// RESTANTE DO CÓDIGO (DASHBOARD / WMS)
// ──────────────────────────────────────────────

app.get('/api/produtos', async (req, res) => {
  // Timeout de 2 minutos para não travar o navegador
  req.setTimeout(120000);
  res.setTimeout(120000);
  try {
    const trintaMinutos = 1800000;
    if (cacheProdutos && (Date.now() - ultimoCacheHora < trintaMinutos)) {
        console.log(`   [API] /api/produtos — cache OK (${cacheProdutos.length} produtos)`);
        return res.json(cacheProdutos);
    }

    console.log("   [API] /api/produtos — buscando do Bling (sem cache)...");
    const token = await obterAccessToken();
    const produtos = await buscarEstoque(token);

    cacheProdutos = produtos;
    ultimoCacheHora = Date.now();
    console.log(`   [API] /api/produtos — ${produtos.length} produtos carregados e cacheados.`);
    salvarCatalogoNoDisco(produtos);

    res.json(produtos);
  } catch (error) {
    console.error("   [API] /api/produtos ERRO:", error.message);
    if (cacheProdutos) return res.json(cacheProdutos);
    res.status(500).json({ error: "Erro ao buscar produtos do Bling: " + error.message });
  }
});

// 🔍 DEBUG — Estado do cache de produtos
app.get('/api/checkout/cache-status', (req, res) => {
    if (process.env.NODE_ENV === 'production') return res.status(404).json({ erro: 'Não encontrado' });
    // Conta tipos de produtos no cache
    const tipos = {};
    let amostra = [];
    if (cacheProdutos && cacheProdutos.length > 0) {
        for (const p of cacheProdutos) {
            const t = p.tipo || 'sem_tipo';
            tipos[t] = (tipos[t] || 0) + 1;
        }
        amostra = cacheProdutos.slice(0, 5).map(p => ({
            id: p.id, codigo: p.codigo, descricao: p.descricao, tipo: p.tipo, gtin: p.gtin
        }));
    }
    res.json({
        cacheProdutos: cacheProdutos ? cacheProdutos.length : 0,
        ultimoCacheHora: ultimoCacheHora ? new Date(ultimoCacheHora).toLocaleString('pt-BR') : 'nunca',
        sincronizandoCatalogo,
        catalogoCacheExiste: fs.existsSync(CATALOGO_CACHE_FILE),
        tipos,
        amostra
    });
});

// 🔄 PRE-CARREGA PRODUTOS (chamado pelo checkout ao abrir)
app.get('/api/checkout/precarregar-produtos', async (req, res) => {
    if (cacheProdutos && cacheProdutos.length > 0) {
        return res.json({ status: 'ok', total: cacheProdutos.length, fonte: 'memoria' });
    }

    // Tenta disco
    const doDisco = lerCatalogoDoDisco();
    if (doDisco && doDisco.produtos && doDisco.produtos.length > 0) {
        cacheProdutos = doDisco.produtos;
        ultimoCacheHora = doDisco.atualizadoEm || Date.now();
        return res.json({ status: 'ok', total: cacheProdutos.length, fonte: 'disco' });
    }

    // Tenta API
    try {
        const token = await obterAccessToken();
        cacheProdutos = await buscarEstoque(token);
        ultimoCacheHora = Date.now();
        salvarCatalogoNoDisco(cacheProdutos);
        return res.json({ status: 'ok', total: cacheProdutos.length, fonte: 'api' });
    } catch (e) {
        console.error('   [Precarregar] Falha:', e.message);
        return res.json({ status: 'erro', mensagem: e.message, total: 0 });
    }
});

// 🔍 DEBUG — Testa busca de produtos manualmente
app.get('/api/checkout/debug-busca', async (req, res) => {
    if (process.env.NODE_ENV === 'production') return res.status(404).json({ erro: 'Não encontrado' });
    const termo = (req.query.q || '').toLowerCase().trim();
    try {
        const cacheLen = cacheProdutos ? cacheProdutos.length : 0;
        const discoExiste = fs.existsSync(CATALOGO_CACHE_FILE);

        // Amostra do cache
        let amostra = [];
        if (cacheProdutos && cacheProdutos.length > 0) {
            amostra = cacheProdutos.slice(0, 5).map(p => ({
                id: p.id, codigo: p.codigo, descricao: p.descricao, nome: p.nome, tipo: p.tipo, gtin: p.gtin
            }));
        }

        // Tipos no cache
        const tipos = {};
        if (cacheProdutos) {
            for (const p of cacheProdutos) {
                tipos[p.tipo || 'undefined'] = (tipos[p.tipo || 'undefined'] || 0) + 1;
            }
        }

        // Resultados da busca
        let resultados = [];
        if (termo && cacheProdutos) {
            resultados = cacheProdutos.filter(p => {
                const codigo = (p.codigo || '').toLowerCase();
                const descricao = (p.descricao || p.nome || '').toLowerCase();
                const gtin = (p.gtin || '').toLowerCase();
                return codigo.includes(termo) || descricao.includes(termo) || gtin.includes(termo);
            }).slice(0, 5).map(p => ({ codigo: p.codigo, descricao: p.descricao, tipo: p.tipo, gtin: p.gtin }));
        }

        res.json({ cacheLen, discoExiste, sincronizandoCatalogo, tipos, amostra, termoBuscado: termo, resultados });
    } catch (e) {
        res.status(500).json({ erro: e.message });
    }
});

// 🔍 BUSCA DE PRODUTOS PARA TROCA NO CHECKOUT
app.get('/api/checkout/buscar-produtos', async (req, res) => {
    const termo = (req.query.q || '').toLowerCase().trim();
    if (termo.length < 2) return res.json([]);

    try {
        // Usa cache se disponível, senão tenta disco, senão busca da API
        if (!cacheProdutos || cacheProdutos.length === 0) {
            const doDisco = lerCatalogoDoDisco();
            if (doDisco && doDisco.produtos && doDisco.produtos.length > 0) {
                cacheProdutos = doDisco.produtos;
                ultimoCacheHora = doDisco.atualizadoEm || Date.now();
                console.log(`   [Busca Troca] Cache carregado do disco: ${cacheProdutos.length} produtos`);
            } else {
                const token = await obterAccessToken();
                cacheProdutos = await buscarEstoque(token);
                ultimoCacheHora = Date.now();
            }
        }

        console.log(`   [Busca Troca] Termo: "${termo}" | Cache: ${cacheProdutos ? cacheProdutos.length : 'NULL'} produtos`);

        if (!cacheProdutos || cacheProdutos.length === 0) {
            console.log(`   [Busca Troca] ⚠️ CACHE VAZIO! Retornando erro para o frontend.`);
            return res.json({ erro: 'cache_vazio', mensagem: 'Catálogo não carregado. Aguarde a sincronização ou acesse o Dashboard primeiro.' });
        }

        // Log de diagnóstico: amostra dos primeiros 3 produtos para verificar estrutura
        if (cacheProdutos.length > 0) {
            const amostra = cacheProdutos.slice(0, 3).map(p => ({
                codigo: p.codigo, descricao: p.descricao, tipo: p.tipo, nome: p.nome, gtin: p.gtin
            }));
            console.log(`   [Busca Troca] Amostra do cache:`, JSON.stringify(amostra));
        }

        // Verifica se existem variações (tipo V) no cache — se não existem, todos são produtos simples
        const temVariacoes = cacheProdutos.some(p => p.tipo === 'V');

        const filtrados = cacheProdutos.filter(p => {
            // Só ignora produtos pai (tipo P) se o catálogo TEM variações (tipo V)
            // Se todos são tipo P, são produtos simples e devem ser buscáveis
            if (temVariacoes && p.tipo === 'P') {
                const desc2 = (p.descricao || p.nome || '').toLowerCase();
                if (!desc2.includes('cor')) return false;
            }
            const desc = (p.descricao || p.nome || '').toLowerCase();
            const codigo = (p.codigo || '').toLowerCase();
            const gtin = (p.gtin || '').toLowerCase();
            return codigo.includes(termo) || desc.includes(termo) || gtin.includes(termo);
        });

        // Ordena: prioriza matches na descrição que começam com o termo (ex: "107-")
        // Depois agrupa por nome do modelo para manter variações juntas
        filtrados.sort((a, b) => {
            const descA = (a.descricao || a.nome || '').toLowerCase();
            const descB = (b.descricao || b.nome || '').toLowerCase();
            // Prioridade 1: descrição começa com o termo (ex: "107- TRI MABEL")
            const aDescStarts = descA.startsWith(termo + '-') || descA.startsWith(termo + ' ') ? 0 : 1;
            const bDescStarts = descB.startsWith(termo + '-') || descB.startsWith(termo + ' ') ? 0 : 1;
            if (aDescStarts !== bDescStarts) return aDescStarts - bDescStarts;
            // Prioridade 2: código começa com o termo
            const codA = (a.codigo || '').toLowerCase();
            const codB = (b.codigo || '').toLowerCase();
            const aCodStarts = codA.startsWith(termo) ? 0 : 1;
            const bCodStarts = codB.startsWith(termo) ? 0 : 1;
            if (aCodStarts !== bCodStarts) return aCodStarts - bCodStarts;
            // Prioridade 3: ordem alfabética por descrição
            return descA.localeCompare(descB);
        });

        const resultados = filtrados.slice(0, 50); // máximo 50 resultados

        console.log(`   [Busca Troca] Encontrados: ${resultados.length} resultado(s)`);
        if (resultados.length > 0) {
            console.log(`   [Busca Troca] Primeiro resultado:`, JSON.stringify({ codigo: resultados[0].codigo, descricao: resultados[0].descricao, tipo: resultados[0].tipo }));
        }

        res.json(resultados.map(p => ({
            id: p.id,
            codigo: p.codigo,
            descricao: p.descricao || p.nome || 'Sem descrição',
            gtin: p.gtin,
            saldoFisicoTotal: p.saldoFisicoTotal
        })));
    } catch (e) {
        console.error('Erro na busca de produtos para troca:', e.message, e.stack);
        res.status(500).json({ erro: e.message, produtos: [] });
    }
});

// 📲 ALERTA DE ESTOQUE — Gera mensagem formatada para WhatsApp
app.get('/api/alerta-estoque', async (req, res) => {
  try {
    if (!cacheProdutos || cacheProdutos.length === 0) {
      return res.status(400).json({ erro: "Catálogo vazio. Aguarde a sincronização." });
    }

    // Filtra apenas SKUs que existem no catálogo UpSeller (marketplace ativo)
    const catalogoUpseller = lerCatalogoUpSeller();
    const skusUpseller = catalogoUpseller?.skus ? new Set(catalogoUpseller.skus.map(s => s.toLowerCase())) : null;

    let produtosAlerta = cacheProdutos.filter(p => {
      if (p.saldoFisicoTotal > CONFIG.limiteMax) return false;
      // Ignora refs da blacklist
      const ref = p.codigo ? p.codigo.split(/[-_]/)[0].replace(/^0+/, '') : '';
      if (CONFIG.ignorarRefs.includes(ref)) return false;
      // Se tem catálogo UpSeller, só alerta SKUs que vendem nos marketplaces
      if (skusUpseller && p.codigo) {
        if (!skusUpseller.has(p.codigo.toLowerCase())) return false;
      }
      return true;
    });

    const zerados = produtosAlerta.filter(p => p.saldoFisicoTotal === 0);
    const baixos = produtosAlerta.filter(p => p.saldoFisicoTotal > 0 && p.saldoFisicoTotal <= CONFIG.limiteMax);

    // Agrupa por família (referência = parte antes do primeiro "-")
    function agruparPorFamilia(lista) {
      const familias = new Map();
      for (const p of lista) {
        const partes = (p.codigo || '').split('-');
        const ref = partes[0] || 'SEM-REF';
        // Extrai nome limpo da descrição
        let nomeLimpo = (p.descricao || '').replace(/^[\d]+\s*[-_]\s*/, '').replace(/COR:\s*[^;]+;?/i, '').replace(/TAM(?:ANHO)?:\s*[^;]+;?/i, '').replace(/;/g, '').replace(/\s{2,}/g, ' ').trim();
        if (!familias.has(ref)) {
          familias.set(ref, { ref, nome: nomeLimpo, variantes: [] });
        }
        const cor = partes[1] || '-';
        const tam = partes[2] || '-';
        familias.get(ref).variantes.push({ cor, tam, estoque: p.saldoFisicoTotal, sku: p.codigo });
      }
      return Array.from(familias.values());
    }

    function formatarFamilia(familia) {
      const variantes = familia.variantes
        .map(v => `   ${v.cor}-${v.tam} (${v.estoque} un.)`)
        .join('\n');
      return `📦 *${familia.ref} — ${familia.nome}*\n${variantes}`;
    }

    const MAX_POR_SECAO = 20;
    const familiasZeradas = agruparPorFamilia(zerados);
    const familiasBaixas = agruparPorFamilia(baixos);

    const agora = new Date();
    const dataFormatada = agora.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const horaFormatada = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    let msg = '';
    msg += `━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    msg += `📊 *ALERTA DE ESTOQUE*\n`;
    msg += `📅 ${dataFormatada} às ${horaFormatada}\n`;
    msg += `━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    msg += `📋 *Resumo:*\n`;
    msg += `   🛑 ${zerados.length} produto${zerados.length !== 1 ? 's' : ''} zerado${zerados.length !== 1 ? 's' : ''}\n`;
    msg += `   🟡 ${baixos.length} produto${baixos.length !== 1 ? 's' : ''} com estoque baixo\n\n`;

    // SEÇÃO ZERADOS
    if (familiasZeradas.length > 0) {
      msg += `🛑 *ESTOQUE ZERADO — PRIORIDADE*\n`;
      msg += `─────────────────────────\n`;
      const exibir = familiasZeradas.slice(0, MAX_POR_SECAO);
      msg += exibir.map(f => formatarFamilia(f)).join('\n\n');
      if (familiasZeradas.length > MAX_POR_SECAO) {
        msg += `\n\n   _...e mais ${familiasZeradas.length - MAX_POR_SECAO} família(s) zerada(s)_`;
      }
      msg += `\n\n`;
    }

    // SEÇÃO BAIXOS
    if (familiasBaixas.length > 0) {
      msg += `🟡 *ESTOQUE BAIXO (1-${CONFIG.limiteMax} un.)*\n`;
      msg += `─────────────────────────\n`;
      const exibir = familiasBaixas.slice(0, MAX_POR_SECAO);
      msg += exibir.map(f => formatarFamilia(f)).join('\n\n');
      if (familiasBaixas.length > MAX_POR_SECAO) {
        msg += `\n\n   _...e mais ${familiasBaixas.length - MAX_POR_SECAO} família(s) com estoque baixo_`;
      }
      msg += `\n\n`;
    }

    if (zerados.length === 0 && baixos.length === 0) {
      msg += `✅ *Nenhum produto em alerta!* Estoque saudável.\n\n`;
    }

    msg += `━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    msg += `🕐 Última sync: ${ultimoCacheHora ? new Date(ultimoCacheHora).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }) : 'N/A'}\n`;
    msg += `📦 Total monitorado: ${cacheProdutos.length} SKUs\n`;
    if (skusUpseller) msg += `🏪 Filtro: Apenas marketplace (${skusUpseller.size} SKUs ativos)\n`;
    msg += `━━━━━━━━━━━━━━━━━━━━━━━━━`;

    const whatsappUrl = `https://api.whatsapp.com/send?text=${encodeURIComponent(msg)}`;

    res.json({
      sucesso: true,
      mensagem: msg,
      whatsappUrl,
      resumo: { zerados: zerados.length, baixos: baixos.length, total: produtosAlerta.length }
    });

  } catch (e) {
    console.error("❌ [Alerta] Erro ao gerar alerta:", e.message);
    res.status(500).json({ erro: "Erro ao gerar alerta de estoque." });
  }
});

// 📦 ROTA CACHE — Retorna catálogo instantaneamente do cache
app.get('/api/catalogo-cache', (req, res) => {
  if (cacheProdutos && cacheProdutos.length > 0) {
    return res.json({
      produtos: cacheProdutos,
      atualizadoEm: ultimoCacheHora,
      fonte: 'cache'
    });
  }
  res.json({ produtos: [], atualizadoEm: 0, fonte: 'vazio' });
});

// 🔄 ROTA SYNC — Força sincronização em segundo plano e retorna status
app.post('/api/catalogo-sync', async (req, res) => {
  if (sincronizandoCatalogo) {
    return res.json({ status: 'em_andamento', mensagem: 'Sincronização já em andamento.' });
  }
  // Dispara em segundo plano, não bloqueia a resposta
  sincronizarCatalogoEmSegundoPlano();
  res.json({ status: 'iniciado', mensagem: 'Sincronização iniciada em segundo plano.' });
});

// ──────────────────────────────────────────────
// 📤 UPLOAD DO CATÁLOGO UPSELLER (Excel do armazém)
// ──────────────────────────────────────────────
app.post('/api/upseller/upload-catalogo', upload.single('arquivo'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ erro: 'Nenhum arquivo enviado.' });
    }

    console.log(`\n📤 [UpSeller] Upload recebido: ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`);

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    if (!rows || rows.length === 0) {
      return res.status(400).json({ erro: 'Planilha vazia ou formato inválido.' });
    }

    // Detecta a coluna de SKU (primeira coluna que contenha "SKU" no header)
    const headers = Object.keys(rows[0]);
    const skuCol = headers.find(h => h.toUpperCase().includes('SKU')) || headers[0];

    const skus = [];
    for (const row of rows) {
      const sku = String(row[skuCol] || '').trim();
      if (sku && sku.length > 1) {
        skus.push(sku);
      }
    }

    if (skus.length === 0) {
      return res.status(400).json({ erro: 'Nenhum SKU encontrado na planilha.' });
    }

    salvarCatalogoUpSeller(skus);

    console.log(`   [UpSeller] ${skus.length} SKUs importados do catálogo.`);
    res.json({ sucesso: true, total: skus.length, amostra: skus.slice(0, 10) });

  } catch (e) {
    console.error('❌ [UpSeller] Erro no upload:', e.message);
    res.status(500).json({ erro: 'Erro ao processar arquivo: ' + e.message });
  }
});

// Rota para ver status do catálogo UpSeller
app.get('/api/upseller/catalogo-status', (req, res) => {
  const catalogo = lerCatalogoUpSeller();
  if (!catalogo || !catalogo.skus || catalogo.skus.length === 0) {
    return res.json({ temCatalogo: false, total: 0 });
  }
  res.json({
    temCatalogo: true,
    total: catalogo.skus.length,
    atualizadoEm: catalogo.atualizadoEm,
    amostra: catalogo.skus.slice(0, 5)
  });
});

app.get('/api/wms/produto/:codigo', async (req, res) => {
  try {
    const codigoBipado = req.params.codigo;
    const token = await obterAccessToken();
    let produto = null;

    // 1️⃣ Busca por código (SKU) no Bling
    try {
      const respCodigo = await axios.get(`https://www.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(codigoBipado)}`, { headers: { Authorization: `Bearer ${token}` } });
      if (respCodigo.data?.data?.length > 0) produto = respCodigo.data.data[0];
    } catch (e) {}

    // 2️⃣ Se não achou, busca por GTIN (código de barras / EAN)
    if (!produto) {
      try {
        const respGtin = await axios.get(`https://www.bling.com.br/Api/v3/produtos?gtin=${encodeURIComponent(codigoBipado)}`, { headers: { Authorization: `Bearer ${token}` } });
        if (respGtin.data?.data?.length > 0) produto = respGtin.data.data[0];
      } catch (e) {}
    }

    // 3️⃣ Se não achou no Bling, tenta o cache local (match por codigo ou gtin)
    if (!produto && cacheProdutos) {
      const termoLower = codigoBipado.toLowerCase();
      const cacheMatch = cacheProdutos.find(p =>
        String(p.codigo).toLowerCase() === termoLower ||
        String(p.gtin || '').toLowerCase() === termoLower
      );
      if (cacheMatch) {
        return res.json({ id: cacheMatch.id, codigo: cacheMatch.codigo, nome: cacheMatch.descricao, estoqueAtual: cacheMatch.saldoFisicoTotal || 0, fotoUrl: "" });
      }
    }

    if (!produto) {
      return res.status(404).json({ erro: 'Produto não encontrado' });
    }

    // Busca saldo de estoque do depósito SEDE
    let estoqueAtual = 0;
    try {
      const respEstoque = await axios.get(`https://www.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=${produto.id}`, { headers: { Authorization: `Bearer ${token}` } });
      const saldoData = respEstoque.data?.data?.[0];
      if (saldoData?.depositos && Array.isArray(saldoData.depositos)) {
        const depSede = saldoData.depositos.find(d => d.id === DEPOSITO_SEDE_ID);
        estoqueAtual = depSede?.saldoFisico ?? depSede?.saldoVirtual ?? 0;
      } else {
        estoqueAtual = saldoData?.saldoFisicoTotal || 0;
      }
    } catch (e) {}

    res.json({ id: produto.id, codigo: produto.codigo || codigoBipado, gtin: produto.gtin || '', nome: produto.nome, estoqueAtual, fotoUrl: produto.imagemURL || "" });
  } catch (error) {
    res.status(500).json({ erro: 'Erro interno' });
  }
});

// Rota para listar depósitos do Bling
app.get('/api/debug-depositos', async (req, res) => {
  if (process.env.NODE_ENV === 'production') return res.status(404).json({ erro: 'Não encontrado' });
  try {
    const token = await obterAccessToken();
    const resp = await axios.get('https://www.bling.com.br/Api/v3/depositos', { headers: { Authorization: `Bearer ${token}` } });
    res.json(resp.data);
  } catch (e) {
    res.status(500).json({ erro: e.message });
  }
});

app.post('/api/wms/entrada', async (req, res) => {
  try {
    const { idProduto, quantidade, operacao } = req.body;

    // 🛡️ Validação de entrada — evita valores inválidos
    const qtd = parseFloat(quantidade);
    const id = parseInt(idProduto);
    if (!id || isNaN(id) || id <= 0) return res.status(400).json({ erro: 'ID do produto inválido.' });
    if (!qtd || isNaN(qtd) || qtd <= 0 || qtd > 99999) return res.status(400).json({ erro: 'Quantidade inválida (deve ser entre 1 e 99999).' });

    const token = await obterAccessToken();
    const depositoId = DEPOSITO_SEDE_ID;
    const tipoOperacao = operacao === 'S' ? 'S' : 'E';
    await axios.post("https://www.bling.com.br/Api/v3/estoques", {
      produto: { id },
      deposito: { id: depositoId },
      operacao: tipoOperacao,
      quantidade: qtd,
      observacoes: tipoOperacao === 'E' ? "Entrada via WMS Local" : "Saída/Correção via WMS Local"
    }, { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 });
    atualizarCacheEstoque(id, qtd, tipoOperacao);
    res.json({ sucesso: true });
  } catch (error) {
    console.error(`⚠️ [WMS] Erro na movimentação:`, error.response?.data || error.message);
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

  // Detecta ngrok automaticamente e exibe links externos
  setTimeout(async () => {
    try {
      const resp = await fetch("http://127.0.0.1:4040/api/tunnels");
      const data = await resp.json();
      const tunnel = data.tunnels.find(t => t.proto === "https") || data.tunnels[0];
      if (tunnel) {
        const url = tunnel.public_url;
        console.log("");
        console.log("======================================================");
        console.log("   🌍 NGROK DETECTADO — LINKS EXTERNOS:");
        console.log("======================================================");
        console.log("");
        console.log(`   🌐 Dashboard: ${url}/dashboard.html`);
        console.log(`   📦 WMS:       ${url}/wms.html`);
        console.log(`   🛒 Checkout:  ${url}/checkout.html`);
        console.log("");
        console.log("   📋 Copie e envie para a equipe!");
        console.log("======================================================");
      }
    } catch {
      // ngrok não está rodando, sem problema
    }
  }, 2000);

  // 🔄 Sincroniza catálogo com Bling 5 segundos após iniciar
  setTimeout(() => {
    console.log("\n   [Auto-Sync] Atualizando catálogo em segundo plano...");
    sincronizarCatalogoEmSegundoPlano();
  }, 5000);

  // 🔄 Re-sincroniza a cada 30 minutos automaticamente
  setInterval(() => {
    console.log("\n   [Auto-Sync] Sincronização periódica (30 min)...");
    sincronizarCatalogoEmSegundoPlano();
  }, 30 * 60 * 1000);

  // 🧹 Limpa operações expiradas a cada 10 minutos (evita memory leak)
  setInterval(() => {
    const antes = operacoesFinalizadas.size;
    limparOperacoesExpiradas();
    const removidos = antes - operacoesFinalizadas.size;
    if (removidos > 0) console.log(`   [Cleanup] ${removidos} operação(ões) expirada(s) removida(s) da memória`);
  }, 10 * 60 * 1000);
});

// 🛡️ Captura erros não tratados para evitar crash total do servidor
process.on('uncaughtException', (err) => {
  console.error('❌ [ERRO FATAL NÃO CAPTURADO]:', err.message);
  console.error(err.stack);
});
process.on('unhandledRejection', (reason) => {
  console.error('❌ [PROMISE REJEITADA NÃO TRATADA]:', reason);
});

// WHATSAPP — DESATIVADO TEMPORARIAMENTE
// const wppClient = new Client({ authStrategy: new LocalAuth(), puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] } });
// wppClient.on("qr", (qr) => { qrcode.generate(qr, { small: true }); });
// wppClient.on("ready", () => { console.log("✅ [Robô] WhatsApp conectado!"); });
// wppClient.initialize();
console.log("ℹ️  [WhatsApp] Robô desativado temporariamente.");

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

// WhatsApp verificação desativada temporariamente
// async function executarVerificacao() { ... }
// cron.schedule("0 * * * *", executarVerificacao, { timezone: "America/Sao_Paulo" });