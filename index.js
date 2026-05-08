/**
 * ============================================================
 * AJ MODAS — ROBÔ DE ESTOQUE + DASHBOARD WEB + EXPORTAÇÃO UPSELLER
 * + MÓDULO DE CONFERÊNCIA (HÍBRIDO: BLING + PLANILHAS UPSELLER/BAGY + PDF)
 * ============================================================
 */

const archiver       = require('archiver');
const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode         = require("qrcode-terminal");
const cron           = require("node-cron");
const axios          = require("axios");
const fs             = require("fs");
const path           = require("path");
const express        = require("express");
const cors           = require("cors");
const xlsx           = require("xlsx"); 
const multer         = require("multer"); 
const pdfParse       = require("pdf-parse"); // Adicionado suporte a PDF

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
    nomeDoGrupo: "Estoque Marketplace",
    nomeDoGrupoSede: "Sede Giovana"
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

// 🔄 CONTROLE DE SINCRONIZAÇÃO — Timer gerenciável com estado exposto
const SYNC_INTERVALO_MS = 30 * 60 * 1000; // 30 minutos
const SYNC_COOLDOWN_MS = 2 * 60 * 1000;   // 2 min entre syncs manuais
const syncState = {
  timerRef: null,            // referência do setInterval atual
  ultimaSyncOk: 0,           // timestamp da última sync bem-sucedida
  ultimaSyncForçada: 0,      // timestamp da última sync manual (cooldown)
  proximaSync: 0,            // timestamp estimado da próxima sync automática
};

function iniciarTimerSync() {
  if (syncState.timerRef) clearInterval(syncState.timerRef);
  syncState.proximaSync = Date.now() + SYNC_INTERVALO_MS;
  syncState.timerRef = setInterval(() => {
    console.log("\n   [Auto-Sync] Sincronização periódica (30 min)...");
    syncState.proximaSync = Date.now() + SYNC_INTERVALO_MS;
    sincronizarCatalogoEmSegundoPlano();
  }, SYNC_INTERVALO_MS);
}

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
    syncState.ultimaSyncOk = Date.now();
    salvarCatalogoNoDisco(produtos);
    console.log(`   [Sync] Sincronização concluída! ${produtos.length} produtos atualizados.`);

    // 🚨 Dispara motor de alertas proativos após sync bem-sucedido
    avaliarEDispararAlertas().catch(e => console.error('   [AlertaManager] Erro:', e.message));
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
    "https://api.bling.com.br/Api/v3/oauth/token",
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
      const resp = await blingRequest("https://api.bling.com.br/Api/v3/produtos", accessToken, { pagina, limite: 100, tipo: 'T' });

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
      const resp = await blingRequest("https://api.bling.com.br/Api/v3/estoques/saldos", accessToken, params);
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

// 📊 Status da sincronização — alimenta cronômetro no front-end
app.get('/api/sync-status', (req, res) => {
    const agora = Date.now();
    const faltaMs = Math.max(0, syncState.proximaSync - agora);
    const minutos = Math.floor(faltaMs / 60000);
    const segundos = Math.floor((faltaMs % 60000) / 1000);

    res.json({
        ultimaSyncOk: syncState.ultimaSyncOk || null,
        ultimaSyncFormatada: syncState.ultimaSyncOk
            ? new Date(syncState.ultimaSyncOk).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })
            : 'nunca',
        proximaSyncMs: faltaMs,
        proximaSyncFormatada: `${minutos}m ${segundos}s`,
        sincronizando: sincronizandoCatalogo,
        totalProdutos: cacheProdutos ? cacheProdutos.length : 0
    });
});

// 🔴 Forçar sincronização manual (botão de pânico)
app.post('/api/forcar-sync', async (req, res) => {
    // Cooldown de 2 minutos entre syncs manuais
    const agora = Date.now();
    const tempoDesdeUltima = agora - syncState.ultimaSyncForçada;
    if (tempoDesdeUltima < SYNC_COOLDOWN_MS) {
        const restante = Math.ceil((SYNC_COOLDOWN_MS - tempoDesdeUltima) / 1000);
        return res.status(429).json({
            erro: `Aguarde ${restante}s antes de forçar nova sincronização.`,
            cooldownRestante: restante
        });
    }

    if (sincronizandoCatalogo) {
        return res.status(409).json({ erro: 'Sincronização já em andamento. Aguarde.' });
    }

    console.log('\n   [Sync] 🔴 Sincronização FORÇADA pela equipe!');
    syncState.ultimaSyncForçada = agora;

    // Reseta o timer para evitar sync duplo
    iniciarTimerSync();

    try {
        await sincronizarCatalogoEmSegundoPlano();
        res.json({
            ok: true,
            mensagem: 'Sincronização concluída!',
            totalProdutos: cacheProdutos ? cacheProdutos.length : 0,
            atualizadoEm: new Date(syncState.ultimaSyncOk).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }),
            proximaSyncMs: syncState.proximaSync - Date.now()
        });
    } catch (e) {
        console.error('   [Sync] Falha na sync forçada:', e.message);
        res.status(500).json({ erro: 'Falha na sincronização: ' + e.message });
    }
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
                const url = `https://api.bling.com.br/Api/v3/pedidos/vendas?dataInicial=${dataInicial}&pagina=${pagina}&limite=100`;
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
                const resp = await axios.get(`https://api.bling.com.br/Api/v3/pedidos/vendas?numeroLoja=${numero}`, { headers: { Authorization: `Bearer ${token}` } });
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
        const respDetalhes = await axios.get(`https://api.bling.com.br/Api/v3/pedidos/vendas/${pedidoId}`, { headers: { Authorization: `Bearer ${token}` } });
        
        const itensBling = respDetalhes.data.data.itens.map(i => {
            const sku = i.codigo || i.produto?.codigo || "S/COD";
            const produtoId = i.produto?.id || null;
            // Busca GTIN e código Bling no cache para conferência por código de barras
            let gtin = '';
            let codigoBling = '';
            if (cacheProdutos && produtoId) {
                const cached = cacheProdutos.find(p => p.id === produtoId);
                if (cached) {
                    gtin = cached.gtin || '';
                    codigoBling = cached.codigo || '';
                }
            }
            // Se não achou por ID, tenta por código (com e sem zeros à esquerda)
            if (!codigoBling && cacheProdutos) {
                const skuLower = String(sku).toLowerCase();
                const skuSemZeros = skuLower.replace(/^0+/, '') || '0';
                const cached = cacheProdutos.find(p => {
                    const cod = String(p.codigo).toLowerCase();
                    return cod === skuLower || cod.replace(/^0+/, '') === skuSemZeros;
                });
                if (cached) {
                    gtin = gtin || cached.gtin || '';
                    codigoBling = cached.codigo || '';
                }
            }
            return {
                sku,
                gtin,
                codigoBling,
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
                        const respBusca = await axios.get(`https://api.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(troca.originalSku)}`, { headers: { Authorization: `Bearer ${token}` }});
                        if (respBusca.data?.data?.length > 0) {
                            troca.originalProdutoId = respBusca.data.data[0].id;
                            console.log(`   🔍 Resolvido ID do item original "${troca.originalSku}" → ${troca.originalProdutoId}`);
                        }
                    } catch (e) {}
                }
                // ENTRADA do item removido (devolver ao estoque)
                if (troca.originalProdutoId) {
                    try {
                        await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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
                        await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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
                        const resp = await axios.get(`https://api.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(item.sku)}`, { headers: { Authorization: `Bearer ${token}` }});
                        if (resp.data?.data?.length > 0) prodId = resp.data.data[0].id;
                    } catch (e) {}
                }
                if (prodId) {
                    try {
                        await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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
                        await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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

            const respPedido = await axios.get(`https://api.bling.com.br/Api/v3/pedidos/vendas/${id}`, {
                headers: { Authorization: `Bearer ${token}` }
            });

            let dadosPedido = respPedido.data.data;

            dadosPedido.loja = { id: 205344151 };       // ID da Loja SEDE
            dadosPedido.vendedor = { id: 15596386514 }; // ID do Vendedor SITE

            try {
                await axios.put(`https://api.bling.com.br/Api/v3/pedidos/vendas/${id}`, dadosPedido, {
                    headers: { Authorization: `Bearer ${token}` }
                });
                console.log(`✅ Loja e Vendedor atualizados com sucesso no Bling!`);
            } catch (errPut) {
                console.error(`⚠️ Erro ao injetar Loja/Vendedor (mas o pacote será finalizado mesmo assim).`);
            }

            try {
                await axios.patch(`https://api.bling.com.br/Api/v3/pedidos/vendas/${id}/situacoes/9`, {}, {
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
                        const respProd = await axios.get(`https://api.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(item.sku)}`, { headers: { Authorization: `Bearer ${token}` }});
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
                                const respRef = await axios.get(`https://api.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(parsed.ref)}`, { headers: { Authorization: `Bearer ${token}` }});
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
                        await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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
// 🟢 EXPORTAÇÃO UPSELLER (LOTE MULTI-REF + AUDITORIA) - BLINDADO
// ──────────────────────────────────────────────
app.get('/api/exportar-upseller', async (req, res) => {
    req.setTimeout(120000);
    res.setTimeout(120000);
    try {
        // REGRA 1: CORTE DE REQUISIÇÃO (VELOCIDADE)
        // Usa exclusivamente a variável global de cache (NENHUMA CHAMADA AO BLING AQUI)
        if (!cacheProdutos || cacheProdutos.length === 0) {
            return res.status(400).json({ erro: "Cache de produtos vazio. Aguarde a sincronização em segundo plano." });
        }

        const LIMIAR_SEGURANCA = 30;
        
        // Extrai o array de referências (ex: ?refs=33,108,45) - LIMPO (sem zeros a esquerda)
        const refsQuery = req.query.refs 
            ? req.query.refs.split(',').map(r => r.trim().replace(/^0+/, '')).filter(r => r) 
            : null;
        
        console.log(`\n📦 [UpSeller] Iniciando Exportação... Limiar: ${LIMIAR_SEGURANCA}un`);
        if (refsQuery) console.log(`   [UpSeller] 🎯 LOTE ATIVADO (Barreira de Ferro): [${refsQuery.join(', ')}]`);

        const catalogoUpSeller = lerCatalogoUpSeller();
        if (!catalogoUpSeller || !catalogoUpSeller.skus || catalogoUpSeller.skus.length === 0) {
            console.log(`[Aviso] Tentativa de exportação, mas o catálogo base da UpSeller não foi carregado na memória.`);
            return res.status(400).json({ erro: "O catálogo base da UpSeller não está carregado. Faça o upload da planilha primeiro." });
        }
        
        const skusUpSellerMap = new Map();
        if (catalogoUpSeller && catalogoUpSeller.skus) {
            catalogoUpSeller.skus.forEach(item => {
                const sku = typeof item === 'string' ? item : item.sku;
                const armazem = typeof item === 'string' ? '' : item.armazem;
                const parsed = parseUpSellerSku(sku);
                if (parsed) skusUpSellerMap.set(normalizarChaveMatch(parsed.ref, parsed.cor, parsed.tam), { sku, armazem });
            });
        }

        // --- HIGIENE E QUARENTENA ---
        // REGRA 2: QUARENTENA DE DUPLICATAS
        // REGRA 3: SKUs ÓRFÃOS (GG ESCONDIDO)
        const blingAgrupado = new Map();
        for (const p of cacheProdutos) {
            if (!p.descricao) continue;
            const matchRef = p.descricao.match(/^(\d+)/);
            if (!matchRef) continue;
            const ref = matchRef[1].replace(/^0+/, '');
            
            // Filtra pela Barreira de Ferro
            if (refsQuery && refsQuery.length > 0 && !refsQuery.includes(ref)) continue;
            
            const matchCor = p.descricao.match(/\bCOR[:\s]+([^,;]+)/i);
            const cor = matchCor ? matchCor[1].trim() : '';
            const matchTam = p.descricao.match(/\bTAM(?:ANHO)?[:\s]+([^,;\s]+)/i);
            const tam = matchTam ? matchTam[1].trim() : '';
            
            // Só agrupa variações com cor
            if (!cor) continue;

            const chave = normalizarChaveMatch(ref, cor, tam);
            if (!blingAgrupado.has(chave)) blingAgrupado.set(chave, []);
            blingAgrupado.get(chave).push(p);
        }

        const anomaliasDuplicatas = new Set();
        let logAnomalias = `\n⚠️ ANOMALIAS DETECTADAS (AÇÃO NECESSÁRIA NO BLING/UPSELLER)\n----------------------------------------------------\n`;
        let logOrfaos = `\n⚠️ AÇÃO NECESSÁRIA: SKUS NÃO ENCONTRADOS NA UPSELLER (Cadastrar ou Corrigir Nome)\n---------------------------------------------------------------------------------\n`;
        let logSemArmazem = `\n⚠️ AÇÃO NECESSÁRIA: SKUS SEM ARMAZÉM VINCULADO (Entrar na UpSeller e vincular Galpão)\n---------------------------------------------------------------------------------\n`;
        let qtdAnomalias = 0;
        let qtdOrfaos = 0;
        let qtdSemArmazem = 0;

        const mapaEstoqueReal = new Map();

        for (const [chave, listaBling] of blingAgrupado.entries()) {
            const ativos = listaBling.filter(p => p.tipo === 'V' || (p.tipo === 'P' && listaBling.length === 1));
            
            // REGRA 2: QUARENTENA DE DUPLICATAS
            if (ativos.length > 1) {
                qtdAnomalias++;
                anomaliasDuplicatas.add(chave);
                const codigos = ativos.map(v => v.codigo).join(', ');
                const nomeVisual = ativos[0].descricao;
                logAnomalias += `- DUPLICIDADE: "${nomeVisual}" possui ${ativos.length} SKUs ativos no Bling (${codigos}). Saldo zerado por segurança.\n`;
                mapaEstoqueReal.set(chave, 0); // Força 0 na quarentena
            } else if (ativos.length === 1) {
                const estoqueReal = parseInt(ativos[0].saldoFisicoTotal) || 0;
                mapaEstoqueReal.set(chave, estoqueReal);
                
                const upSellerInfo = skusUpSellerMap.get(chave);
                if (!upSellerInfo) {
                    if (estoqueReal > 0) {
                        qtdOrfaos++;
                        logOrfaos += `- FALTA NA UPSELLER: "${ativos[0].descricao}" existe no Bling com ${estoqueReal} peças, mas não no catálogo UpSeller.\n`;
                    }
                } else if (!upSellerInfo.armazem) {
                    qtdSemArmazem++;
                    logSemArmazem += `- SEM ARMAZÉM: O SKU "${upSellerInfo.sku}" ("${ativos[0].descricao}") está na UpSeller mas a coluna Armazém está vazia.\n`;
                }
            }
        }

        // --- GERAÇÃO DA PLANILHA ---
        const dadosPlanilha = [[
            "SKU*",
            "Estoque Baixo\n(Não será atualizado se não for preenchido)",
            "Qtd. Total Atualizado\n(Não será atualizado se não for preenchido)",
            "Custo Médio Atualizado\n(Não será atualizado se não for preenchido)"
        ]];

        let logRepostos = `\n🟢 PEÇAS REPOSTAS (MÁSCARA ATIVADA | ESTOQUE >= ${LIMIAR_SEGURANCA})\n----------------------------------------------------\n`;
        let logBaixos = `\n🔴 PEÇAS BAIXAS/ZERADAS (SALDO REAL | ESTOQUE < ${LIMIAR_SEGURANCA})\n----------------------------------------------------\n`;
        let logKits = `\n⚠️ KITS (ENVIADO PADRÃO 100)\n----------------------------------------------------\n`;
        let logSemMatch = `\n❌ SEM MATCH (ZERADOS POR SEGURANÇA)\n----------------------------------------------------\n`;

        let qtdAtivo = 0, qtdZerado = 0, qtdIgnorados = 0, qtdSemMatch = 0, qtdKit = 0;

        // Função utilitária para extrair a REF estrita
        const extrairRefEstrita = (sku) => {
            if (!sku) return null;
            const partes = sku.split('-');
            if (partes.length === 0) return null;
            return partes[0].trim().replace(/^0+/, '');
        };

        if (catalogoUpSeller && catalogoUpSeller.skus) {
            for (const item of catalogoUpSeller.skus) {
                const skuReal = typeof item === 'string' ? item : item.sku;
                const armazem = typeof item === 'string' ? '' : item.armazem;

                if (refsQuery && refsQuery.length > 0) {
                    const temMatch = refsQuery.some(ref => {
                        // Regex inteligente: ignora zeros à esquerda, e exige que o próximo caractere após a ref não seja um número (evita que 3 ache 31)
                        const regex = new RegExp(`^0*${ref}(?:[^0-9]|$)`, 'i');
                        return regex.test(skuReal);
                    });
                    if (!temMatch) continue;
                }

                const parsed = parseUpSellerSku(skuReal);
                if (!parsed) {
                    qtdIgnorados++;
                    continue;
                }

                const chaveUpSeller = normalizarChaveMatch(parsed.ref, parsed.cor, parsed.tam);

                if (anomaliasDuplicatas.has(chaveUpSeller) || !armazem) {
                    dadosPlanilha.push([skuReal, "", 0, ""]);
                    continue;
                }

                if (!mapaEstoqueReal.has(chaveUpSeller)) {
                    if (/kit/i.test(skuReal)) {
                        qtdKit++;
                        dadosPlanilha.push([skuReal, "", 100, ""]);
                        logKits += `- ${skuReal} -> Enviado: 100\n`;
                    } else {
                        qtdSemMatch++;
                        dadosPlanilha.push([skuReal, "", 0, ""]);
                        logSemMatch += `- ${skuReal} -> Enviado: 0\n`;
                    }
                    continue;
                }

                const quantidadeReal = mapaEstoqueReal.get(chaveUpSeller);
                let quantidadeUpSeller = 0;

                if (/kit/i.test(skuReal)) {
                    quantidadeUpSeller = 100;
                    qtdKit++;
                    logKits += `- ${skuReal} | Real: ${quantidadeReal} -> Enviado: 100\n`;
                } else if (quantidadeReal >= LIMIAR_SEGURANCA) {
                    quantidadeUpSeller = 2000 + quantidadeReal;
                    qtdAtivo++;
                    logRepostos += `- ${skuReal.padEnd(25)} | Real: ${quantidadeReal.toString().padStart(3)} -> Enviado: ${quantidadeUpSeller}\n`;
                } else {
                    quantidadeUpSeller = quantidadeReal;
                    qtdZerado++;
                    logBaixos += `- ${skuReal.padEnd(25)} | Real: ${quantidadeReal.toString().padStart(3)} -> Enviado: ${quantidadeUpSeller}\n`;
                }

                dadosPlanilha.push([skuReal, "", quantidadeUpSeller, ""]);
            }
        }

        const totalExportados = qtdAtivo + qtdZerado + qtdSemMatch + qtdKit;
        if (refsQuery && totalExportados === 0) {
             return res.status(404).json({ erro: `A referência ${refsQuery.join(', ')} não foi localizada no catálogo.` });
        }

        // REGRA 4: NOVO RELATÓRIO TXT
        let relatorioFinal = `====================================================\n`;
        relatorioFinal += `📊 RELATÓRIO DE AUDITORIA E EXPORTAÇÃO UPSELLER\n`;
        relatorioFinal += `====================================================\n`;
        relatorioFinal += `Data da Geração: ${new Date().toLocaleString('pt-BR')}\n`;
        relatorioFinal += `Lote Processado: ${refsQuery ? refsQuery.join(', ') : 'Catálogo Completo'}\n`;
        relatorioFinal += `Regra Base.....: >= ${LIMIAR_SEGURANCA} (Ativa Máscara +2000) | < ${LIMIAR_SEGURANCA} (Envia Real)\n`;
        relatorioFinal += `====================================================\n\n`;
        
        relatorioFinal += `RESUMO ESTATÍSTICO:\n`;
        relatorioFinal += `- Peças Repostas (Máscara ativa).....: ${qtdAtivo}\n`;
        relatorioFinal += `- Peças Baixas/Zeradas (Saldo real)..: ${qtdZerado}\n`;
        relatorioFinal += `- Peças tipo Kit.....................: ${qtdKit}\n`;
        relatorioFinal += `- Ignorados (Sem match/Inválidos)....: ${qtdSemMatch}\n`;
        relatorioFinal += `- Anomalias (Duplicatas/Órfãos)......: ${qtdAnomalias + qtdOrfaos}\n`;
        relatorioFinal += `- SKUs Sem Armazém...................: ${qtdSemArmazem}\n`;
        relatorioFinal += `----------------------------------------------------\n`;

        if (qtdOrfaos > 0) relatorioFinal += logOrfaos;
        if (qtdSemArmazem > 0) relatorioFinal += logSemArmazem;
        if (qtdAnomalias > 0) relatorioFinal += logAnomalias;
        if (qtdAtivo > 0) relatorioFinal += logRepostos;
        if (qtdZerado > 0) relatorioFinal += logBaixos;
        if (qtdKit > 0) relatorioFinal += logKits;
        if (qtdSemMatch > 0) relatorioFinal += logSemMatch;

        const worksheet = xlsx.utils.aoa_to_sheet(dadosPlanilha);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
        const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        const dataAtual = new Date().toISOString().slice(0,10);
        const nomeRef = refsQuery ? `_Lote` : `_Completo`;
        
        res.setHeader('Content-Disposition', `attachment; filename="UpSeller_Exportacao${nomeRef}_${dataAtual}.zip"`);
        res.setHeader('Content-Type', 'application/zip');

        const archive = archiver('zip', { zlib: { level: 9 } });
        archive.on('error', function(err) { throw err; });
        archive.pipe(res);
        
        archive.append(excelBuffer, { name: `Update_warehouse_SKU${nomeRef}_${dataAtual}.xlsx` });
        archive.append(relatorioFinal, { name: `Relatorio_Auditoria_${dataAtual}.txt` });
        
        await archive.finalize();

        console.log(`✅ [UpSeller] ZIP com Relatório de Auditoria gerado!`);

    } catch (e) {
        console.error("❌ Erro ao exportar ZIP UpSeller:", e.message);
        res.status(500).json({ erro: "Erro interno ao gerar o pacote ZIP." });
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
    const skusUpseller = catalogoUpseller?.skus ? new Set(catalogoUpseller.skus.map(s => (typeof s === 'string' ? s : s.sku).toLowerCase())) : null;

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
    const armazemCol = headers.find(h => h.toUpperCase().includes('ARMAZ') || h.toUpperCase().includes('WAREHOUSE'));

    const skus = [];
    for (const row of rows) {
      const sku = String(row[skuCol] || '').trim();
      const armazem = armazemCol ? String(row[armazemCol] || '').trim() : '';
      if (sku && sku.length > 1) {
        skus.push({ sku, armazem });
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
      const respCodigo = await axios.get(`https://api.bling.com.br/Api/v3/produtos?codigo=${encodeURIComponent(codigoBipado)}`, { headers: { Authorization: `Bearer ${token}` } });
      if (respCodigo.data?.data?.length > 0) produto = respCodigo.data.data[0];
    } catch (e) {}

    // 2️⃣ Se não achou, busca por GTIN (código de barras / EAN)
    if (!produto) {
      try {
        const respGtin = await axios.get(`https://api.bling.com.br/Api/v3/produtos?gtin=${encodeURIComponent(codigoBipado)}`, { headers: { Authorization: `Bearer ${token}` } });
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
      const respEstoque = await axios.get(`https://api.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=${produto.id}`, { headers: { Authorization: `Bearer ${token}` } });
      const saldoData = respEstoque.data?.data?.[0];
      if (saldoData?.depositos && Array.isArray(saldoData.depositos)) {
        const depSede = saldoData.depositos.find(d => d.id === DEPOSITO_SEDE_ID);
        if (depSede) {
          estoqueAtual = depSede?.saldoFisico ?? depSede?.saldoVirtual ?? 0;
        } else {
          estoqueAtual = saldoData?.saldoFisicoTotal || 0;
        }
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
    const resp = await axios.get('https://api.bling.com.br/Api/v3/depositos', { headers: { Authorization: `Bearer ${token}` } });
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
    await axios.post("https://api.bling.com.br/Api/v3/estoques", {
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

app.post('/api/wms/entrada-lote', async (req, res) => {
  try {
    const { lote, operacao } = req.body;
    if (!lote || !Array.isArray(lote) || lote.length === 0) {
      return res.status(400).json({ erro: 'Lote vazio ou inválido.' });
    }

    const tipoOperacao = operacao === 'S' ? 'S' : 'E';
    const depositoId = DEPOSITO_SEDE_ID;
    
    // Constrói array no formato exato da API V3
    const payloadBling = lote.map(item => ({
      produto: { id: parseInt(item.id) },
      deposito: { id: depositoId },
      operacao: tipoOperacao,
      quantidade: parseFloat(item.quantidade),
      observacoes: tipoOperacao === 'E' ? "Entrada em Lote via WMS Local" : "Saída/Correção em Lote via WMS Local"
    }));

    const token = await obterAccessToken();
    
    await axios.post("https://api.bling.com.br/Api/v3/estoques", payloadBling, { 
      headers: { Authorization: `Bearer ${token}` }, 
      timeout: 30000 // Timeout maior para requisição em lote
    });

    // Atualiza cache local instantaneamente
    lote.forEach(item => {
      atualizarCacheEstoque(parseInt(item.id), parseFloat(item.quantidade), tipoOperacao);
    });

    res.json({ sucesso: true, total: lote.length });
  } catch (error) {
    console.error(`⚠️ [WMS] Erro na movimentação em lote:`, error.response?.data || error.message);
    res.status(500).json({ erro: 'Erro ao salvar lote no Bling' });
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

  // 🔄 Re-sincroniza a cada 30 minutos automaticamente (com timer gerenciável)
  iniciarTimerSync();

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

// ══════════════════════════════════════════════════
// 📱 WHATSAPP — ALERTA DE ESTOQUE MARKETPLACE
// ══════════════════════════════════════════════════

const DEPOSITO_BASE_ID = 14887397316;

// Referências monitoradas (lista fixa de produtos prioritários)
const REFS_MONITORADAS = [
  '08', '15', '29', '31', '54', '104', '107', '108',
  '110', '129', '138', '146', '160', '163', '171', '172', '176', '188'
];

const MINIMO_SEDE = 30;
const MINIMO_BASE = 60;

// Thresholds de alerta por depósito (do mais crítico ao menos crítico)
const THRESHOLDS_SEDE = [0, 15];
const THRESHOLDS_BASE = [0, 30];

// Grupos WhatsApp para alertas automáticos
const GRUPO_ALERTA_SEDE = 'Estoque Marketplace';
const GRUPO_ALERTA_BASE = 'Estoque Base de Reposição';

const wppClient = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] }
});

wppClient.on("qr", (qr) => {
  console.log("\n📱 [WhatsApp] Escaneie o QR Code abaixo para conectar:");
  qrcode.generate(qr, { small: true });
});

wppClient.on("ready", () => {
  console.log("✅ [WhatsApp] Conectado ao grupo 'Estoque Marketplace'!");
});

wppClient.on("disconnected", async (reason) => {
  console.log(`⚠️ [SRE] Conexão perdida. Motivo: ${reason}. Tentando Auto-Healing em 10s...`);
  try { await wppClient.destroy(); } catch (e) { /* ignore */ }
  setTimeout(async () => {
      console.log(`[SRE] Auto-Healing acionado. Reinicializando bot.`);
      try { await wppClient.initialize(); } catch (e) { console.error(`[SRE] Falha no Auto-Healing:`, e.message); }
  }, 10000);
});

wppClient.on("auth_failure", async (msg) => {
  console.error(`⚠️ [SRE] Falha de autenticação. Tentando Auto-Healing em 10s... | Detalhe:`, msg);
  try { await wppClient.destroy(); } catch (e) { /* ignore */ }
  setTimeout(async () => {
      console.log(`[SRE] Auto-Healing acionado. Reinicializando bot após falha de auth.`);
      try { await wppClient.initialize(); } catch (e) { console.error(`[SRE] Falha no Auto-Healing:`, e.message); }
  }, 10000);
});

// 📩 Listener de comandos no grupo WhatsApp (registrado ANTES do initialize)
let _processandoComando = false;

// Grupos autorizados a usar o comando !estoque baseado em cache
const GRUPOS_COMANDO_CACHE = [
  'Estoque Marketplace',
  'Estoque Base de Reposição'
];
// IDs imutáveis (deixar vazio e preencher via console.log do DEBUG para transição segura)
const GRUPOS_CACHE_IDS = [];

wppClient.on('message_create', async (msg) => {
  try {
    if (!msg.body || msg.body.trim() === '') return;

    const texto = msg.body.trim().toLowerCase();
    if (!texto.startsWith('!estoque')) return;

    if (_processandoComando) return;

    const chat = await msg.getChat();
    if (!chat.isGroup) return;

    // --- REGRAS SRE: LOG DE TRANSIÇÃO E IDs IMUTÁVEIS ---
    console.log(`[DEBUG WhatsApp] Grupo: "${chat.name}" | ID Oficial: ${chat.id._serialized}`);

    const idGrupoRemetente = chat.id._serialized;
    const isGrupoSede = (chat.name === CONFIG.whatsapp.nomeDoGrupoSede) || (CONFIG.whatsapp.idGrupoSede && idGrupoRemetente === CONFIG.whatsapp.idGrupoSede);
    const isGrupoCache = GRUPOS_COMANDO_CACHE.includes(chat.name) || GRUPOS_CACHE_IDS.includes(idGrupoRemetente);

    const partes = texto.split(/\s+/);
    const arg = partes[1] || null;

    // ─────────────────────────────────────────────────────────
    // 🏪 GRUPO "Sede Giovana" — consulta em TEMPO REAL na API
    // ─────────────────────────────────────────────────────────
    if (isGrupoSede) {
      if (!arg) {
        await chat.sendMessage('📋 *Uso:* !estoque <referência>\nEx: !estoque 180');
        return;
      }

      console.log(`\n📩 [WhatsApp/Sede] Comando recebido: "${msg.body}" — ref: ${arg}`);
      _processandoComando = true;

      await chat.sendMessage(`⏳ Consultando estoque ref ${arg} em tempo real...`);

      const resultado = await buscarEstoqueBlingTempoReal(arg);

      if (!resultado) {
        await chat.sendMessage(`⚠️ Nenhuma variação encontrada para referência "${arg}".`);
        return;
      }

      const mensagem = formatarEstoqueSede(resultado.nomeProduto, resultado.referencia, resultado.porTamanho);
      await chat.sendMessage(mensagem);

      console.log(`   [WhatsApp/Sede] ✅ Ref ${arg}: ${resultado.totalVariacoes} variações enviadas.`);
      return;
    }

    // ─────────────────────────────────────────────────────────
    // 📦 GRUPOS COM CACHE — Marketplace + Base de Reposição
    // ─────────────────────────────────────────────────────────
    if (!isGrupoCache) return;

    console.log(`\n📩 [WhatsApp] Comando recebido: "${msg.body}"`);
    _processandoComando = true;

    if (!cacheProdutos || cacheProdutos.length === 0) {
      await chat.sendMessage('⚠️ Cache de produtos vazio. Aguarde a sincronização.');
      return;
    }

    // Determina filtro: !estoque [sede|base] [<ref>]
    let filtroDeposito = null;
    let filtroRef = null;

    if (arg === 'sede' || arg === 'base') {
      filtroDeposito = arg;
      if (partes[2]) filtroRef = partes[2]; // !estoque sede 104 / !estoque base 104
    } else if (arg) {
      filtroRef = arg; // !estoque 104
    }

    // Valida referência
    if (filtroRef && !REFS_MONITORADAS.includes(filtroRef)) {
      await chat.sendMessage(`⚠️ Referência "${filtroRef}" não monitorada.\n\n📋 *Refs:* ${REFS_MONITORADAS.join(', ')}`);
      return;
    }

    const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

    // ─── NOVO: Comando específico !estoque sede <ref> ou !estoque base <ref>
    // Resposta = bloco isolado do modelo solicitado, sem misturar outras refs
    if (filtroDeposito && filtroRef) {
      await chat.sendMessage(`⏳ Consultando ${filtroDeposito.toUpperCase()} ref ${filtroRef}...`);

      let saldosBaseConsulta = null;
      if (filtroDeposito === 'base') {
        try {
          const token = await obterAccessToken();
          const variacoesRef = cacheProdutos.filter(p => {
            const r = extrairRef(p.descricao);
            return r === filtroRef && /COR:/i.test(p.descricao || '');
          });
          if (variacoesRef.length === 0) {
            await chat.sendMessage(`⚠️ Nenhuma variação encontrada para ref ${filtroRef}.`);
            return;
          }
          saldosBaseConsulta = await buscarSaldosPorDeposito(token, variacoesRef.map(p => p.id), DEPOSITO_BASE_ID);
        } catch (e) {
          await chat.sendMessage('⚠️ Erro ao buscar saldos BASE.');
          return;
        }
      }

      const bloco = montarBlocoModelo(filtroRef, filtroDeposito, saldosBaseConsulta, `📋 *CONSULTA ${dataHora}*`);
      if (bloco) {
        await chat.sendMessage(bloco);
        console.log(`   [WhatsApp] ✅ Bloco enviado: ${filtroDeposito.toUpperCase()} ref ${filtroRef}`);
      } else {
        await chat.sendMessage(`⚠️ Nenhuma variação encontrada para ref ${filtroRef}.`);
      }
      return;
    }

    await chat.sendMessage(`⏳ Consultando estoque...`);

    // Gera alertas agrupados por conjunto (visão geral, sem ref específica)
    const { conjuntosSede, conjuntosBase } = await gerarAlertasPorConjunto(filtroRef, filtroDeposito);

    if (!filtroDeposito || filtroDeposito === 'sede') {
      if (conjuntosSede.length > 0) {
        await enviarPorLotes(chat, `🏢 *ESTOQUE SEDE — ${dataHora}*\n📊 ${conjuntosSede.length} conjunto(s) abaixo de ${MINIMO_SEDE} pçs/variação`, conjuntosSede);
      } else {
        await chat.sendMessage(`✅ *SEDE:* Todos os estoques monitorados OK!`);
      }
      if (!filtroDeposito) await delay(2000);
    }

    if (!filtroDeposito || filtroDeposito === 'base') {
      if (conjuntosBase.length > 0) {
        await enviarPorLotes(chat, `🏭 *ESTOQUE BASE — ${dataHora}*\n📊 ${conjuntosBase.length} conjunto(s) abaixo de ${MINIMO_BASE} pçs/variação`, conjuntosBase);
      } else {
        await chat.sendMessage(`✅ *BASE:* Todos os estoques monitorados OK!`);
      }
    }

    console.log(`   [WhatsApp] ✅ Comando processado. SEDE: ${conjuntosSede.length} conjuntos | BASE: ${conjuntosBase.length} conjuntos`);
  } catch (e) {
    console.error('   [WhatsApp] Erro ao processar comando:', e.message);
  } finally {
    _processandoComando = false;
  }
});

// Inicializa WhatsApp com retry (Puppeteer pode perder o contexto na primeira conexão)
(async () => {
  const MAX_TENTATIVAS = 3;
  for (let tentativa = 1; tentativa <= MAX_TENTATIVAS; tentativa++) {
    try {
      await wppClient.initialize();
      break;
    } catch (err) {
      console.error(`   [WhatsApp] Falha ao inicializar (tentativa ${tentativa}/${MAX_TENTATIVAS}):`, err.message);
      if (tentativa < MAX_TENTATIVAS) {
        const espera = tentativa * 5000;
        console.log(`   [WhatsApp] Tentando novamente em ${espera / 1000}s...`);
        await delay(espera);
      } else {
        console.error('   [WhatsApp] ⚠️ Não foi possível inicializar após 3 tentativas. O restante do sistema continua funcionando.');
      }
    }
  }
})();

// ══════════════════════════════════════════════════
// 🏪 SEDE GIOVANA — BUSCA DE ESTOQUE EM TEMPO REAL
// ══════════════════════════════════════════════════

// Busca estoque direto na API do Bling (sem cache) para uma referência específica
async function buscarEstoqueBlingTempoReal(referencia) {
  const token = await obterAccessToken();

  // 1) Busca produtos cuja descrição começa com a referência
  const todosProdutos = [];
  let pagina = 1;
  let temMais = true;

  while (temMais) {
    const resp = await blingRequest("https://api.bling.com.br/Api/v3/produtos", token, {
      nome: `${referencia}-`,
      tipo: 'T',
      limite: 100,
      pagina
    });
    const data = resp.data?.data ?? [];
    todosProdutos.push(...data);
    if (data.length < 100) temMais = false;
    else pagina++;
    if (temMais) await delay(500);
  }

  // 2) Filtra: somente variações da referência exata (com COR: na descrição)
  const variacoes = todosProdutos.filter(p => {
    const desc = p.nome || p.descricao || '';
    const ref = extrairRef(desc);
    return ref === referencia && /COR:/i.test(desc);
  });

  if (variacoes.length === 0) return null;

  // 3) Busca saldos do depósito SEDE em tempo real
  const saldoMap = new Map();
  for (let i = 0; i < variacoes.length; i += 50) {
    const batch = variacoes.slice(i, i + 50);
    const params = new URLSearchParams();
    for (const p of batch) params.append('idsProdutos[]', p.id);

    const resp = await blingRequest("https://api.bling.com.br/Api/v3/estoques/saldos", token, params);
    const saldos = resp.data?.data || [];

    for (const s of saldos) {
      const idProd = s.produto?.id;
      if (!idProd) continue;
      let saldo = 0;
      if (s.depositos && Array.isArray(s.depositos)) {
        const dep = s.depositos.find(d => d.id === DEPOSITO_SEDE_ID);
        if (dep) saldo = dep.saldoFisico ?? dep.saldoVirtual ?? 0;
      }
      saldoMap.set(idProd, saldo);
    }
    if (i + 50 < variacoes.length) await delay(500);
  }

  // 4) Agrupa por tamanho → cor
  const nomeProduto = extrairNomeLimpo(variacoes[0].nome || variacoes[0].descricao || '')
    .replace(/^\d+[-\s]*/, '').trim();

  const porTamanho = new Map();
  for (const p of variacoes) {
    const desc = p.nome || p.descricao || '';
    const { cor, tam } = extrairCorTam(desc);
    const saldo = saldoMap.get(p.id) || 0;
    if (!porTamanho.has(tam)) porTamanho.set(tam, []);
    porTamanho.get(tam).push({ cor, saldo });
  }

  return { nomeProduto, referencia, porTamanho, totalVariacoes: variacoes.length };
}

// Formata a resposta de estoque para o grupo Sede Giovana (formato limpo, sem emoji)
function formatarEstoqueSede(nomeProduto, ref, porTamanho) {
  let msg = `${nomeProduto}\n\nREF: -${ref}\n`;

  const ordemTamanhos = ['PP', 'P', 'M', 'G', 'GG', 'XG', 'XXG', 'EG', 'EGG'];
  const tamanhos = [...porTamanho.keys()].sort((a, b) => {
    const ia = ordemTamanhos.indexOf(a.toUpperCase());
    const ib = ordemTamanhos.indexOf(b.toUpperCase());
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  for (const tam of tamanhos) {
    const cores = porTamanho.get(tam).filter(v => v.saldo > 0).sort((a, b) => a.cor.localeCompare(b.cor));
    if (cores.length === 0) continue;
    msg += `\n${tam}\n`;
    for (const { cor, saldo } of cores) {
      msg += `${cor} ${saldo}\n`;
    }
  }

  return msg.trim();
}

// Extrai referência numérica do início da descrição (ex: "31- CONJUNTO SUÍÇA" → "31")
function extrairRef(descricao) {
  const match = (descricao || '').match(/^(\d+)/);
  return match ? match[1] : null;
}

// Extrai cor e tamanho da descrição Bling
function extrairCorTam(descricao) {
  let cor = '-', tam = '-';
  const matchCor = (descricao || '').match(/COR:\s*([^;]+)/i);
  if (matchCor) cor = matchCor[1].trim();
  const matchTam = (descricao || '').match(/TAM(?:ANHO)?:\s*([^;,\s]+)/i);
  if (matchTam) tam = matchTam[1].trim();
  return { cor, tam };
}

// Extrai nome limpo do produto (sem cor/tamanho)
function extrairNomeLimpo(descricao) {
  return (descricao || 'Sem descrição')
    .replace(/COR:\s*[^;]+/i, '')
    .replace(/TAM(?:ANHO)?:\s*[^;,\s]+/i, '')
    .replace(/;/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

// Busca saldos de um depósito específico para uma lista de produto IDs
async function buscarSaldosPorDeposito(token, produtoIds, depositoId) {
  const mapa = new Map();
  const batchSize = 100;
  for (let i = 0; i < produtoIds.length; i += batchSize) {
    const batch = produtoIds.slice(i, i + batchSize);
    const idsParam = batch.join('&idsProdutos[]=');
    try {
      const resp = await blingRequest(
        `https://api.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=${idsParam}`,
        token
      );
      const saldos = resp.data?.data || [];
      for (const s of saldos) {
        const idProduto = s.produto?.id;
        if (!idProduto) continue;
        let saldo = 0;
        if (s.depositos && Array.isArray(s.depositos)) {
          const dep = s.depositos.find(d => d.id === depositoId);
          if (dep) saldo = dep.saldoFisico ?? dep.saldoVirtual ?? 0;
        }
        mapa.set(idProduto, saldo);
      }
    } catch (e) {
      console.error(`   [Alerta] Erro ao buscar saldos (batch ${i}):`, e.message);
    }
    if (i + batchSize < produtoIds.length) await delay(500);
  }
  return mapa;
}

// Formata bloco de um conjunto (referência) com todas suas variações em alerta
function formatarBlocoConjunto(ref, nomeProduto, variacoes) {
  let bloco = `📦 *${ref} – ${nomeProduto}*\n`;
  for (const v of variacoes) {
    const icone = v.saldo === 0 ? '🚨' : '⚠️';
    const status = v.saldo === 0 ? 'ZERADO' : `${v.saldo} pçs`;
    bloco += `${icone} ${v.cor} | ${v.tam} → *${status}*\n`;
  }
  return bloco.trim();
}

// Gera alertas agrupados por referência (conjunto), separados por depósito
async function gerarAlertasPorConjunto(filtroRef, filtroDeposito) {
  const produtosMonitorados = cacheProdutos.filter(p => {
    const ref = extrairRef(p.descricao);
    if (!ref) return false;
    if (!(p.descricao || '').match(/COR:/i)) return false;
    if (filtroRef) return ref === filtroRef;
    return REFS_MONITORADAS.includes(ref);
  });

  if (produtosMonitorados.length === 0) {
    return { conjuntosSede: [], conjuntosBase: [] };
  }

  // Busca saldos BASE se necessário
  let saldosBase = new Map();
  if (!filtroDeposito || filtroDeposito === 'base') {
    try {
      const token = await obterAccessToken();
      saldosBase = await buscarSaldosPorDeposito(token, produtosMonitorados.map(p => p.id), DEPOSITO_BASE_ID);
    } catch (e) {
      console.error('   [Alerta] Erro ao buscar saldos BASE:', e.message);
    }
  }

  // Agrupa por referência
  const mapaSede = new Map(); // ref → { nome, variacoes[] }
  const mapaBase = new Map();

  for (const p of produtosMonitorados) {
    const ref = extrairRef(p.descricao);
    const { cor, tam } = extrairCorTam(p.descricao);
    const nomeLimpo = extrairNomeLimpo(p.descricao).replace(/^\d+[-\s]*/, '').trim();
    const saldoSede = p.saldoFisicoTotal || 0;
    const saldoBase = saldosBase.get(p.id) || 0;

    if ((!filtroDeposito || filtroDeposito === 'sede') && saldoSede < MINIMO_SEDE) {
      if (!mapaSede.has(ref)) mapaSede.set(ref, { nome: nomeLimpo, variacoes: [] });
      mapaSede.get(ref).variacoes.push({ cor, tam, saldo: saldoSede });
    }
    if ((!filtroDeposito || filtroDeposito === 'base') && saldoBase < MINIMO_BASE) {
      if (!mapaBase.has(ref)) mapaBase.set(ref, { nome: nomeLimpo, variacoes: [] });
      mapaBase.get(ref).variacoes.push({ cor, tam, saldo: saldoBase });
    }
  }

  // Converte para array de blocos formatados
  const conjuntosSede = [];
  for (const [ref, dados] of mapaSede) {
    conjuntosSede.push(formatarBlocoConjunto(ref, dados.nome, dados.variacoes));
  }

  const conjuntosBase = [];
  for (const [ref, dados] of mapaBase) {
    conjuntosBase.push(formatarBlocoConjunto(ref, dados.nome, dados.variacoes));
  }

  return { conjuntosSede, conjuntosBase };
}

// Envia conjuntos em lotes de 5 referências por mensagem, com pausa entre lotes
async function enviarPorLotes(chat, cabecalho, conjuntos) {
  const LOTE = 5;
  await chat.sendMessage(cabecalho);
  await delay(800);

  for (let i = 0; i < conjuntos.length; i += LOTE) {
    const lote = conjuntos.slice(i, i + LOTE);
    const msg = lote.join('\n\n━━━━━━━━━━━━━━━━━━\n\n');
    const pagina = Math.floor(i / LOTE) + 1;
    const totalPaginas = Math.ceil(conjuntos.length / LOTE);
    await chat.sendMessage(`📄 *${pagina}/${totalPaginas}*\n\n${msg}`);
    await delay(2000); // 2 segundos entre lotes
  }
}

// ══════════════════════════════════════════════════════════════
// 🚨 ALERTA MANAGER — Motor de alertas proativos pós-sync
// ══════════════════════════════════════════════════════════════

// Estado anti-spam: Map<"produtoId_deposito", nivelAlertado>
// Exemplo: "12345_sede" → 15 (já alertou quando caiu para <=15)
// Reseta quando estoque sobe acima do maior threshold (reposição)
const alertaEstado = new Map();

/**
 * Determina o nível de alerta para um saldo dado os thresholds.
 * Thresholds devem estar ordenados do mais crítico ao menos: [0, 15, 30]
 * Retorna o threshold atingido ou null se estoque está saudável.
 */
function determinarNivel(saldo, thresholds) {
  // Percorre do menos crítico ao mais crítico para achar o threshold mais alto atingido
  // depois refina para o mais crítico
  for (const t of thresholds) {
    if (saldo === 0 && t === 0) return 0;
    if (t > 0 && saldo > 0 && saldo <= t) return t;
  }
  if (saldo === 0) return 0;
  return null; // saudável
}

/**
 * Verifica se deve alertar para esta variação (anti-spam).
 * Retorna { deveAlertar: boolean, nivel: number|null }
 */
function verificarEstadoAlerta(produtoId, deposito, saldo, thresholds) {
  const chave = `${produtoId}_${deposito}`;
  const nivelAtual = determinarNivel(saldo, thresholds);
  const maiorThreshold = Math.max(...thresholds);

  // Estoque saudável → limpa estado (reposição detectada)
  if (nivelAtual === null) {
    if (alertaEstado.has(chave)) {
      alertaEstado.delete(chave);
      console.log(`   [Alerta] ✅ Reposição detectada: ${chave} (saldo: ${saldo})`);
    }
    return { deveAlertar: false, nivel: null };
  }

  // Estoque subiu mas ainda dentro de threshold → verifica se subiu de nível
  const nivelAnterior = alertaEstado.get(chave);

  if (nivelAnterior === undefined) {
    // Nunca alertou → alerta agora
    alertaEstado.set(chave, nivelAtual);
    return { deveAlertar: true, nivel: nivelAtual };
  }

  if (nivelAtual < nivelAnterior) {
    // Caiu para nível mais crítico (ex: de 30 → 15, ou de 15 → 0)
    alertaEstado.set(chave, nivelAtual);
    return { deveAlertar: true, nivel: nivelAtual };
  }

  // Estoque subiu mas voltou a cair para o mesmo nível → reposição parcial
  // Se subiu acima do nível anterior e voltou a cair, reseta e alerta
  // Se continua no mesmo nível → anti-spam, não alerta
  return { deveAlertar: false, nivel: nivelAtual };
}

/**
 * Monta um bloco completo (header + grade) de UMA referência em UM depósito.
 * Reutilizado tanto pelos alertas automáticos quanto pelos comandos manuais
 * !estoque sede <ref> e !estoque base <ref>.
 * @param {string} ref - Referência (ex: "108")
 * @param {string} deposito - 'sede' ou 'base'
 * @param {Map|null} saldosBase - Mapa de saldos BASE (somente se deposito === 'base')
 * @param {string} headerLabel - Texto do cabeçalho (ex: "🚨 ALERTA — 06/04 15:30")
 * @returns {string|null} Bloco formatado pronto para sendMessage, ou null se ref não existir
 */
function montarBlocoModelo(ref, deposito, saldosBase, headerLabel) {
  const variacoes = cacheProdutos.filter(p => {
    const r = extrairRef(p.descricao);
    return r === ref && /COR:/i.test(p.descricao || '');
  });

  if (variacoes.length === 0) return null;

  const nomeLimpo = extrairNomeLimpo(variacoes[0].descricao)
    .replace(/^\d+[-\s]*/, '').trim();

  const ordemTam = ['PP', 'P', 'M', 'G', 'GG', 'XG', 'XXG', 'EG', 'EGG'];
  const porTam = new Map();

  for (const p of variacoes) {
    const { cor, tam } = extrairCorTam(p.descricao);
    const saldo = deposito === 'base'
      ? (saldosBase ? (saldosBase.get(p.id) || 0) : 0)
      : (p.saldoFisicoTotal || 0);
    if (!porTam.has(tam)) porTam.set(tam, []);
    porTam.get(tam).push({ cor, saldo });
  }

  const tamanhos = [...porTam.keys()].sort((a, b) => {
    const ia = ordemTam.indexOf(a.toUpperCase());
    const ib = ordemTam.indexOf(b.toUpperCase());
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  const labelDep = deposito === 'base' ? 'BASE' : 'SEDE';
  let bloco = `${headerLabel} — *${labelDep}*\n`;
  bloco += `📦 *${ref} – ${nomeLimpo}*\n`;
  bloco += `━━━━━━━━━━━━━━━━━━`;

  for (const tam of tamanhos) {
    const cores = porTam.get(tam).sort((a, b) => a.cor.localeCompare(b.cor));
    bloco += `\n\n*${tam}*`;
    for (const { cor, saldo } of cores) {
      const icone = saldo === 0 ? '🚨' : saldo <= 5 ? '⚠️' : '✅';
      bloco += `\n${icone} ${cor}: ${saldo}`;
    }
  }
  return bloco;
}

/**
 * Motor principal: varre o cache aplicando thresholds e dispara alertas.
 * Chamado automaticamente após cada sincronização bem-sucedida.
 * NOVA REGRA: envia 1 mensagem por modelo (referência), não consolida modelos diferentes.
 */
async function avaliarEDispararAlertas() {
  if (!cacheProdutos || cacheProdutos.length === 0) {
    console.log('   [AlertaManager] Cache vazio, pulando.');
    return;
  }

  // Filtra variações monitoradas (Lista VIP + tem COR: na descrição)
  const monitorados = cacheProdutos.filter(p => {
    const ref = extrairRef(p.descricao);
    return ref && /COR:/i.test(p.descricao || '') && REFS_MONITORADAS.includes(ref);
  });

  if (monitorados.length === 0) {
    console.log('   [AlertaManager] Nenhum produto monitorado no cache.');
    return;
  }

  console.log(`   [AlertaManager] Analisando ${monitorados.length} variações da Lista VIP...`);

  // Busca saldos BASE (SEDE já está no cache via saldoFisicoTotal)
  let saldosBase = new Map();
  try {
    const token = await obterAccessToken();
    saldosBase = await buscarSaldosPorDeposito(token, monitorados.map(p => p.id), DEPOSITO_BASE_ID);
  } catch (e) {
    console.error('   [AlertaManager] Erro ao buscar saldos BASE:', e.message);
  }

  // Coleta refs em alerta por depósito (Map<ref, 'ruptura'|'critico'>)
  // Anti-spam continua per-variação, mas o agrupamento de envio é por ref.
  const refsAlertaSede = new Map();
  const refsAlertaBase = new Map();

  function classificarRef(mapa, ref, nivel) {
    const tipoAtual = mapa.get(ref);
    if (nivel === 0 || tipoAtual === 'ruptura') mapa.set(ref, 'ruptura');
    else mapa.set(ref, 'critico');
  }

  for (const p of monitorados) {
    const ref = extrairRef(p.descricao);
    const saldoSede = p.saldoFisicoTotal || 0;
    const saldoBase = saldosBase.get(p.id) || 0;

    const checkSede = verificarEstadoAlerta(p.id, 'sede', saldoSede, THRESHOLDS_SEDE);
    if (checkSede.deveAlertar) classificarRef(refsAlertaSede, ref, checkSede.nivel);

    const checkBase = verificarEstadoAlerta(p.id, 'base', saldoBase, THRESHOLDS_BASE);
    if (checkBase.deveAlertar) classificarRef(refsAlertaBase, ref, checkBase.nivel);
  }

  console.log(`   [AlertaManager] Resultado → SEDE: ${refsAlertaSede.size} modelo(s) | BASE: ${refsAlertaBase.size} modelo(s) | Estado: ${alertaEstado.size} variações rastreadas`);

  if (refsAlertaSede.size === 0 && refsAlertaBase.size === 0) {
    console.log('   [AlertaManager] ✅ Sem novos alertas (anti-spam ativo).');
    return;
  }

  // Busca os grupos de WhatsApp
  let chats;
  try {
    chats = await wppClient.getChats();
  } catch (e) {
    console.error('   [AlertaManager] WhatsApp não conectado:', e.message);
    return;
  }

  const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

  // --- SEDE: 1 mensagem por modelo ---
  if (refsAlertaSede.size > 0) {
    const grupoSede = chats.find(c => c.isGroup && c.name === GRUPO_ALERTA_SEDE);
    if (grupoSede) {
      for (const [ref, tipo] of refsAlertaSede) {
        const icone = tipo === 'ruptura' ? '🚨' : '🔴';
        const label = tipo === 'ruptura' ? 'RUPTURA' : 'CRÍTICO';
        const header = `${icone} *${label} ${dataHora}*`;
        const bloco = montarBlocoModelo(ref, 'sede', null, header);
        if (bloco) {
          await grupoSede.sendMessage(bloco);
          await delay(1500);
        }
      }
      console.log(`   [AlertaManager] 📤 SEDE: ${refsAlertaSede.size} modelo(s) enviado(s).`);
    } else {
      console.error(`   [AlertaManager] ⚠️ Grupo "${GRUPO_ALERTA_SEDE}" não encontrado!`);
    }
    await delay(2000);
  }

  // --- BASE: 1 mensagem por modelo ---
  if (refsAlertaBase.size > 0) {
    const grupoBase = chats.find(c => c.isGroup && c.name === GRUPO_ALERTA_BASE);
    if (grupoBase) {
      for (const [ref, tipo] of refsAlertaBase) {
        const icone = tipo === 'ruptura' ? '🚨' : '🔴';
        const label = tipo === 'ruptura' ? 'RUPTURA' : 'CRÍTICO';
        const header = `${icone} *${label} ${dataHora}*`;
        const bloco = montarBlocoModelo(ref, 'base', saldosBase, header);
        if (bloco) {
          await grupoBase.sendMessage(bloco);
          await delay(1500);
        }
      }
      console.log(`   [AlertaManager] 📤 BASE: ${refsAlertaBase.size} modelo(s) enviado(s).`);
    } else {
      console.error(`   [AlertaManager] ⚠️ Grupo "${GRUPO_ALERTA_BASE}" não encontrado!`);
    }
  }
}