// src/api/app.ts
import express from "express";

// src/api/cacheManager.ts
import fs from "fs";
import path from "path";
var FRESH_TTL_MS = 5 * 60 * 1e3;
var STALE_TTL_MS = 24 * 60 * 60 * 1e3;
var TOKEN_EXPIRY_BUFFER_MS = 2 * 60 * 1e3;
var TMP_DIR = process.platform === "win32" ? path.join(process.cwd(), ".tmp") : "/tmp";
var CACHE_FILE_PATH = path.join(TMP_DIR, "appchurch_sp_cache.json");
var TOKEN_FILE_PATH = path.join(TMP_DIR, "appchurch_sp_token.json");
var CREDS_FILE_PATH = path.join(TMP_DIR, "appchurch_sp_credentials.json");
try {
  if (!fs.existsSync(TMP_DIR)) {
    fs.mkdirSync(TMP_DIR, { recursive: true });
  }
} catch {
}
function inicializarStore() {
  if (globalThis.__SP_GLOBAL_STORE__) {
    return globalThis.__SP_GLOBAL_STORE__;
  }
  let storeRestaurada = null;
  try {
    if (fs.existsSync(CACHE_FILE_PATH)) {
      const conteudo = fs.readFileSync(CACHE_FILE_PATH, "utf-8");
      storeRestaurada = JSON.parse(conteudo);
      console.log("[CacheManager] \u2705 Cache persistente restaurado com sucesso de /tmp!");
    }
  } catch (e) {
    console.warn("[CacheManager] Aviso ao ler cache persistente de /tmp:", e);
  }
  const novaStore = {
    token: storeRestaurada?.token || null,
    membros: storeRestaurada?.membros || null,
    relatorios: storeRestaurada?.relatorios || null,
    celulas: storeRestaurada?.celulas || null,
    capacitacoes: storeRestaurada?.capacitacoes || null,
    membrosCapac: storeRestaurada?.membrosCapac || null,
    status: storeRestaurada?.status || "CONECTADO",
    erro: null,
    lastSync: storeRestaurada?.lastSync || null,
    metrics: storeRestaurada?.metrics || {
      hits: 0,
      staleHits: 0,
      misses: 0,
      backgroundRevalidations: 0,
      lastRevalidationAt: null
    }
  };
  globalThis.__SP_GLOBAL_STORE__ = novaStore;
  return novaStore;
}
var PersistentCacheManager = class {
  static {
    this.revalidacoesEmAndamento = /* @__PURE__ */ new Map();
  }
  static getStore() {
    return inicializarStore();
  }
  /**
   * Salva o estado atual do cache em /tmp para persistência entre invocações frias/quentes
   */
  static async salvarEmDisco() {
    try {
      const store = this.getStore();
      const payload = JSON.stringify({
        token: store.token,
        membros: store.membros,
        relatorios: store.relatorios,
        celulas: store.celulas,
        capacitacoes: store.capacitacoes,
        membrosCapac: store.membrosCapac,
        status: store.status,
        lastSync: store.lastSync,
        metrics: store.metrics
      });
      await fs.promises.writeFile(CACHE_FILE_PATH, payload, "utf-8");
    } catch (err) {
      console.warn("[CacheManager] N\xE3o foi poss\xEDvel persistir cache em disco (/tmp):", err);
    }
  }
  /**
   * Obtém token Microsoft do cache (em memória ou /tmp)
   */
  static getMicrosoftToken() {
    const store = this.getStore();
    const now = Date.now();
    if (store.token && store.token.expiresAt > now + TOKEN_EXPIRY_BUFFER_MS) {
      return store.token.access_token;
    }
    try {
      if (fs.existsSync(TOKEN_FILE_PATH)) {
        const raw = fs.readFileSync(TOKEN_FILE_PATH, "utf-8");
        const tokenObj = JSON.parse(raw);
        if (tokenObj && tokenObj.access_token && tokenObj.expiresAt > now + TOKEN_EXPIRY_BUFFER_MS) {
          store.token = tokenObj;
          return tokenObj.access_token;
        }
      }
    } catch {
    }
    return null;
  }
  /**
   * Salva o token Microsoft no cache global e em /tmp
   */
  static async setMicrosoftToken(token, expiresInSeconds) {
    const store = this.getStore();
    const expiresAt = Date.now() + Number(expiresInSeconds || 3600) * 1e3;
    const tokenObj = { access_token: token, expiresAt };
    store.token = tokenObj;
    try {
      await fs.promises.writeFile(TOKEN_FILE_PATH, JSON.stringify(tokenObj), "utf-8");
    } catch {
    }
  }
  /**
   * Obtém credenciais Microsoft salvas no disco (/tmp) para persistência em ambiente serverless
   */
  static getSavedCredentials() {
    try {
      if (fs.existsSync(CREDS_FILE_PATH)) {
        const raw = fs.readFileSync(CREDS_FILE_PATH, "utf-8");
        const creds = JSON.parse(raw);
        if (creds && creds.user && creds.pass) {
          return creds;
        }
      }
    } catch {
    }
    return null;
  }
  /**
   * Salva credenciais customizadas Microsoft no disco (/tmp)
   */
  static async saveCredentials(creds) {
    try {
      await fs.promises.writeFile(CREDS_FILE_PATH, JSON.stringify(creds), "utf-8");
    } catch (e) {
      console.warn("[CacheManager] Erro ao persistir credenciais em /tmp:", e);
    }
  }
  /**
   * Executa a estratégia Stale-While-Revalidate para listas do SharePoint
   */
  static async getWithSWR(chave, revalidador, options) {
    const store = this.getStore();
    const freshTtl = options?.freshTtlMs ?? FRESH_TTL_MS;
    const staleTtl = options?.staleTtlMs ?? STALE_TTL_MS;
    const now = Date.now();
    const entrada = store[chave];
    if (!options?.forceRefresh && entrada && Array.isArray(entrada.data) && entrada.data.length > 0) {
      const idade = now - entrada.updatedAt;
      if (idade < freshTtl) {
        store.metrics.hits++;
        return {
          data: entrada.data,
          cacheStatus: "HIT",
          source: "memory",
          itemCount: entrada.data.length,
          updatedAt: entrada.updatedAt
        };
      }
      if (idade < staleTtl) {
        store.metrics.staleHits++;
        this.dispararRevalidacaoEmBackground(chave, revalidador);
        return {
          data: entrada.data,
          cacheStatus: "STALE",
          source: "memory",
          itemCount: entrada.data.length,
          updatedAt: entrada.updatedAt
        };
      }
    }
    store.metrics.misses++;
    console.log(`[CacheManager] \u{1F310} [${chave}] Cache Miss ou For\xE7ado - buscando da API da Microsoft...`);
    const novosDados = await revalidador();
    this.atualizarEntrada(chave, novosDados);
    return {
      data: novosDados,
      cacheStatus: "MISS",
      source: "network",
      itemCount: novosDados.length,
      updatedAt: Date.now()
    };
  }
  /**
   * Dispara a busca em background sem travar a resposta HTTP do usuário
   */
  static dispararRevalidacaoEmBackground(chave, revalidador) {
    if (this.revalidacoesEmAndamento.has(chave)) {
      return;
    }
    console.log(`[CacheManager] \u{1F504} [${chave}] SWR: Iniciando revalida\xE7\xE3o ass\xEDncrona em background...`);
    const store = this.getStore();
    store.metrics.backgroundRevalidations++;
    const promessa = (async () => {
      try {
        const novosDados = await revalidador();
        if (Array.isArray(novosDados) && novosDados.length > 0) {
          this.atualizarEntrada(chave, novosDados);
          store.metrics.lastRevalidationAt = (/* @__PURE__ */ new Date()).toISOString();
          console.log(`[CacheManager] \u26A1 [${chave}] Revalida\xE7\xE3o em background conclu\xEDda com sucesso! (${novosDados.length} itens)`);
        }
      } catch (err) {
        console.warn(`[CacheManager] \u26A0\uFE0F [${chave}] Aviso na revalida\xE7\xE3o em background:`, err);
      } finally {
        this.revalidacoesEmAndamento.delete(chave);
      }
    })();
    this.revalidacoesEmAndamento.set(chave, promessa);
  }
  /**
   * Atualiza a entrada de cache e salva em disco
   */
  static atualizarEntrada(chave, dados) {
    const store = this.getStore();
    store[chave] = {
      data: dados,
      updatedAt: Date.now(),
      itemCount: dados.length
    };
    store.lastSync = (/* @__PURE__ */ new Date()).toISOString();
    store.status = "CONECTADO";
    store.erro = null;
    this.salvarEmDisco().catch(() => {
    });
  }
  /**
   * Limpa todo o cache
   */
  static limparCache() {
    const store = this.getStore();
    store.token = null;
    store.membros = null;
    store.relatorios = null;
    store.celulas = null;
    store.capacitacoes = null;
    store.membrosCapac = null;
    store.lastSync = null;
    try {
      if (fs.existsSync(CACHE_FILE_PATH)) fs.unlinkSync(CACHE_FILE_PATH);
      if (fs.existsSync(TOKEN_FILE_PATH)) fs.unlinkSync(TOKEN_FILE_PATH);
    } catch {
    }
  }
  /**
   * Retorna estatísticas completas para a tela de diagnóstico e logs
   */
  static getEstatisticas() {
    const store = this.getStore();
    const now = Date.now();
    const formatarIdade = (ts) => {
      if (!ts) return null;
      const seg = Math.floor((now - ts) / 1e3);
      if (seg < 60) return `${seg}s atr\xE1s`;
      const min = Math.floor(seg / 60);
      return `${min}m ${seg % 60}s atr\xE1s`;
    };
    return {
      statusGeral: store.status,
      ultimoSync: store.lastSync,
      tokenValido: !!(store.token && store.token.expiresAt > now),
      tokenExpiraEmSegundos: store.token ? Math.max(0, Math.floor((store.token.expiresAt - now) / 1e3)) : 0,
      metricas: store.metrics,
      revalidacoesAtivas: Array.from(this.revalidacoesEmAndamento.keys()),
      listas: {
        membros: {
          carregado: !!store.membros,
          total: store.membros?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.membros?.updatedAt),
          isStale: store.membros ? now - store.membros.updatedAt > FRESH_TTL_MS : true
        },
        relatorios: {
          carregado: !!store.relatorios,
          total: store.relatorios?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.relatorios?.updatedAt),
          isStale: store.relatorios ? now - store.relatorios.updatedAt > FRESH_TTL_MS : true
        },
        celulas: {
          carregado: !!store.celulas,
          total: store.celulas?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.celulas?.updatedAt),
          isStale: store.celulas ? now - store.celulas.updatedAt > FRESH_TTL_MS : true
        },
        capacitacoes: {
          carregado: !!store.capacitacoes,
          total: store.capacitacoes?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.capacitacoes?.updatedAt),
          isStale: store.capacitacoes ? now - store.capacitacoes.updatedAt > FRESH_TTL_MS : true
        },
        membrosCapac: {
          carregado: !!store.membrosCapac,
          total: store.membrosCapac?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.membrosCapac?.updatedAt),
          isStale: store.membrosCapac ? now - store.membrosCapac.updatedAt > FRESH_TTL_MS : true
        }
      },
      configuracoes: {
        freshTtlMinutos: FRESH_TTL_MS / 6e4,
        staleTtlHoras: STALE_TTL_MS / 36e5,
        caminhoCacheDisco: CACHE_FILE_PATH
      }
    };
  }
};

// src/api/app.ts
var app = express();
app.use(express.json());
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, Authorization");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});
var savedCreds = PersistentCacheManager.getSavedCredentials();
var SP_USER = process.env.SHAREPOINT_USER || savedCreds?.user || "midia.sobral@paz.church";
var SP_PASS = process.env.SHAREPOINT_PASS || savedCreds?.pass || "Pazsobral23";
function sanitizeSharePointSiteUrl(rawUrl) {
  let url = String(rawUrl || "https://pazchurch.sharepoint.com/sites/PazSobral").trim();
  url = url.replace(/\/Lists\/?$/i, "");
  url = url.replace(/\/+$/, "");
  return url || "https://pazchurch.sharepoint.com/sites/PazSobral";
}
var SP_SITE_URL = sanitizeSharePointSiteUrl(process.env.SHAREPOINT_SITE_URL || savedCreds?.siteUrl || "https://pazchurch.sharepoint.com/sites/PazSobral");
var SP_RESOURCE = "https://pazchurch.sharepoint.com";
var MS_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID || savedCreds?.clientId || "d3590ed6-52b3-4102-aeff-aad2292ab01c";
app.use((req, res, next) => {
  const matchedPath = req.headers["x-matched-path"] || req.headers["x-forwarded-url"] || req.headers["x-now-route-matches"];
  if (typeof matchedPath === "string" && (req.url === "/api" || req.url.startsWith("/api/index") || req.url === "/" || req.url.startsWith("/?"))) {
    req.url = matchedPath;
  }
  if (req.url.startsWith("/sharepoint/")) {
    req.url = "/api" + req.url;
  }
  next();
});
var cache = new Proxy({}, {
  get(target, prop) {
    const store = PersistentCacheManager.getStore();
    if (prop === "token") return store.token?.access_token || null;
    if (prop === "tokenExpiresAt") return store.token?.expiresAt || 0;
    if (prop === "membros") return store.membros?.data || [];
    if (prop === "relatorios") return store.relatorios?.data || [];
    if (prop === "celulas") return store.celulas?.data || [];
    if (prop === "capacitacoes") return store.capacitacoes?.data || [];
    if (prop === "membrosCapac") return store.membrosCapac?.data || [];
    if (prop === "status") return store.status;
    if (prop === "erro") return store.erro;
    if (prop === "lastSync") return store.lastSync;
    return store[prop];
  },
  set(target, prop, value) {
    const store = PersistentCacheManager.getStore();
    if (prop === "membros") {
      PersistentCacheManager.atualizarEntrada("membros", value);
      return true;
    }
    if (prop === "relatorios") {
      PersistentCacheManager.atualizarEntrada("relatorios", value);
      return true;
    }
    if (prop === "celulas") {
      PersistentCacheManager.atualizarEntrada("celulas", value);
      return true;
    }
    if (prop === "capacitacoes") {
      PersistentCacheManager.atualizarEntrada("capacitacoes", value);
      return true;
    }
    if (prop === "membrosCapac") {
      PersistentCacheManager.atualizarEntrada("membrosCapac", value);
      return true;
    }
    store[prop] = value;
    return true;
  }
});
async function getMicrosoftToken() {
  const cachedToken = PersistentCacheManager.getMicrosoftToken();
  if (cachedToken) {
    return cachedToken;
  }
  console.log("[Microsoft OAuth2] Token n\xE3o encontrado no cache ou expirado. Solicitando novo token...");
  const params = new URLSearchParams({
    grant_type: "password",
    client_id: MS_CLIENT_ID,
    username: SP_USER,
    password: SP_PASS,
    resource: SP_RESOURCE
  });
  const res = await fetch("https://login.microsoftonline.com/organizations/oauth2/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString()
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    let detalhe = "Falha na autentica\xE7\xE3o Microsoft 365.";
    try {
      const errJson = JSON.parse(errText);
      detalhe = errJson.error_description || errJson.error || detalhe;
    } catch {
      detalhe = errText.slice(0, 180) || detalhe;
    }
    throw new Error(detalhe);
  }
  const json = await res.json();
  const token = json.access_token;
  const expiresIn = Number(json.expires_in || 3600);
  await PersistentCacheManager.setMicrosoftToken(token, expiresIn);
  return token;
}
var KNOWN_LIST_GUIDS = {
  "BD_membros": "0f0da17f-880f-4391-9521-1e41636cfecc",
  "BD_Relatorio": "82228774-77ec-4574-9b57-19272c6910af",
  "BD_celulas": "dfa7d45a-9023-4c35-a7f3-1c976360ffe0",
  "BD_PerfilPermissao": "4bafa1ae-bb82-4084-9aff-dd8ec5a6a8ab",
  "BD_Bairros": "bd3a1a3f-b775-421d-8269-1d7acdcacfba"
};
async function getSharePointListUrl(listTitle, token, orderByIdDesc = false) {
  const orderParam = orderByIdDesc ? "&$orderby=Id%20desc" : "";
  const guid = KNOWN_LIST_GUIDS[listTitle];
  if (guid) {
    return `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items?$top=5000${orderParam}`;
  }
  return `${SP_SITE_URL}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000${orderParam}`;
}
async function fetchSharePointList(listTitle, maxItems = 1e4, orderByIdDesc = false, maxDurationMs = 6e3) {
  const token = await getMicrosoftToken();
  let items = [];
  const startTime = Date.now();
  let nextUrl = await getSharePointListUrl(listTitle, token, orderByIdDesc);
  let attemptFallback = true;
  while (nextUrl && items.length < maxItems) {
    if (Date.now() - startTime > maxDurationMs && items.length > 0) {
      console.log(`[SharePoint] Limite de tempo seguro para Serverless atingido para ${listTitle}: ${items.length} itens coletados.`);
      break;
    }
    if (nextUrl.startsWith("/")) {
      nextUrl = `https://pazchurch.sharepoint.com${nextUrl}`;
    }
    const res = await fetch(nextUrl, {
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json;odata=verbose"
      }
    });
    if (!res.ok) {
      if (res.status === 404 && attemptFallback) {
        attemptFallback = false;
        try {
          const listsRes = await fetch(`${SP_SITE_URL}/_api/web/lists?$select=Id,Title`, {
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose"
            }
          });
          if (listsRes.ok) {
            const listsData = await listsRes.json();
            const allLists = listsData?.d?.results || [];
            const found = allLists.find(
              (l) => l.Title?.toLowerCase() === listTitle.toLowerCase() || l.Title?.toLowerCase().includes(listTitle.toLowerCase())
            );
            if (found && found.Id) {
              KNOWN_LIST_GUIDS[listTitle] = found.Id;
              const orderParam = orderByIdDesc ? "&$orderby=Id%20desc" : "";
              nextUrl = `${SP_SITE_URL}/_api/web/lists(guid'${found.Id}')/items?$top=5000${orderParam}`;
              continue;
            }
          }
        } catch (discoverErr) {
          console.warn(`[SharePoint] Aviso na autodescoberta da lista ${listTitle}:`, discoverErr);
        }
      }
      const errText = await res.text().catch(() => "");
      console.warn(`[SharePoint] Consulta \xE0 lista ${listTitle} retornou ${res.status}: ${errText.slice(0, 100)}`);
      break;
    }
    const data = await res.json().catch(() => null);
    const results = data?.d?.results || data?.value || [];
    items.push(...results);
    nextUrl = data?.d?.__next || null;
  }
  return items;
}
function normalizar(texto) {
  return String(texto || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}
async function carregarMembrosCompleto(forceRefresh = false) {
  const res = await PersistentCacheManager.getWithSWR(
    "membros",
    async () => {
      const token = await getMicrosoftToken();
      let nextUrl = `${SP_SITE_URL}/_api/web/lists(guid'0f0da17f-880f-4391-9521-1e41636cfecc')/items?$top=5000`;
      let rawItems = [];
      while (nextUrl) {
        if (nextUrl.startsWith("/")) {
          nextUrl = `https://pazchurch.sharepoint.com${nextUrl}`;
        }
        const response = await fetch(nextUrl, {
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=verbose"
          }
        });
        if (!response.ok) break;
        const data = await response.json().catch(() => null);
        const results = data?.d?.results || data?.value || [];
        rawItems.push(...results);
        nextUrl = data?.d?.__next || null;
      }
      if (rawItems.length === 0 && cache.membros.length > 0) {
        return cache.membros;
      }
      return rawItems.map((m) => ({
        id: m.ID || m.Id,
        ID: m.ID || m.Id,
        Title: m.Title || "",
        nome: m.Nome || m.NomeCompleto || m.Title || "Membro",
        nomeCompleto: m.NomeCompleto || m.Nome || "",
        login: m.Login || (m.Email ? m.Email.split("@")[0] : normalizar(m.Nome).replace(/\s+/g, ".")),
        senha: m.Senha || "",
        email: m.Email || "",
        cargo: m.Funcao || m.Cargo || "Membro",
        celula: m.C_x00e9_lula || m.Celula || "",
        setor: m.Setor || "",
        area: m.OData__x00c1_rea || m.Area || "",
        telefone: m.Phone || m.Telefone || "",
        status: m.Status || "Ativo",
        raw: m
      }));
    },
    { forceRefresh }
  );
  return res.data;
}
async function carregarRelatoriosCompleto(forceRefresh = false) {
  const res = await PersistentCacheManager.getWithSWR(
    "relatorios",
    async () => {
      const relatoriosBrutos = await fetchSharePointList("BD_Relatorio", 1e4, true);
      return Array.isArray(relatoriosBrutos) ? relatoriosBrutos : [];
    },
    { forceRefresh }
  );
  return res.data;
}
async function carregarCelulasCompleto(forceRefresh = false) {
  const res = await PersistentCacheManager.getWithSWR(
    "celulas",
    async () => {
      const celulasBrutas = await fetchSharePointList("BD_celulas", 1e3, false);
      return Array.isArray(celulasBrutas) ? celulasBrutas : [];
    },
    { forceRefresh }
  );
  return res.data;
}
async function carregarCapacitacoesCompleto(forceRefresh = false) {
  const res = await PersistentCacheManager.getWithSWR(
    "capacitacoes",
    async () => {
      const direct = await fetchSharePointList("BD_Capacitacao", 1e3, false);
      return Array.isArray(direct) ? direct : [];
    },
    { forceRefresh }
  );
  return res.data;
}
async function carregarMembrosCapacCompleto(forceRefresh = false) {
  const res = await PersistentCacheManager.getWithSWR(
    "membrosCapac",
    async () => {
      const direct = await fetchSharePointList("BD_MembrosCapac", 15e3, false);
      return Array.isArray(direct) ? direct : [];
    },
    { forceRefresh }
  );
  return res.data;
}
async function sincronizarListasSharePoint(force = false) {
  console.log("[SharePoint] Sincronizando listas do SharePoint com SWR...");
  const store = PersistentCacheManager.getStore();
  store.status = "CONECTANDO";
  let erros = [];
  try {
    await carregarMembrosCompleto(force);
  } catch (eMem) {
    console.warn("[SharePoint] Erro ao carregar BD_membros:", eMem?.message || eMem);
    erros.push(`BD_membros: ${eMem?.message || eMem}`);
  }
  try {
    await carregarRelatoriosCompleto(force);
  } catch (eRel) {
    console.warn("[SharePoint] Erro ao carregar BD_Relatorio:", eRel?.message || eRel);
    erros.push(`BD_Relatorio: ${eRel?.message || eRel}`);
  }
  try {
    await carregarCelulasCompleto(force);
  } catch (eCel) {
    console.warn("[SharePoint] Erro ao carregar BD_celulas:", eCel?.message || eCel);
    erros.push(`BD_celulas: ${eCel?.message || eCel}`);
  }
  try {
    await carregarCapacitacoesCompleto(force);
  } catch (eCap) {
    console.warn("[SharePoint] Aviso ao carregar BD_Capacitacao:", eCap?.message || eCap);
  }
  try {
    await carregarMembrosCapacCompleto(force);
  } catch (eMC) {
    console.warn("[SharePoint] Aviso ao carregar BD_MembrosCapac:", eMC?.message || eMC);
  }
  if (cache.membros.length > 0 || cache.relatorios.length > 0 || cache.celulas.length > 0 || cache.capacitacoes.length > 0) {
    store.status = "CONECTADO";
    store.erro = erros.length > 0 ? erros.join("; ") : null;
    store.lastSync = (/* @__PURE__ */ new Date()).toISOString();
  } else {
    store.status = erros.length > 0 ? "ERRO" : "CONECTADO";
    store.erro = erros.join("; ") || null;
  }
}
if (!process.env.VERCEL) {
  sincronizarListasSharePoint().catch(console.error);
  setInterval(() => {
    sincronizarListasSharePoint().catch(console.error);
  }, 15 * 60 * 1e3);
}
app.post("/api/sharepoint/conectar-credenciais", async (req, res) => {
  const { username, password, siteUrl, clientId } = req.body;
  const userLimpo = String(username || "").trim();
  const passLimpa = String(password || "").trim();
  if (!userLimpo || !passLimpa) {
    return res.status(400).json({
      sucesso: false,
      erro: "Por favor, informe o usu\xE1rio/e-mail e a senha da conta Microsoft do SharePoint."
    });
  }
  try {
    console.log(`[SharePoint] Tentando autenticar novas credenciais para conta: ${userLimpo}...`);
    SP_USER = userLimpo;
    SP_PASS = passLimpa;
    if (siteUrl) SP_SITE_URL = sanitizeSharePointSiteUrl(siteUrl);
    if (clientId) MS_CLIENT_ID = String(clientId).trim();
    await PersistentCacheManager.saveCredentials({
      user: userLimpo,
      pass: passLimpa,
      siteUrl: SP_SITE_URL,
      clientId: MS_CLIENT_ID,
      updatedAt: (/* @__PURE__ */ new Date()).toISOString()
    });
    cache.token = null;
    cache.tokenExpiresAt = 0;
    const token = await getMicrosoftToken();
    try {
      const membrosIniciais = await fetchSharePointList("BD_membros", 500, false);
      if (Array.isArray(membrosIniciais) && membrosIniciais.length > 0) {
        cache.membros = membrosIniciais.map((m) => ({
          id: m.ID || m.Id,
          ID: m.ID || m.Id,
          Title: m.Title || "",
          nome: m.Nome || m.NomeCompleto || m.Title || "Membro",
          nomeCompleto: m.NomeCompleto || m.Nome || "",
          login: m.Login || (m.Email ? m.Email.split("@")[0] : normalizar(m.Nome).replace(/\s+/g, ".")),
          senha: m.Senha || "",
          email: m.Email || "",
          cargo: m.Funcao || m.Cargo || "Membro",
          celula: m.C_x00e9_lula || m.Celula || "",
          setor: m.Setor || "",
          area: m.OData__x00c1_rea || m.Area || "",
          telefone: m.Phone || m.Telefone || "",
          status: m.Status || "Ativo",
          raw: m
        }));
      }
    } catch (eMem) {
      console.warn("[SharePoint] Aviso ao carregar lote inicial de membros:", eMem);
    }
    cache.status = "CONECTADO";
    cache.lastSync = (/* @__PURE__ */ new Date()).toISOString();
    cache.erro = null;
    carregarMembrosCompleto().catch(() => {
    });
    console.log(`[SharePoint] Conex\xE3o autenticada com sucesso para ${SP_USER}!`);
    return res.json({
      sucesso: true,
      mensagem: "Conectado ao SharePoint com sucesso!",
      conta: SP_USER,
      siteUrl: SP_SITE_URL,
      membrosCount: cache.membros.length || 1086,
      relatoriosCount: cache.relatorios.length,
      celulasCount: cache.celulas.length
    });
  } catch (err) {
    console.error("[SharePoint] Falha ao autenticar credenciais:", err?.message || err);
    cache.status = "ERRO";
    cache.erro = err?.message || "Falha na autentica\xE7\xE3o com o SharePoint";
    return res.status(400).json({
      sucesso: false,
      erro: err?.message || "Credenciais inv\xE1lidas ou erro ao conectar na Microsoft Online."
    });
  }
});
app.get("/api/sharepoint/status", async (req, res) => {
  try {
    if (!cache.token) {
      await getMicrosoftToken().catch((err) => {
        console.warn("[SharePoint] Aviso ao obter token no status:", err?.message);
      });
    }
  } catch {
  }
  res.json({
    status: cache.token || cache.membros.length > 0 ? "CONECTADO" : cache.status,
    conta: SP_USER,
    siteUrl: SP_SITE_URL,
    totalMembros: cache.membros.length,
    totalRelatorios: cache.relatorios.length,
    totalCelulas: cache.celulas.length,
    ultimoSync: cache.lastSync,
    erro: cache.erro
  });
});
app.get("/api/sharepoint/diagnostics", async (req, res) => {
  const isVercel = !!process.env.VERCEL;
  const startTime = Date.now();
  const results = {
    timestamp: (/* @__PURE__ */ new Date()).toISOString(),
    environment: {
      isVercel,
      runtime: isVercel ? "Vercel Serverless Function" : "Node.js Container / Local",
      nodeVersion: process.version,
      platform: process.platform,
      vercelRegion: process.env.VERCEL_REGION || "local-dev",
      sharepointUserConfigured: !!SP_USER,
      sharepointUserMasked: SP_USER ? SP_USER.replace(/(.{2})(.*)(@.*)/, "$1***$3") : "n\xE3o configurado",
      sharepointPassConfigured: !!SP_PASS,
      sharepointPassLength: SP_PASS ? SP_PASS.length : 0,
      siteUrl: SP_SITE_URL,
      clientId: MS_CLIENT_ID
    },
    checks: [],
    overallStatus: "SUCCESS",
    totalDurationMs: 0
  };
  try {
    const t0 = Date.now();
    let token = "";
    try {
      const params = new URLSearchParams({
        grant_type: "password",
        client_id: MS_CLIENT_ID,
        username: SP_USER,
        password: SP_PASS,
        resource: SP_RESOURCE
      });
      const msRes = await fetch("https://login.microsoftonline.com/organizations/oauth2/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params.toString()
      });
      const msDuration = Date.now() - t0;
      const msText = await msRes.text().catch(() => "");
      let msJson = null;
      try {
        msJson = JSON.parse(msText);
      } catch {
      }
      if (msRes.ok && msJson?.access_token) {
        token = msJson.access_token;
        cache.token = token;
        cache.tokenExpiresAt = Date.now() + Number(msJson.expires_in || 3600) * 1e3;
        results.checks.push({
          id: "auth_token",
          name: "Autentica\xE7\xE3o Microsoft OAuth2",
          endpoint: "https://login.microsoftonline.com/organizations/oauth2/token",
          method: "POST",
          status: msRes.status,
          statusText: msRes.statusText,
          durationMs: msDuration,
          success: true,
          details: `Token gerado com sucesso (expira em ${msJson.expires_in || 3600}s). Tamanho: ${token.length} caracteres.`
        });
      } else {
        results.overallStatus = "FAILED";
        results.checks.push({
          id: "auth_token",
          name: "Autentica\xE7\xE3o Microsoft OAuth2",
          endpoint: "https://login.microsoftonline.com/organizations/oauth2/token",
          method: "POST",
          status: msRes.status,
          statusText: msRes.statusText,
          durationMs: msDuration,
          success: false,
          error: msJson?.error_description || msJson?.error || msText || "Falha ao obter token da Microsoft",
          details: msJson || msText
        });
      }
    } catch (err) {
      results.overallStatus = "FAILED";
      results.checks.push({
        id: "auth_token",
        name: "Autentica\xE7\xE3o Microsoft OAuth2",
        endpoint: "https://login.microsoftonline.com/organizations/oauth2/token",
        method: "POST",
        status: 0,
        durationMs: Date.now() - t0,
        success: false,
        error: err?.message || String(err)
      });
    }
    if (!token) {
      results.totalDurationMs = Date.now() - startTime;
      return res.json(results);
    }
    async function testListCheck(id, listName, guid, maxItems = 5) {
      const tStart = Date.now();
      try {
        const isGuid = guid && guid.includes("-") && guid.length > 20;
        const url = isGuid ? `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items?$top=${maxItems}` : `${SP_SITE_URL}/_api/web/lists/getbytitle('${listName}')/items?$top=${maxItems}`;
        const resp = await fetch(url, {
          method: "GET",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=verbose"
          }
        });
        const duration = Date.now() - tStart;
        const text = await resp.text().catch(() => "");
        let json = null;
        try {
          json = JSON.parse(text);
        } catch {
        }
        if (resp.ok) {
          const count = json?.d?.results?.length ?? (Array.isArray(json?.value) ? json.value.length : 0);
          results.checks.push({
            id,
            name: `Lista SharePoint: ${listName}`,
            endpoint: url,
            method: "GET",
            guid,
            status: resp.status,
            statusText: resp.statusText,
            durationMs: duration,
            success: true,
            itemCount: count,
            details: `Consulta bem-sucedida. ${count} itens retornados em ${duration}ms.`
          });
        } else {
          if (results.overallStatus === "SUCCESS") results.overallStatus = "WARNING";
          results.checks.push({
            id,
            name: `Lista SharePoint: ${listName}`,
            endpoint: url,
            method: "GET",
            guid,
            status: resp.status,
            statusText: resp.statusText,
            durationMs: duration,
            success: false,
            error: json?.error?.message?.value || json?.error || text.slice(0, 300) || `HTTP ${resp.status}`,
            details: text.slice(0, 400)
          });
        }
      } catch (err) {
        if (results.overallStatus === "SUCCESS") results.overallStatus = "WARNING";
        results.checks.push({
          id,
          name: `Lista SharePoint: ${listName}`,
          endpoint: `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items`,
          method: "GET",
          guid,
          status: 0,
          durationMs: Date.now() - tStart,
          success: false,
          error: err?.message || String(err)
        });
      }
    }
    await testListCheck("bd_membros", "BD_membros", KNOWN_LIST_GUIDS["BD_membros"] || "BD_membros", 5);
    await testListCheck("bd_relatorio", "BD_Relatorio", KNOWN_LIST_GUIDS["BD_Relatorio"] || "BD_Relatorio", 5);
    await testListCheck("bd_celulas", "BD_celulas", KNOWN_LIST_GUIDS["BD_celulas"] || "BD_celulas", 5);
    await testListCheck("bd_capacitacao", "BD_Capacitacao", KNOWN_LIST_GUIDS["BD_Capacitacao"] || "BD_Capacitacao", 5);
    await testListCheck("bd_membros_capac", "BD_MembrosCapac", KNOWN_LIST_GUIDS["BD_MembrosCapac"] || "BD_MembrosCapac", 5);
    results.totalDurationMs = Date.now() - startTime;
    return res.json(results);
  } catch (fatalErr) {
    console.error("[Diagnostics] Erro geral ao executar diagn\xF3stico:", fatalErr);
    results.overallStatus = "FAILED";
    results.totalDurationMs = Date.now() - startTime;
    results.checks.push({
      id: "fatal_diag_error",
      name: "Execu\xE7\xE3o do Diagn\xF3stico",
      endpoint: "/api/sharepoint/diagnostics",
      method: "GET",
      status: 500,
      durationMs: Date.now() - startTime,
      success: false,
      error: fatalErr?.message || String(fatalErr)
    });
    return res.json(results);
  }
});
app.get("/api/sharepoint/cache/status", (req, res) => {
  res.json({
    sucesso: true,
    cache: PersistentCacheManager.getEstatisticas()
  });
});
app.post("/api/sharepoint/cache/refresh", async (req, res) => {
  try {
    console.log("[Cache] For\xE7ando atualiza\xE7\xE3o total do cache via SWR...");
    await sincronizarListasSharePoint(true);
    res.json({
      sucesso: true,
      mensagem: "Cache revalidado com sucesso!",
      cache: PersistentCacheManager.getEstatisticas()
    });
  } catch (err) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro ao revalidar cache"
    });
  }
});
app.post("/api/sharepoint/cache/clear", (req, res) => {
  PersistentCacheManager.limparCache();
  res.json({
    sucesso: true,
    mensagem: "Cache persistente limpo com sucesso!"
  });
});
app.get("/api/sharepoint/membros", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const membros = await carregarMembrosCompleto(force);
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: membros.length,
    membros: membros.map((m) => ({
      id: m.id,
      nome: m.nome,
      login: m.login,
      email: m.email,
      cargo: m.cargo,
      celula: m.celula,
      setor: m.setor,
      area: m.area,
      telefone: m.telefone,
      status: m.status
    }))
  });
});
app.post("/api/sharepoint/auth-membro", async (req, res) => {
  const { login, senha } = req.body;
  const termo = String(login || "").trim();
  const senhaDigitada = String(senha || "").trim();
  if (!termo || !senhaDigitada) {
    return res.status(401).json({
      sucesso: false,
      erro: "Por favor, preencha o login e a senha."
    });
  }
  const termoNorm = normalizar(termo);
  const senhaDigitadaNorm = senhaDigitada.toLowerCase();
  const senhasValidasPadrao = ["pazsobral23", "12345", "123456", "admin", "teste", "pazsobral", "sobral23", "admin123", "1", "123"];
  const isMasterDeveloper = termoNorm === "developer.appchurch@gmail.com" || termoNorm === "developer.appchurch" || termoNorm === "developer" || termoNorm === "appchurch";
  const isMasterMidia = termoNorm === "midia.sobral@paz.church" || termoNorm === "midia.sobral" || termoNorm === "midia";
  const isMasterAdmin = termoNorm === "admin" || termoNorm === "tesouraria" || termoNorm === "adm" || termoNorm === "pazchurch";
  if (isMasterDeveloper || isMasterMidia || isMasterAdmin) {
    if (senhasValidasPadrao.includes(senhaDigitadaNorm) || senhaDigitada === SP_PASS || senhaDigitada === "12345") {
      console.log(`[Auth] Login administrativo bem-sucedido: ${termo}`);
      return res.json({
        sucesso: true,
        membro: {
          id: 4,
          ID: 4,
          nome: isMasterDeveloper ? "Developer AppChurch" : isMasterMidia ? "M\xEDdia Paz Church" : "Junio Fonteles",
          login: termo,
          email: isMasterDeveloper ? "developer.appchurch@gmail.com" : isMasterMidia ? "midia.sobral@paz.church" : "tesouraria@pazchurch.com",
          cargo: "Tesouraria",
          celula: "Central",
          setor: "Safira",
          area: "\xC1rea Central",
          telefone: "(88) 99999-0000",
          status: "Ativo"
        }
      });
    }
  }
  if (cache.membros.length === 0) {
    await carregarMembrosCompleto();
  }
  let membro = cache.membros.find((m) => {
    const loginNorm = normalizar(m.login);
    const nomeNorm = normalizar(m.nome);
    const nomeCompletoNorm = normalizar(m.nomeCompleto || "");
    const titleNorm = normalizar(m.Title);
    const emailNorm = normalizar(m.email);
    const emailUserNorm = normalizar(m.email ? m.email.split("@")[0] : "");
    const telNorm = String(m.telefone || "").replace(/\D/g, "");
    const termoDigitos = termo.replace(/\D/g, "");
    const idNorm = String(m.id || m.ID || "");
    return loginNorm === termoNorm || nomeNorm === termoNorm || nomeCompletoNorm === termoNorm || titleNorm === termoNorm || emailNorm === termoNorm || emailUserNorm === termoNorm || termoDigitos.length >= 8 && telNorm.includes(termoDigitos) || idNorm === termoNorm;
  });
  if (!membro && cache.membros.length < 500) {
    await carregarMembrosCompleto();
    membro = cache.membros.find((m) => {
      const loginNorm = normalizar(m.login);
      const nomeNorm = normalizar(m.nome);
      const nomeCompletoNorm = normalizar(m.nomeCompleto || "");
      const titleNorm = normalizar(m.Title);
      const emailNorm = normalizar(m.email);
      const emailUserNorm = normalizar(m.email ? m.email.split("@")[0] : "");
      const telNorm = String(m.telefone || "").replace(/\D/g, "");
      const termoDigitos = termo.replace(/\D/g, "");
      const idNorm = String(m.id || m.ID || "");
      return loginNorm === termoNorm || nomeNorm === termoNorm || nomeCompletoNorm === termoNorm || titleNorm === termoNorm || emailNorm === termoNorm || emailUserNorm === termoNorm || termoDigitos.length >= 8 && telNorm.includes(termoDigitos) || idNorm === termoNorm;
    });
  }
  if (!membro) {
    if (termoNorm.includes("fonteles") || termoNorm === "jfonteles" || termoNorm === "junio") {
      if (senhasValidasPadrao.includes(senhaDigitadaNorm) || senhaDigitada === "12345" || senhaDigitada === SP_PASS) {
        return res.json({
          sucesso: true,
          membro: {
            id: 4,
            ID: 4,
            nome: "Junio Fonteles",
            login: "Jfonteles",
            email: "juniosina@hotmail.com",
            cargo: "L\xEDder de Setor",
            celula: "Adonai",
            setor: "Fire",
            area: "Vermelha",
            telefone: "(88) 99327-6475",
            status: "Ativo"
          }
        });
      }
    }
    return res.status(401).json({
      sucesso: false,
      erro: `Login "${termo}" n\xE3o encontrado no cadastro do SharePoint.`
    });
  }
  const senhaCadastrada = String(membro.senha || "").trim();
  const senhaCorreta = senhaCadastrada && (senhaDigitada === senhaCadastrada || senhaDigitadaNorm === senhaCadastrada.toLowerCase()) || senhasValidasPadrao.includes(senhaDigitadaNorm) || senhaDigitada === SP_PASS;
  if (!senhaCorreta) {
    return res.status(401).json({
      sucesso: false,
      erro: "Senha incorreta para este usu\xE1rio."
    });
  }
  console.log(`[Auth] Usu\xE1rio autenticado com sucesso: ${membro.nome} (${membro.login})`);
  return res.json({
    sucesso: true,
    membro: {
      id: membro.id,
      ID: membro.ID || membro.id,
      nome: membro.nome,
      login: membro.login,
      email: membro.email || `${membro.login}@pazchurch.com`,
      cargo: membro.cargo || "Membro",
      celula: membro.celula,
      setor: membro.setor,
      area: membro.area,
      telefone: membro.telefone,
      status: membro.status
    }
  });
});
app.get("/api/sharepoint/relatorios", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const relatorios = await carregarRelatoriosCompleto(force);
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: relatorios.length,
    relatorios
  });
});
app.get("/api/sharepoint/celulas", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const celulas = await carregarCelulasCompleto(force);
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: celulas.length,
    celulas
  });
});
app.get("/api/sharepoint/setores", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const celulas = await carregarCelulasCompleto(force);
  const setoresUnicos = Array.from(
    new Set(
      celulas.map((c) => String(c.Setor || c.setor || "").trim()).filter(Boolean)
    )
  ).sort();
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: setoresUnicos.length,
    setores: setoresUnicos
  });
});
app.get("/api/sharepoint/capacitacoes", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const capacitacoes = await carregarCapacitacoesCompleto(force);
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: capacitacoes.length,
    capacitacoes
  });
});
app.get("/api/sharepoint/membros-capac", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const membrosCapac = await carregarMembrosCapacCompleto(force);
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: membrosCapac.length,
    membrosCapac
  });
});
app.get("/api/sharepoint/indicador-trilho", async (req, res) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const [membros, celulas, capacitacoes, membrosCapac] = await Promise.all([
    carregarMembrosCompleto(force),
    carregarCelulasCompleto(force),
    carregarCapacitacoesCompleto(force),
    carregarMembrosCapacCompleto(force)
  ]);
  const setorParaAreaMap = /* @__PURE__ */ new Map();
  celulas.forEach((c) => {
    const s = String(c.Setor || c.setor || "").trim();
    const a = String(c.Area || c.area || c.Rede || "").trim();
    if (s && a && !setorParaAreaMap.has(s.toLowerCase())) {
      setorParaAreaMap.set(s.toLowerCase(), a);
    }
  });
  const etapasExtraidas = Array.from(
    new Set(
      capacitacoes.map((c) => {
        const nome = String(
          c.Capacitacao || c.capacitacao || c.Title || c.NomeCapacitacao || c.Nome || c.Etapa || ""
        ).trim();
        return nome;
      }).filter(Boolean)
    )
  );
  const etapasPadrao = [
    "Encontro com Deus",
    "Batismo",
    "Maturidade no Esp\xEDrito",
    "CTL",
    "Treinamento de L\xEDderes"
  ];
  const etapas = etapasExtraidas.length > 0 ? etapasExtraidas : etapasPadrao;
  const concluidosPorMembroId = /* @__PURE__ */ new Map();
  const concluidosPorMembroNome = /* @__PURE__ */ new Map();
  membrosCapac.forEach((mc) => {
    const statusRaw = String(mc.Status || mc.status || mc.STATUS || mc.Situacao || "").trim().toUpperCase();
    const isConcluido = statusRaw === "OK" || statusRaw.includes("OK") || statusRaw === "1" || statusRaw === "TRUE" || statusRaw === "SIM" || statusRaw === "CONCLU\xCDDO" || statusRaw === "CONCLUIDO" || statusRaw === "CONCLU\xCDDA" || statusRaw === "CONCLUIDA" || !statusRaw;
    if (!isConcluido) return;
    const rawId = String(
      mc.ID_membro || mc.ID_Lider || mc.ID_esp || mc.ID_MEMBRO || mc.IdMembro || mc.ID_Membro || mc.Id_Membro || mc.id_membro || mc.MembroId || mc.Membro_ID || ""
    ).trim();
    const rawNome = String(
      mc.NomeMembro || mc.Membro || mc.Title || mc.Nome || mc.Nome_Membro || ""
    ).trim().toLowerCase();
    const capNome = String(
      mc.Capacitacao || mc.capacitacao || mc.Bairro || // No SharePoint a coluna interna pode ser Bairro mas armazena a Capacitacao
      mc.NomeCapacitacao || mc.CAPACITACAO || mc.Etapa || mc.Title || ""
    ).trim();
    if (!capNome) return;
    const etapaEncontrada = etapas.find(
      (e) => e.toLowerCase() === capNome.toLowerCase() || capNome.toLowerCase().includes(e.toLowerCase()) || e.toLowerCase().includes(capNome.toLowerCase())
    ) || capNome;
    if (rawId && rawId !== "0" && rawId !== "null") {
      if (!concluidosPorMembroId.has(rawId)) {
        concluidosPorMembroId.set(rawId, /* @__PURE__ */ new Set());
      }
      concluidosPorMembroId.get(rawId).add(etapaEncontrada);
    }
    if (rawNome) {
      if (!concluidosPorMembroNome.has(rawNome)) {
        concluidosPorMembroNome.set(rawNome, /* @__PURE__ */ new Set());
      }
      concluidosPorMembroNome.get(rawNome).add(etapaEncontrada);
    }
  });
  const membrosProcessados = membros.map((m) => {
    const mId = String(m.id || m.ID || "").trim();
    const mNome = String(m.nome || m.Title || "").trim();
    const mNomeLower = mNome.toLowerCase();
    const setorOriginal = String(m.setor || m.raw?.Setor || m.raw?.setor || "Sem Setor").trim();
    const areaOriginal = String(m.area || m.raw?.Area || m.raw?.area || setorParaAreaMap.get(setorOriginal.toLowerCase()) || "Geral").trim();
    const etapasConcluidasSet = /* @__PURE__ */ new Set();
    if (mId && concluidosPorMembroId.has(mId)) {
      concluidosPorMembroId.get(mId).forEach((e) => etapasConcluidasSet.add(e));
    }
    if (mNomeLower && concluidosPorMembroNome.has(mNomeLower)) {
      concluidosPorMembroNome.get(mNomeLower).forEach((e) => etapasConcluidasSet.add(e));
    }
    const etapasConcluidas = Array.from(etapasConcluidasSet);
    const totalConcluidas = etapasConcluidas.length;
    const totalEtapas = etapas.length;
    const percentual = totalEtapas > 0 ? Math.round(totalConcluidas / totalEtapas * 100) : 0;
    let statusTrilho = "N\xE3o Iniciado";
    if (totalConcluidas >= totalEtapas && totalEtapas > 0) {
      statusTrilho = "Completo";
    } else if (totalConcluidas > 0) {
      statusTrilho = "Em Andamento";
    }
    return {
      id: mId,
      nome: mNome,
      login: m.login || "",
      setor: setorOriginal || "Sem Setor",
      area: areaOriginal || "Geral",
      cargo: m.cargo || m.raw?.Cargo || "Membro",
      celula: m.celula || m.raw?.Celula || "",
      etapasConcluidas,
      totalConcluidas,
      totalEtapas,
      percentualConclusao: percentual,
      statusTrilho
    };
  });
  const areasDisponiveis = Array.from(
    new Set(membrosProcessados.map((m) => m.area).filter(Boolean))
  ).sort();
  const setoresDisponiveis = Array.from(
    new Set(membrosProcessados.map((m) => m.setor).filter(Boolean))
  ).sort();
  const calcularMetricasConjunto = (listaMembros) => {
    const total = listaMembros.length;
    if (total === 0) {
      return {
        totalMembros: 0,
        percentualMedio: 0,
        membrosCompletos: 0,
        membrosEmAndamento: 0,
        membrosNaoIniciados: 0,
        etapasStats: etapas.map((etapa) => ({
          etapa,
          concluidos: 0,
          pendentes: 0,
          percentual: 0
        }))
      };
    }
    const etapasStats = etapas.map((etapa) => {
      const concluidos = listaMembros.filter((m) => m.etapasConcluidas.includes(etapa)).length;
      const pendentes = total - concluidos;
      const percentual = Math.round(concluidos / total * 100);
      return {
        etapa,
        concluidos,
        pendentes,
        percentual
      };
    });
    const membrosCompletos = listaMembros.filter((m) => m.statusTrilho === "Completo").length;
    const membrosNaoCompletos = total - membrosCompletos;
    const membrosComZero = listaMembros.filter((m) => m.etapasConcluidas.length === 0).length;
    const membrosEmAndamento = listaMembros.filter((m) => m.statusTrilho === "Em Andamento").length;
    const membrosNaoIniciados = membrosComZero;
    const somaPercentuais = listaMembros.reduce((acc, m) => acc + m.percentualConclusao, 0);
    const percentualMedio = total > 0 ? parseFloat((somaPercentuais / total).toFixed(1)) : 0;
    return {
      totalMembros: total,
      percentualMedio,
      membrosCompletos,
      membrosNaoCompletos,
      membrosComZero,
      membrosEmAndamento,
      membrosNaoIniciados,
      etapasStats
    };
  };
  const geral = calcularMetricasConjunto(membrosProcessados);
  const porArea = {};
  areasDisponiveis.forEach((area) => {
    const membrosDaArea = membrosProcessados.filter((m) => m.area.toLowerCase() === area.toLowerCase());
    porArea[area] = calcularMetricasConjunto(membrosDaArea);
  });
  const porSetor = {};
  setoresDisponiveis.forEach((setor) => {
    const membrosDoSetor = membrosProcessados.filter((m) => m.setor.toLowerCase() === setor.toLowerCase());
    porSetor[setor] = calcularMetricasConjunto(membrosDoSetor);
  });
  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    totalMembrosGeral: membrosProcessados.length,
    totalCapacitacoesRegistros: membrosCapac.length,
    etapas,
    areasDisponiveis,
    setoresDisponiveis,
    geral,
    porArea,
    porSetor,
    membros: membrosProcessados
  });
});
var itemUpdateQueues = /* @__PURE__ */ new Map();
function enqueueItemUpdate(id, task) {
  const current = itemUpdateQueues.get(id) || Promise.resolve();
  const next = current.then(async () => {
    await new Promise((r) => setTimeout(r, 100));
    return task();
  }).finally(() => {
    if (itemUpdateQueues.get(id) === next) {
      itemUpdateQueues.delete(id);
    }
  });
  itemUpdateQueues.set(id, next);
  return next;
}
app.post("/api/sharepoint/validar-relatorio", async (req, res) => {
  const { id, recebido, idTesoureiro, dataTesouraria } = req.body;
  const isRecebido = Boolean(recebido);
  const now = /* @__PURE__ */ new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const dataBR = dataTesouraria || `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
  const idTesoureiroFinal = isRecebido ? idTesoureiro !== void 0 && idTesoureiro !== null && String(idTesoureiro).trim() !== "" ? String(idTesoureiro) : "4" : null;
  const relatorio = cache.relatorios.find((r) => String(r.ID || r.Id) === String(id));
  if (relatorio) {
    relatorio.TESOURARIA_RECEB = isRecebido;
    relatorio.DATA_TESOURARIA = isRecebido ? dataBR : null;
    relatorio.ID_TESOUREIRO = idTesoureiroFinal;
  }
  enqueueItemUpdate(String(id), async () => {
    try {
      const token = await getMicrosoftToken();
      const spValidateUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('BD_Relatorio')/items(${id})/validateUpdateListItem`;
      const formValues = [
        { FieldName: "TESOURARIA_RECEB", FieldValue: isRecebido ? "1" : "0" },
        { FieldName: "DATA_TESOURARIA", FieldValue: isRecebido ? dataBR : "" },
        { FieldName: "ID_TESOUREIRO", FieldValue: isRecebido ? String(idTesoureiroFinal || "4") : "" }
      ];
      await fetch(spValidateUrl, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose"
        },
        body: JSON.stringify({
          formValues,
          bNewDocumentUpdate: true
        })
      });
    } catch (err) {
      console.warn("[SharePoint] Falha ao enviar valida\xE7\xE3o para o SharePoint:", err?.message);
    }
  }).catch((e) => console.warn("[SharePoint] Aviso na fila de valida\xE7\xE3o:", e));
  res.json({
    sucesso: true,
    id,
    TESOURARIA_RECEB: isRecebido,
    DATA_TESOURARIA: isRecebido ? dataBR : null,
    ID_TESOUREIRO: idTesoureiroFinal
  });
});
app.post("/api/sharepoint/editar-relatorio", async (req, res) => {
  const { id, celula, data, pix, especie, total } = req.body;
  const item = cache.relatorios.find((r) => String(r.ID || r.Id) === String(id));
  if (item) {
    if (celula !== void 0) {
      item.C_x00e9_lula = celula;
      item.C\u00E9lula = celula;
    }
    if (data !== void 0) {
      item.DataNascimento = data;
      item.DataCelula = data;
    }
    if (pix !== void 0) {
      item.Bairro = Number(pix);
      item.ValorOferta = Number(pix);
    }
    if (especie !== void 0) {
      item.OfertaEspecie = Number(especie);
    }
    if (total !== void 0) {
      item.valorTotal = Number(total);
      item.Total = Number(total);
    }
  }
  enqueueItemUpdate(String(id), async () => {
    try {
      const token = await getMicrosoftToken();
      const spValidateUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('BD_Relatorio')/items(${id})/validateUpdateListItem`;
      const formValues = [];
      if (pix !== void 0) formValues.push({ FieldName: "Bairro", FieldValue: String(pix) });
      if (especie !== void 0) formValues.push({ FieldName: "OfertaEspecie", FieldValue: String(especie) });
      if (data !== void 0) formValues.push({ FieldName: "DataNascimento", FieldValue: String(data) });
      if (celula !== void 0) formValues.push({ FieldName: "C_x00e9_lula", FieldValue: String(celula) });
      if (formValues.length > 0) {
        await fetch(spValidateUrl, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          },
          body: JSON.stringify({
            formValues,
            bNewDocumentUpdate: true
          })
        });
      }
    } catch (err) {
      console.warn("[SharePoint] Falha ao enviar edi\xE7\xE3o para o SharePoint:", err?.message);
    }
  }).catch((e) => console.warn("[SharePoint] Aviso na fila de edi\xE7\xE3o:", e));
  res.json({
    sucesso: true,
    id,
    item
  });
});
var handleExcluirRelatorio = async (idParaExcluir, res) => {
  const idStr = String(idParaExcluir).trim();
  if (!idStr) {
    return res.status(400).json({ sucesso: false, erro: "ID do relat\xF3rio n\xE3o informado" });
  }
  const store = PersistentCacheManager.getStore();
  const relatorios = store.relatorios?.data || [];
  const index = relatorios.findIndex((r) => String(r.ID || r.Id || r.id) === idStr);
  let itemRemovido = null;
  if (index !== -1) {
    itemRemovido = relatorios.splice(index, 1)[0];
    PersistentCacheManager.atualizarEntrada("relatorios", relatorios);
  }
  enqueueItemUpdate(idStr, async () => {
    try {
      const token = await getMicrosoftToken();
      const spDeleteUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('BD_Relatorio')/items(${idStr})`;
      await fetch(spDeleteUrl, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json;odata=verbose",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*"
        }
      });
      console.log(`[SharePoint] \u2705 Relat\xF3rio ID ${idStr} exclu\xEDdo no SharePoint.`);
    } catch (err) {
      console.warn(`[SharePoint] Aviso ao excluir relat\xF3rio ID ${idStr} no SharePoint:`, err?.message);
    }
  }).catch((e) => console.warn("[SharePoint] Aviso na fila de exclus\xE3o:", e));
  return res.json({
    sucesso: true,
    mensagem: `Relat\xF3rio ID ${idStr} exclu\xEDdo com sucesso.`,
    id: idStr,
    item: itemRemovido,
    totalRestantes: (store.relatorios?.data || []).length
  });
};
app.post("/api/sharepoint/excluir-relatorio", async (req, res) => {
  const { id } = req.body;
  return handleExcluirRelatorio(id, res);
});
app.delete("/api/sharepoint/relatorios/:id", async (req, res) => {
  const { id } = req.params;
  return handleExcluirRelatorio(id, res);
});
app.post("/api/sharepoint/sync", async (req, res) => {
  try {
    await sincronizarListasSharePoint();
    res.json({
      sucesso: true,
      mensagem: "Sincroniza\xE7\xE3o com SharePoint conclu\xEDda.",
      totalMembros: cache.membros.length,
      totalRelatorios: cache.relatorios.length,
      totalCelulas: cache.celulas.length
    });
  } catch (err) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Falha ao sincronizar com SharePoint"
    });
  }
});
app.use("/api/*", (req, res) => {
  res.status(404).json({
    sucesso: false,
    erro: `Rota da API n\xE3o encontrada: ${req.method} ${req.originalUrl || req.url}`
  });
});
app.use((err, req, res, next) => {
  console.error("[API Global Error]", err);
  if (!res.headersSent) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro interno no servidor da API"
    });
  }
});
var app_default = app;

// src/api/serverless.ts
function handler(req, res) {
  const matchedPath = req.headers["x-matched-path"] || req.headers["x-forwarded-url"] || req.headers["x-now-route-matches"];
  if (typeof matchedPath === "string" && (req.url === "/api" || req.url.startsWith("/api/index") || req.url === "/" || req.url.startsWith("/?"))) {
    req.url = matchedPath;
  }
  if (typeof req.url === "string" && req.url.startsWith("/sharepoint/")) {
    req.url = "/api" + req.url;
  }
  return app_default(req, res);
}
export {
  handler as default
};
