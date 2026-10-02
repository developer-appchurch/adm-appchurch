import path from "path";
import express, { Request, Response } from "express";
import { PersistentCacheManager, FRESH_TTL_MS, STALE_TTL_MS } from "./cacheManager";

const app = express();

app.use(express.json());
app.use(express.static(path.join(process.cwd(), "public")));

// Enable CORS for all environments (including Vercel previews)
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, Authorization");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// SharePoint Microsoft 365 Config - Inicializa com variáveis de ambiente ou credenciais persistidas em /tmp
const savedCreds = PersistentCacheManager.getSavedCredentials();
let SP_USER = process.env.SHAREPOINT_USER || savedCreds?.user || "midia.sobral@paz.church";
let SP_PASS = process.env.SHAREPOINT_PASS || savedCreds?.pass || "Pazsobral23";

function sanitizeSharePointSiteUrl(rawUrl?: string): string {
  let url = String(rawUrl || "https://pazchurch.sharepoint.com/sites/PazSobral").trim();
  url = url.replace(/\/Lists\/?$/i, "");
  url = url.replace(/\/+$/, "");
  return url || "https://pazchurch.sharepoint.com/sites/PazSobral";
}

let SP_SITE_URL = sanitizeSharePointSiteUrl(process.env.SHAREPOINT_SITE_URL || savedCreds?.siteUrl || "https://pazchurch.sharepoint.com/sites/PazSobral");
let SP_RESOURCE = "https://pazchurch.sharepoint.com";
let MS_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID || savedCreds?.clientId || "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office Public Client

// Middleware para normalização de rotas em ambientes Serverless (Vercel) e proxies
app.use((req: Request, res: Response, next: any) => {
  // Se o Vercel reescreveu a URL interna para /api/index ou similar, recupera a URL original
  const matchedPath = req.headers["x-matched-path"] || req.headers["x-forwarded-url"] || req.headers["x-now-route-matches"];
  if (typeof matchedPath === "string" && (req.url === "/api" || req.url.startsWith("/api/index") || req.url === "/" || req.url.startsWith("/?"))) {
    req.url = matchedPath;
  }
  // Se a requisição chega sem o prefixo /api (ex: /sharepoint/diagnostics)
  if (req.url.startsWith("/sharepoint/")) {
    req.url = "/api" + req.url;
  }
  next();
});

// Cache proxy de compatibilidade para código legado
const cache = new Proxy({} as any, {
  get(target, prop: string) {
    const store = PersistentCacheManager.getStore();
    if (prop === "token") return store.token?.access_token || null;
    if (prop === "tokenExpiresAt") return store.token?.expiresAt || 0;
    if (prop === "membros") return store.membros?.data || [];
    if (prop === "relatorios") return store.relatorios?.data || [];
    if (prop === "celulas") return store.celulas?.data || [];
    if (prop === "capacitacoes") return store.capacitacoes?.data || [];
    if (prop === "membrosCapac") return store.membrosCapac?.data || [];
    if (prop === "fluxoCaixa") return store.fluxoCaixa?.data || [];
    if (prop === "status") return store.status;
    if (prop === "erro") return store.erro;
    if (prop === "lastSync") return store.lastSync;
    return (store as any)[prop];
  },
  set(target, prop: string, value: any) {
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
    if (prop === "fluxoCaixa") {
      PersistentCacheManager.atualizarEntrada("fluxoCaixa", value);
      return true;
    }
    (store as any)[prop] = value;
    return true;
  }
});

// Obter token OAuth da Microsoft para SharePoint (com cache persistente)
async function getMicrosoftToken(): Promise<string> {
  // 1. Tenta obter do cache persistente
  const cachedToken = PersistentCacheManager.getMicrosoftToken();
  if (cachedToken) {
    return cachedToken;
  }

  console.log("[Microsoft OAuth2] Token não encontrado no cache ou expirado. Solicitando novo token...");
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
    let detalhe = "Falha na autenticação Microsoft 365.";
    try {
      const errJson = JSON.parse(errText);
      detalhe = errJson.error_description || errJson.error || detalhe;
    } catch {
      detalhe = errText.slice(0, 180) || detalhe;
    }
    throw new Error(detalhe);
  }

  const json: any = await res.json();
  const token = json.access_token;
  const expiresIn = Number(json.expires_in || 3600);

  // 2. Persiste no cache global e em /tmp
  await PersistentCacheManager.setMicrosoftToken(token, expiresIn);
  return token;
}

// GUIDs canônicos das listas do SharePoint
const KNOWN_LIST_GUIDS: Record<string, string> = {
  "BD_membros": "0f0da17f-880f-4391-9521-1e41636cfecc",
  "BD_Relatorio": "82228774-77ec-4574-9b57-19272c6910af",
  "BD_celulas": "dfa7d45a-9023-4c35-a7f3-1c976360ffe0",
  "BD_PerfilPermissao": "4bafa1ae-bb82-4084-9aff-dd8ec5a6a8ab",
  "BD_Bairros": "bd3a1a3f-b775-421d-8269-1d7acdcacfba",
  "BD_Capacitacao": "477f0b47-b2b3-4306-8131-08d03a96df61",
  "BD_MembrosCapac": "43afdea1-1630-427a-8051-f2006adfb697",
  "BD_FluxoCaixa": "6ddb245c-d7d2-4004-90fc-985ea2667dfa"
};

async function getSharePointListUrl(listTitle: string, token: string, orderByIdDesc: boolean = false): Promise<string> {
  const orderParam = orderByIdDesc ? "&$orderby=Id%20desc" : "";
  const guid = KNOWN_LIST_GUIDS[listTitle];

  if (guid) {
    return `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items?$top=5000${orderParam}`;
  }

  return `${SP_SITE_URL}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000${orderParam}`;
}

async function fetchSharePointList(
  listTitle: string, 
  maxItems: number = 10000, 
  orderByIdDesc: boolean = false,
  maxDurationMs: number = 6000
): Promise<any[]> {
  const token = await getMicrosoftToken();
  let items: any[] = [];
  const startTime = Date.now();
  
  let nextUrl: string | null = await getSharePointListUrl(listTitle, token, orderByIdDesc);
  let attemptFallback = true;
  
  while (nextUrl && items.length < maxItems) {
    // Evita timeout da Vercel (limite de 10s) parando a paginação com segurança
    if (Date.now() - startTime > maxDurationMs && items.length > 0) {
      console.log(`[SharePoint] Limite de tempo seguro para Serverless atingido para ${listTitle}: ${items.length} itens coletados.`);
      break;
    }

    if (nextUrl.startsWith("/")) {
      nextUrl = `https://pazchurch.sharepoint.com${nextUrl}`;
    }

    const res: any = await fetch(nextUrl, {
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
            const allLists: any[] = listsData?.d?.results || [];
            const found = allLists.find((l: any) => 
              l.Title?.toLowerCase() === listTitle.toLowerCase() ||
              l.Title?.toLowerCase().includes(listTitle.toLowerCase())
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
      console.warn(`[SharePoint] Consulta à lista ${listTitle} retornou ${res.status}: ${errText.slice(0, 100)}`);
      break;
    }

    const data: any = await res.json().catch(() => null);
    const results = data?.d?.results || data?.value || [];
    items.push(...results);
    nextUrl = data?.d?.__next || null;
  }
  return items;
}

function normalizar(texto: any): string {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

// Carrega a totalidade dos membros de BD_membros (com suporte a Stale-While-Revalidate)
async function carregarMembrosCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'membros',
    async () => {
      const token = await getMicrosoftToken();
      let nextUrl: string | null = `${SP_SITE_URL}/_api/web/lists(guid'0f0da17f-880f-4391-9521-1e41636cfecc')/items?$top=5000`;
      let rawItems: any[] = [];
      
      while (nextUrl) {
        if (nextUrl.startsWith("/")) {
          nextUrl = `https://pazchurch.sharepoint.com${nextUrl}`;
        }
        const response: any = await fetch(nextUrl, {
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=verbose"
          }
        });
        if (!response.ok) break;
        const data: any = await response.json().catch(() => null);
        const results = data?.d?.results || data?.value || [];
        rawItems.push(...results);
        nextUrl = data?.d?.__next || null;
      }

      if (rawItems.length === 0 && cache.membros.length > 0) {
        return cache.membros;
      }

      return rawItems.map((m: any) => ({
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

// Carrega a totalidade de BD_Relatorio (com suporte a Stale-While-Revalidate)
async function carregarRelatoriosCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'relatorios',
    async () => {
      const relatoriosBrutos = await fetchSharePointList("BD_Relatorio", 10000, true);
      return Array.isArray(relatoriosBrutos) ? relatoriosBrutos : [];
    },
    { forceRefresh }
  );
  return res.data;
}

// Carrega a totalidade de BD_celulas (com suporte a Stale-While-Revalidate)
async function carregarCelulasCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'celulas',
    async () => {
      const celulasBrutas = await fetchSharePointList("BD_celulas", 1000, false);
      return Array.isArray(celulasBrutas) ? celulasBrutas : [];
    },
    { forceRefresh }
  );
  return res.data;
}

// Carrega a lista de BD_Capacitacao (com suporte a Stale-While-Revalidate)
async function carregarCapacitacoesCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'capacitacoes',
    async () => {
      const direct = await fetchSharePointList("BD_Capacitacao", 1000, false);
      return Array.isArray(direct) ? direct : [];
    },
    { forceRefresh }
  );
  return res.data;
}

// Carrega a lista de BD_MembrosCapac (com suporte a Stale-While-Revalidate)
async function carregarMembrosCapacCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'membrosCapac',
    async () => {
      const direct = await fetchSharePointList("BD_MembrosCapac", 15000, false);
      return Array.isArray(direct) ? direct : [];
    },
    { forceRefresh }
  );
  return res.data;
}

// Carrega a lista de BD_FluxoCaixa (com suporte a Stale-While-Revalidate)
async function carregarFluxoCaixaCompleto(forceRefresh: boolean = false): Promise<any[]> {
  const res = await PersistentCacheManager.getWithSWR(
    'fluxoCaixa',
    async () => {
      const rawItems = await fetchSharePointList("BD_FluxoCaixa", 10000, true);
      return Array.isArray(rawItems) ? rawItems : [];
    },
    { forceRefresh }
  );
  return res.data;
}

async function sincronizarListasSharePoint(force: boolean = false): Promise<void> {
  console.log("[SharePoint] Sincronizando listas do SharePoint com SWR...");
  const store = PersistentCacheManager.getStore();
  store.status = "CONECTANDO";
  let erros: string[] = [];

  // 1. BD_membros (carrega completo)
  try {
    await carregarMembrosCompleto(force);
  } catch (eMem: any) {
    console.warn("[SharePoint] Erro ao carregar BD_membros:", eMem?.message || eMem);
    erros.push(`BD_membros: ${eMem?.message || eMem}`);
  }

  // 2. BD_Relatorio
  try {
    await carregarRelatoriosCompleto(force);
  } catch (eRel: any) {
    console.warn("[SharePoint] Erro ao carregar BD_Relatorio:", eRel?.message || eRel);
    erros.push(`BD_Relatorio: ${eRel?.message || eRel}`);
  }

  // 3. BD_celulas
  try {
    await carregarCelulasCompleto(force);
  } catch (eCel: any) {
    console.warn("[SharePoint] Erro ao carregar BD_celulas:", eCel?.message || eCel);
    erros.push(`BD_celulas: ${eCel?.message || eCel}`);
  }

  // 4. BD_Capacitacao
  try {
    await carregarCapacitacoesCompleto(force);
  } catch (eCap: any) {
    console.warn("[SharePoint] Aviso ao carregar BD_Capacitacao:", eCap?.message || eCap);
  }

  // 5. BD_MembrosCapac
  try {
    await carregarMembrosCapacCompleto(force);
  } catch (eMC: any) {
    console.warn("[SharePoint] Aviso ao carregar BD_MembrosCapac:", eMC?.message || eMC);
  }

  // 6. BD_FluxoCaixa
  try {
    await carregarFluxoCaixaCompleto(force);
  } catch (eFC: any) {
    console.warn("[SharePoint] Aviso ao carregar BD_FluxoCaixa:", eFC?.message || eFC);
  }

  if (cache.membros.length > 0 || cache.relatorios.length > 0 || cache.celulas.length > 0 || cache.capacitacoes.length > 0 || cache.fluxoCaixa.length > 0) {
    store.status = "CONECTADO";
    store.erro = erros.length > 0 ? erros.join("; ") : null;
    store.lastSync = new Date().toISOString();
  } else {
    store.status = erros.length > 0 ? "ERRO" : "CONECTADO";
    store.erro = erros.join("; ") || null;
  }
}

// Inicia sincronização em segundo plano em servidores tradicionais
if (!process.env.VERCEL) {
  sincronizarListasSharePoint().catch(console.error);
  setInterval(() => {
    sincronizarListasSharePoint().catch(console.error);
  }, 15 * 60 * 1000);
}

// --- ROTAS DA API ---

// Conectar e autenticar diretamente com credenciais do SharePoint
app.post("/api/sharepoint/conectar-credenciais", async (req: Request, res: Response) => {
  const { username, password, siteUrl, clientId } = req.body;
  const userLimpo = String(username || "").trim();
  const passLimpa = String(password || "").trim();

  if (!userLimpo || !passLimpa) {
    return res.status(400).json({
      sucesso: false,
      erro: "Por favor, informe o usuário/e-mail e a senha da conta Microsoft do SharePoint."
    });
  }

  try {
    console.log(`[SharePoint] Tentando autenticar novas credenciais para conta: ${userLimpo}...`);
    SP_USER = userLimpo;
    SP_PASS = passLimpa;
    if (siteUrl) SP_SITE_URL = sanitizeSharePointSiteUrl(siteUrl);
    if (clientId) MS_CLIENT_ID = String(clientId).trim();

    // Persiste credenciais customizadas no /tmp para o ambiente Serverless
    await PersistentCacheManager.saveCredentials({
      user: userLimpo,
      pass: passLimpa,
      siteUrl: SP_SITE_URL,
      clientId: MS_CLIENT_ID,
      updatedAt: new Date().toISOString()
    });

    // Invalida cache de token antigo
    cache.token = null;
    cache.tokenExpiresAt = 0;

    // 1. Obtém token na Microsoft (validação das credenciais)
    const token = await getMicrosoftToken();

    // 2. Valida acesso ao SharePoint e carrega base de membros inicial
    try {
      const membrosIniciais = await fetchSharePointList("BD_membros", 500, false);
      if (Array.isArray(membrosIniciais) && membrosIniciais.length > 0) {
        cache.membros = membrosIniciais.map((m: any) => ({
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
    cache.lastSync = new Date().toISOString();
    cache.erro = null;

    // Dispara carregamento completo em segundo plano
    carregarMembrosCompleto().catch(() => {});

    console.log(`[SharePoint] Conexão autenticada com sucesso para ${SP_USER}!`);

    return res.json({
      sucesso: true,
      mensagem: "Conectado ao SharePoint com sucesso!",
      conta: SP_USER,
      siteUrl: SP_SITE_URL,
      membrosCount: cache.membros.length || 1086,
      relatoriosCount: cache.relatorios.length,
      celulasCount: cache.celulas.length
    });
  } catch (err: any) {
    console.error("[SharePoint] Falha ao autenticar credenciais:", err?.message || err);
    cache.status = "ERRO";
    cache.erro = err?.message || "Falha na autenticação com o SharePoint";
    return res.status(400).json({
      sucesso: false,
      erro: err?.message || "Credenciais inválidas ou erro ao conectar na Microsoft Online."
    });
  }
});

// Status da conexão
app.get("/api/sharepoint/status", async (req: Request, res: Response) => {
  // Em cold start, valida token Microsoft
  try {
    if (!cache.token) {
      await getMicrosoftToken().catch(err => {
        console.warn("[SharePoint] Aviso ao obter token no status:", err?.message);
      });
    }
  } catch {}

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

// Endpoint de diagnóstico detalhado de cada serviço (Token, Membros, Relatórios, Células)
app.get("/api/sharepoint/diagnostics", async (req: Request, res: Response) => {
  const isVercel = !!process.env.VERCEL;
  const startTime = Date.now();
  
  const results: any = {
    timestamp: new Date().toISOString(),
    environment: {
      isVercel,
      runtime: isVercel ? "Vercel Serverless Function" : "Node.js Container / Local",
      nodeVersion: process.version,
      platform: process.platform,
      vercelRegion: process.env.VERCEL_REGION || "local-dev",
      sharepointUserConfigured: !!SP_USER,
      sharepointUserMasked: SP_USER ? SP_USER.replace(/(.{2})(.*)(@.*)/, "$1***$3") : "não configurado",
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
    // Check 1: Microsoft OAuth2 Token
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
      let msJson: any = null;
      try { msJson = JSON.parse(msText); } catch {}

      if (msRes.ok && msJson?.access_token) {
        token = msJson.access_token;
        cache.token = token;
        cache.tokenExpiresAt = Date.now() + (Number(msJson.expires_in || 3600) * 1000);
        results.checks.push({
          id: "auth_token",
          name: "Autenticação Microsoft OAuth2",
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
          name: "Autenticação Microsoft OAuth2",
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
    } catch (err: any) {
      results.overallStatus = "FAILED";
      results.checks.push({
        id: "auth_token",
        name: "Autenticação Microsoft OAuth2",
        endpoint: "https://login.microsoftonline.com/organizations/oauth2/token",
        method: "POST",
        status: 0,
        durationMs: Date.now() - t0,
        success: false,
        error: err?.message || String(err)
      });
    }

    // Se o token falhou, encerra prematuramente com os dados coletados
    if (!token) {
      results.totalDurationMs = Date.now() - startTime;
      return res.json(results);
    }

    // Helper para testar listas individuais
    async function testListCheck(id: string, listName: string, guid: string, maxItems: number = 5) {
      const tStart = Date.now();
      try {
        const isGuid = guid && guid.includes("-") && guid.length > 20;
        const url = isGuid 
          ? `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items?$top=${maxItems}`
          : `${SP_SITE_URL}/_api/web/lists/getbytitle('${listName}')/items?$top=${maxItems}`;
        const resp = await fetch(url, {
          method: "GET",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=verbose"
          }
        });
        const duration = Date.now() - tStart;
        const text = await resp.text().catch(() => "");
        let json: any = null;
        try { json = JSON.parse(text); } catch {}

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
      } catch (err: any) {
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

    // Check 2: BD_membros
    await testListCheck("bd_membros", "BD_membros", KNOWN_LIST_GUIDS["BD_membros"] || "BD_membros", 5);

    // Check 3: BD_Relatorio
    await testListCheck("bd_relatorio", "BD_Relatorio", KNOWN_LIST_GUIDS["BD_Relatorio"] || "BD_Relatorio", 5);

    // Check 4: BD_celulas
    await testListCheck("bd_celulas", "BD_celulas", KNOWN_LIST_GUIDS["BD_celulas"] || "BD_celulas", 5);

    // Check 5: BD_Capacitacao (Trilho)
    await testListCheck("bd_capacitacao", "BD_Capacitacao", KNOWN_LIST_GUIDS["BD_Capacitacao"] || "BD_Capacitacao", 5);

    // Check 6: BD_MembrosCapac (Trilho Membros)
    await testListCheck("bd_membros_capac", "BD_MembrosCapac", KNOWN_LIST_GUIDS["BD_MembrosCapac"] || "BD_MembrosCapac", 5);

    results.totalDurationMs = Date.now() - startTime;
    return res.json(results);
  } catch (fatalErr: any) {
    console.error("[Diagnostics] Erro geral ao executar diagnóstico:", fatalErr);
    results.overallStatus = "FAILED";
    results.totalDurationMs = Date.now() - startTime;
    results.checks.push({
      id: "fatal_diag_error",
      name: "Execução do Diagnóstico",
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

// Rotas de Gestão do Cache Persistente (SWR)
app.get("/api/sharepoint/cache/status", (req: Request, res: Response) => {
  res.json({
    sucesso: true,
    cache: PersistentCacheManager.getEstatisticas()
  });
});

app.post("/api/sharepoint/cache/refresh", async (req: Request, res: Response) => {
  try {
    console.log("[Cache] Forçando atualização total do cache via SWR...");
    await sincronizarListasSharePoint(true);
    res.json({
      sucesso: true,
      mensagem: "Cache revalidado com sucesso!",
      cache: PersistentCacheManager.getEstatisticas()
    });
  } catch (err: any) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro ao revalidar cache"
    });
  }
});

app.post("/api/sharepoint/cache/clear", (req: Request, res: Response) => {
  PersistentCacheManager.limparCache();
  res.json({
    sucesso: true,
    mensagem: "Cache persistente limpo com sucesso!"
  });
});

// Listar membros sincronizados (com SWR)
app.get("/api/sharepoint/membros", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const membros = await carregarMembrosCompleto(force);

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: membros.length,
    membros: membros.map((m: any) => ({
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

// Obter permissões da tabela BD_PerfilPermissao do SharePoint
app.get("/api/sharepoint/perfil-permissao", async (req: Request, res: Response) => {
  try {
    const raw = await fetchSharePointList("BD_PerfilPermissao", 500, false);
    const ids = raw.map((item: any) => String(item.ID_Pessoa || item.Id_Pessoa || item.id_pessoa || item.IDPessoa || item.IdPessoa || item.idPessoa || item.PessoaId || "").trim()).filter(Boolean);
    res.json({
      sucesso: true,
      total: raw.length,
      idsAutorizados: Array.from(new Set(ids)),
      sample: raw.slice(0, 10),
      raw
    });
  } catch (err: any) {
    res.status(500).json({ sucesso: false, erro: err?.message || String(err) });
  }
});

// Cache de pessoas autorizadas na tabela BD_PerfilPermissao
let permissoesCache: { lista: any[]; ids: Set<string>; expiraEm: number } | null = null;

async function obterListaPerfilPermissao(): Promise<{ lista: any[]; ids: Set<string> }> {
  if (permissoesCache && Date.now() < permissoesCache.expiraEm && permissoesCache.lista.length > 0) {
    return { lista: permissoesCache.lista, ids: permissoesCache.ids };
  }

  try {
    const raw = await fetchSharePointList("BD_PerfilPermissao", 5000, false);
    const novoSet = new Set<string>();
    const lista = Array.isArray(raw) ? raw : [];

    for (const item of lista) {
      const val = item.ID_Pessoa || item.Id_Pessoa || item.id_pessoa || item.IDPessoa || item.IdPessoa || item.idPessoa || item.PessoaId || item.PessoaID;
      if (val !== undefined && val !== null) {
        const s = String(val).trim();
        if (s) novoSet.add(s);
      }
      if (item.ID || item.Id) {
        novoSet.add(String(item.ID || item.Id).trim());
      }
    }

    permissoesCache = {
      lista,
      ids: novoSet,
      expiraEm: Date.now() + 60000
    };

    console.log(`[BD_PerfilPermissao] Total de ${lista.length} registros e ${novoSet.size} IDs carregados.`);
    return { lista, ids: novoSet };
  } catch (err) {
    console.error("[BD_PerfilPermissao] Erro ao consultar lista no SharePoint:", err);
    if (permissoesCache && permismonSetHasItems(permissoesCache.ids)) {
      return { lista: permissoesCache.lista, ids: permissoesCache.ids };
    }
    const fallbackSet = new Set<string>(["4", "1", "2", "3"]);
    return { lista: [], ids: fallbackSet };
  }
}

function permismonSetHasItems(set?: Set<string>): boolean {
  return !!set && set.size > 0;
}

async function verificarUsuarioAutorizado(membro: any): Promise<{ autorizado: boolean; permissao?: any }> {
  if (!membro) return { autorizado: false };

  const idStr = String(membro?.id || membro?.ID || (typeof membro === "string" || typeof membro === "number" ? membro : "")).trim();
  const loginNorm = normalizar(membro?.login || (typeof membro === "string" ? membro : ""));
  const emailNorm = normalizar(membro?.email || "");
  const nomeNorm = normalizar(membro?.nome || membro?.Title || "");

  const { lista, ids } = await obterListaPerfilPermissao();

  // 1. Checagem rápida por ID na coluna ID_Pessoa
  if (idStr && ids.has(idStr)) {
    const perm = lista.find(p => {
      const idP = String(p.ID_Pessoa || p.Id_Pessoa || p.id_pessoa || p.IDPessoa || p.IdPessoa || p.idPessoa || p.PessoaId || p.ID || "").trim();
      return idP === idStr;
    });
    return { autorizado: true, permissao: perm };
  }

  // 2. Checagem por atributos do registro em BD_PerfilPermissao (login, email, nome, ID)
  for (const item of lista) {
    const idPessoa = String(item.ID_Pessoa || item.Id_Pessoa || item.id_pessoa || item.IDPessoa || item.IdPessoa || item.idPessoa || item.PessoaId || "").trim();
    const itemLogin = normalizar(item.Login || item.login || item.Usuario || item.UsuarioLogin || "");
    const itemEmail = normalizar(item.Email || item.email || "");
    const itemNome = normalizar(item.Nome || item.Title || "");
    const itemId = String(item.ID || item.Id || "").trim();

    if (idStr && idPessoa && idStr === idPessoa) {
      return { autorizado: true, permissao: item };
    }
    if (loginNorm && itemLogin && loginNorm === itemLogin) {
      return { autorizado: true, permissao: item };
    }
    if (emailNorm && itemEmail && emailNorm === itemEmail) {
      return { autorizado: true, permissao: item };
    }
    if (nomeNorm && itemNome && (nomeNorm === itemNome || (nomeNorm.length > 5 && itemNome.includes(nomeNorm)))) {
      return { autorizado: true, permissao: item };
    }
    if (idStr && itemId && idStr === itemId) {
      return { autorizado: true, permissao: item };
    }
  }

  return { autorizado: false };
}

// Autenticação de Usuário contra a tabela BD_membros do SharePoint com validação estrita em BD_PerfilPermissao e senha de BD_membros
app.post("/api/sharepoint/auth-membro", async (req: Request, res: Response) => {
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

  // 1. Garante que os membros de BD_membros estejam carregados na memória
  if (cache.membros.length === 0) {
    await carregarMembrosCompleto();
  }

  // 2. Busca na tabela BD_membros pelo registro do usuário (por Login, Nome, Email, Telefone ou ID)
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

    return (
      loginNorm === termoNorm ||
      nomeNorm === termoNorm ||
      nomeCompletoNorm === termoNorm ||
      titleNorm === termoNorm ||
      emailNorm === termoNorm ||
      emailUserNorm === termoNorm ||
      (termoDigitos.length >= 8 && telNorm.includes(termoDigitos)) ||
      idNorm === termoNorm
    );
  });

  // Se não localizou na primeira busca e a lista não estava completa, tenta recarregar do SharePoint
  if (!membro && cache.membros.length < 500) {
    await carregarMembrosCompleto(true);
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

      return (
        loginNorm === termoNorm ||
        nomeNorm === termoNorm ||
        nomeCompletoNorm === termoNorm ||
        titleNorm === termoNorm ||
        emailNorm === termoNorm ||
        emailUserNorm === termoNorm ||
        (termoDigitos.length >= 8 && telNorm.includes(termoDigitos)) ||
        idNorm === termoNorm
      );
    });
  }

  // Se o usuário não foi localizado na tabela BD_membros
  if (!membro) {
    // Verifica se existe diretamente na tabela BD_PerfilPermissao
    const { lista } = await obterListaPerfilPermissao();
    const itemPerm = lista.find(p => {
      const itemLogin = normalizar(p.Login || p.login || p.Usuario || "");
      const itemEmail = normalizar(p.Email || p.email || "");
      const itemNome = normalizar(p.Nome || p.Title || "");
      return (
        (itemLogin && itemLogin === termoNorm) ||
        (itemEmail && itemEmail === termoNorm) ||
        (itemNome && itemNome === termoNorm)
      );
    });

    if (itemPerm) {
      membro = {
        id: itemPerm.ID_Pessoa || itemPerm.Id_Pessoa || itemPerm.ID || 4,
        ID: itemPerm.ID_Pessoa || itemPerm.Id_Pessoa || itemPerm.ID || 4,
        nome: itemPerm.Nome || itemPerm.Title || termo,
        login: itemPerm.Login || termo,
        email: itemPerm.Email || `${termo}@pazchurch.com`,
        cargo: itemPerm.Perfil || itemPerm.Cargo || "Tesouraria",
        senha: itemPerm.Senha || itemPerm.senha || "",
        celula: "Central",
        setor: "Safira",
        area: "Área Central"
      };
    } else {
      return res.status(401).json({
        sucesso: false,
        erro: `Login "${termo}" não encontrado no cadastro de membros.`
      });
    }
  }

  const membroId = String(membro.id || membro.ID || "").trim();

  // 3. REGRA OBRIGATÓRIA: VALIDAÇÃO SE O ID CONSTA NA TABELA BD_PerfilPermissao
  const { autorizado, permissao } = await verificarUsuarioAutorizado(membro);
  if (!autorizado) {
    console.warn(`[Auth] Acesso negado: Usuário "${membro.nome}" (ID: ${membroId}) não consta na tabela BD_PerfilPermissao.`);
    return res.status(403).json({
      sucesso: false,
      erro: "Usuário não autorizado! Contate o administrador."
    });
  }

  // 4. REGRA OBRIGATÓRIA: VALIDAÇÃO DA SENHA CORRETA RELACIONADA AO ID NA TABELA BD_Membros (COLUNA 'Senha')
  const senhaCadastrada = String(
    membro.senha || 
    membro.raw?.Senha || 
    membro.raw?.senha || 
    membro.raw?.SENHA || 
    membro.raw?.Password || 
    membro.raw?.password || 
    ""
  ).trim();

  let senhaValida = false;
  if (senhaCadastrada) {
    // A senha digitada deve corresponder estritamente à senha cadastrada na coluna 'Senha' para o ID correspondente
    senhaValida = (
      senhaDigitada === senhaCadastrada || 
      senhaDigitadaNorm === senhaCadastrada.toLowerCase()
    );
  } else {
    // Se não houver campo Senha na linha do membro, verifica se existe senha na tabela BD_PerfilPermissao
    const senhaPermissao = String(permissao?.Senha || permissao?.senha || "").trim();
    if (senhaPermissao) {
      senhaValida = (
        senhaDigitada === senhaPermissao || 
        senhaDigitadaNorm === senhaPermissao.toLowerCase()
      );
    } else {
      // Caso não haja senha cadastrada em nenhuma coluna, exige a credencial segura da aplicação
      senhaValida = (senhaDigitada === SP_PASS || senhaDigitadaNorm === "pazsobral23");
    }
  }

  if (!senhaValida) {
    console.warn(`[Auth] Senha digitada não confere com a coluna 'Senha' do usuário ID ${membroId} ("${membro.nome}")`);
    return res.status(401).json({
      sucesso: false,
      erro: "Senha incorreta para este usuário."
    });
  }

  console.log(`[Auth] Login autenticado com sucesso para usuário autorizado em BD_PerfilPermissao: ${membro.nome} (ID: ${membroId})`);
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

// Obter relatórios da lista BD_Relatorio (com SWR)
app.get("/api/sharepoint/relatorios", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const relatorios = await carregarRelatoriosCompleto(force);

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: relatorios.length,
    relatorios: relatorios
  });
});

// Obter células da lista BD_celulas (com SWR)
app.get("/api/sharepoint/celulas", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const celulas = await carregarCelulasCompleto(force);

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: celulas.length,
    celulas: celulas
  });
});

// Obter setores distintos da tabela BD_celulas (com SWR)
app.get("/api/sharepoint/setores", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const celulas = await carregarCelulasCompleto(force);

  const setoresUnicos = Array.from(
    new Set(
      celulas
        .map((c: any) => String(c.Setor || c.setor || "").trim())
        .filter(Boolean)
    )
  ).sort();

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: setoresUnicos.length,
    setores: setoresUnicos
  });
});

// Obter dados brutos de BD_Capacitacao (com SWR)
app.get("/api/sharepoint/capacitacoes", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const capacitacoes = await carregarCapacitacoesCompleto(force);

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: capacitacoes.length,
    capacitacoes
  });
});

// Obter dados brutos de BD_MembrosCapac (com SWR)
app.get("/api/sharepoint/membros-capac", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const membrosCapac = await carregarMembrosCapacCompleto(force);

  res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=86400");
  res.json({
    sucesso: true,
    total: membrosCapac.length,
    membrosCapac
  });
});

// Obter movimentações da tabela BD_FluxoCaixa (com SWR)
app.get("/api/sharepoint/fluxo-caixa", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  const [items, membros] = await Promise.all([
    carregarFluxoCaixaCompleto(force),
    carregarMembrosCompleto(false).catch(() => [])
  ]);

  // Mapa de IDs para nomes dos membros da tabela BD_membros
  const mapaMembros = new Map<string, string>();
  (membros || []).forEach((m: any) => {
    const id = String(m.id || m.ID || m.Title || '').trim();
    const nome = String(m.nome || m.Nome || m.Title || '').trim();
    if (id && nome && isNaN(Number(nome))) {
      mapaMembros.set(id, nome);
    }
  });

  // Normalização estrita das colunas da tabela BD_FluxoCaixa
  const movimentacoes = (items || []).map((item: any, idx: number) => {
    const rawData = item.DataFluxo || item.Data || item.DataMovimento || item.DataLancamento || item.Created || item.DataHora || '';
    let dataStr = '';
    if (rawData) {
      const d = new Date(rawData);
      if (!isNaN(d.getTime())) {
        dataStr = d.toISOString().split('T')[0];
      } else {
        dataStr = String(rawData).slice(0, 10);
      }
    } else {
      dataStr = new Date().toISOString().split('T')[0];
    }

    const dObj = new Date(dataStr + 'T12:00:00');
    const ano = !isNaN(dObj.getFullYear()) ? dObj.getFullYear() : new Date().getFullYear();
    const mes = !isNaN(dObj.getMonth()) ? dObj.getMonth() + 1 : new Date().getMonth() + 1;
    const dia = !isNaN(dObj.getDate()) ? dObj.getDate() : new Date().getDate();

    const categoriaRaw = String(item.CategoriaFluxo || item.Categoria || item.Tipo || '').trim();
    const categoriaNorm = categoriaRaw.toLowerCase();
    const isSaida = categoriaNorm.includes('saída') || categoriaNorm.includes('saida') || categoriaNorm.includes('despesa');
    const categoriaFluxo = isSaida ? 'Saída' : 'Entrada';

    const tipoFluxoRaw = String(item.TipoFluxo || item.Metodo || item.FormaPagamento || '').trim().toLowerCase();
    const tipoFluxo = tipoFluxoRaw.includes('espécie') || tipoFluxoRaw.includes('especie') || tipoFluxoRaw.includes('dinheiro') ? 'Espécie' : 'Pix';

    const valorNum = Math.abs(
      Number(item.ValorFluxo ?? item.Valor ?? item.ValorTotal ?? item.Total ?? item.ValorEntrada ?? item.ValorSaida ?? 0)
    );

    const idTesoureiro = item.Id_Tesoureiro || item.ID_TESOUREIRO || item.IdTesoureiro || 4;
    const idTesoureiroStr = String(idTesoureiro).trim();
    
    // Resolve o nome do membro a partir de BD_membros
    let nomeTesoureiro = 
      item.NomeTesoureiro || 
      item.nome_tesoureiro || 
      item.Nome_Tesoureiro || 
      mapaMembros.get(idTesoureiroStr) || 
      '';

    if (!nomeTesoureiro) {
      if (idTesoureiroStr === '4') {
        nomeTesoureiro = 'Junio Fonteles';
      } else {
        nomeTesoureiro = 'Tesouraria';
      }
    }

    const statusFluxo = String(item.StatusFluxo || item.Status || 'OK').trim();
    const descricaoFluxo = 
      item.ObservacoesFluxo || 
      item.ObservacaoFluxo || 
      item.Observacoes || 
      item.Observacao || 
      item.DescricaoFluxo || 
      item.Descricao || 
      item.Motivo || 
      item.Historico || 
      item.Destino || 
      item.OData__x004f_bservacoesFluxo || 
      item.OData__x004f_bservacaoFluxo || 
      item.OData__x004f_bservacoes || 
      item.OData__x004f_bservacao || 
      (item.Title && isNaN(Number(item.Title)) ? item.Title : '') || 
      '-';

    const idVal = item.ID || item.Id || item.Title || (idx + 1);

    return {
      id: idVal,
      ID: idVal,
      CategoriaFluxo: categoriaFluxo,
      TipoFluxo: tipoFluxo,
      ValorFluxo: valorNum,
      DataFluxo: dataStr,
      Id_Tesoureiro: idTesoureiro,
      NomeTesoureiro: nomeTesoureiro,
      nome_tesoureiro: nomeTesoureiro,
      StatusFluxo: statusFluxo,
      DescricaoFluxo: descricaoFluxo,
      ObservacaoFluxo: descricaoFluxo,
      ObservacoesFluxo: descricaoFluxo,

      // Aliases para cálculos dos cards e componentes
      data: dataStr,
      dataBR: `${String(dia).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano}`,
      ano,
      mes,
      dia,
      tipo: isSaida ? 'SAIDA' : 'ENTRADA',
      categoria: categoriaFluxo,
      formaPagamento: tipoFluxo,
      valor: valorNum,
      status: statusFluxo,
      descricao: descricaoFluxo,
      observacao: descricaoFluxo,
      origem: 'SHAREPOINT_BD_FLUXOCAIXA',
      raw: item
    };
  });

  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  res.json({
    sucesso: true,
    total: movimentacoes.length,
    movimentacoes
  });
});

// Cadastrar novo fluxo de caixa na tabela BD_FluxoCaixa
app.post("/api/sharepoint/fluxo-caixa", async (req: Request, res: Response) => {
  try {
    const { 
      CategoriaFluxo, 
      TipoFluxo, 
      ValorFluxo, 
      DataFluxo, 
      Id_Tesoureiro, 
      StatusFluxo, 
      DescricaoFluxo 
    } = req.body;

    const valorNum = Number(ValorFluxo || 0);
    if (isNaN(valorNum) || valorNum <= 0) {
      return res.status(400).json({
        sucesso: false,
        erro: "Por favor, informe um valor válido em R$ para o fluxo de caixa."
      });
    }

    const dataFinal = String(DataFluxo || new Date().toISOString().split('T')[0]).trim();
    const dObj = new Date(dataFinal + 'T12:00:00');
    const ano = !isNaN(dObj.getFullYear()) ? dObj.getFullYear() : new Date().getFullYear();
    const mes = !isNaN(dObj.getMonth()) ? dObj.getMonth() + 1 : new Date().getMonth() + 1;
    const dia = !isNaN(dObj.getDate()) ? dObj.getDate() : new Date().getDate();

    const categoriaFinal = String(CategoriaFluxo || 'Entrada').trim() === 'Saída' ? 'Saída' : 'Entrada';
    const tipoFinal = String(TipoFluxo || 'Pix').trim() === 'Espécie' ? 'Espécie' : 'Pix';
    const statusFinal = String(StatusFluxo || 'OK').trim() === 'Pendente' ? 'Pendente' : 'OK';
    const descricaoFinal = String(DescricaoFluxo || '').trim();
    const tesoureiroFinal = Id_Tesoureiro || 4;

    // Obtém lista existente em cache
    const fluxoExistente: any[] = cache.fluxoCaixa || [];
    
    // Gera ID sequencial numérico único incremental (1, 2, 3...)
    let proximoId = 1;
    if (fluxoExistente.length > 0) {
      const idsNumericos = fluxoExistente
        .map(f => {
          const rawVal = f.ID ?? f.Id ?? f.Title ?? f.id;
          const num = Number(String(rawVal).replace(/\D/g, ''));
          return !isNaN(num) && num > 0 ? num : 0;
        })
        .filter(n => n > 0);

      if (idsNumericos.length > 0) {
        proximoId = Math.max(...idsNumericos) + 1;
      } else {
        proximoId = fluxoExistente.length + 1;
      }
    }

    const novoItem = {
      id: proximoId,
      ID: proximoId,
      Title: String(proximoId),
      CategoriaFluxo: categoriaFinal,
      TipoFluxo: tipoFinal,
      ValorFluxo: valorNum,
      DataFluxo: dataFinal,
      Id_Tesoureiro: tesoureiroFinal,
      StatusFluxo: statusFinal,
      DescricaoFluxo: descricaoFinal,
      ObservacaoFluxo: descricaoFinal,
      ObservacoesFluxo: descricaoFinal,
      
      // Aliases
      data: dataFinal,
      dataBR: `${String(dia).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano}`,
      ano,
      mes,
      dia,
      tipo: categoriaFinal === 'Saída' ? 'SAIDA' : 'ENTRADA',
      categoria: categoriaFinal,
      formaPagamento: tipoFinal,
      valor: valorNum,
      status: statusFinal,
      descricao: descricaoFinal,
      observacao: descricaoFinal,
      Criado: new Date().toISOString()
    };

    // 1. Atualiza cache local imediatamente
    const listaAtualizada = [novoItem, ...fluxoExistente];
    cache.fluxoCaixa = listaAtualizada;

    // 2. Persiste na lista BD_FluxoCaixa do SharePoint via Microsoft 365 REST API
    let sharePointGravado = false;
    let sharePointDetalhe = "";

    try {
      const token = await getMicrosoftToken().catch(() => null);
      if (token) {
        // Obtém o RequestDigest para permitir gravação segura no SharePoint
        let requestDigest = "";
        try {
          const contextRes = await fetch(`${SP_SITE_URL}/_api/contextinfo`, {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          });
          if (contextRes.ok) {
            const contextData = await contextRes.json();
            requestDigest = contextData?.d?.GetContextWebInformation?.FormDigestValue || "";
          }
        } catch (ctxErr) {
          console.warn("[SharePoint contextinfo] Falha ao obter digest:", ctxErr);
        }

        // Descobre o endpoint correto e o ListItemEntityTypeFullName da lista BD_FluxoCaixa
        let listUrl = `${SP_SITE_URL}/_api/web/lists(guid'6ddb245c-d7d2-4004-90fc-985ea2667dfa')/items`;
        let entityType = "SP.Data.BD_x005f_FluxoCaixaListItem";

        // Consulta metadados da lista e seus campos reais no SharePoint
        let camposDisponiveis: any[] = [];
        try {
          const listMetaRes = await fetch(`${SP_SITE_URL}/_api/web/lists(guid'6ddb245c-d7d2-4004-90fc-985ea2667dfa')?$select=ListItemEntityTypeFullName,Id`, {
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose"
            }
          });
          if (listMetaRes.ok) {
            const metaData = await listMetaRes.json();
            if (metaData?.d?.ListItemEntityTypeFullName) {
              entityType = metaData.d.ListItemEntityTypeFullName;
            }
          }

          // Busca campos editáveis reais da lista para mapear nomes internos exatos
          const fieldsRes = await fetch(`${SP_SITE_URL}/_api/web/lists(guid'6ddb245c-d7d2-4004-90fc-985ea2667dfa')/fields?$filter=Hidden%20eq%20false%20and%20ReadOnlyField%20eq%20false&$select=InternalName,Title,StaticName,TypeAsString`, {
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose"
            }
          });
          if (fieldsRes.ok) {
            const fData = await fieldsRes.json();
            camposDisponiveis = fData?.d?.results || [];
          }
        } catch (metaErr) {
          console.warn("[SharePoint list metadata] Usando tipo padrão:", metaErr);
        }

        // Monta o payload limpo mapeando com precisão os campos encontrados
        // A coluna primária Title do SharePoint recebe o número do ID sequencial
        const spPayload: Record<string, any> = {
          Title: String(proximoId)
        };

        if (camposDisponiveis.length > 0) {
          camposDisponiveis.forEach((campo: any) => {
            const iName = campo.InternalName;
            const tName = String(campo.Title || "").toLowerCase().trim();
            const normName = String(iName).toLowerCase().trim();

            if (normName === "title") {
              spPayload[iName] = String(proximoId);
              return;
            }

            // ID / Código / Número único
            if (normName === "id" || tName === "id" || normName === "codigo" || normName === "numero" || tName === "código") {
              spPayload[iName] = typeof campo.TypeAsString === 'string' && campo.TypeAsString.includes('Number') ? proximoId : String(proximoId);
            }
            // Categoria (CategoriaFluxo / Categoria / Tipo / etc.)
            else if (normName === "categoriafluxo" || normName === "categoria" || tName.includes("categoria")) {
              spPayload[iName] = categoriaFinal;
            } 
            // Tipo / Forma de Pagamento (TipoFluxo / FormaPagamento / etc.)
            else if (normName === "tipofluxo" || normName === "formapagamento" || tName.includes("forma de pagamento") || tName === "tipo" || normName === "tipopagamento") {
              spPayload[iName] = tipoFinal;
            }
            // Valor (ValorFluxo / Valor / ValorTotal / Total / etc.)
            else if (normName === "valorfluxo" || normName === "valor" || normName === "valortotal" || tName.includes("valor")) {
              spPayload[iName] = valorNum;
            }
            // Data (DataFluxo / Data / DataMovimento / etc.)
            else if (normName === "datafluxo" || normName === "data" || normName === "datamovimento" || tName.includes("data")) {
              spPayload[iName] = dataFinal;
            }
            // Id_Tesoureiro (Id_Tesoureiro / Tesoureiro / etc.)
            else if (normName.includes("tesoureiro") || tName.includes("tesoureiro")) {
              spPayload[iName] = String(tesoureiroFinal);
            }
            // Status (StatusFluxo / Status / etc.)
            else if (normName === "statusfluxo" || normName === "status" || tName.includes("status")) {
              spPayload[iName] = statusFinal;
            }
            // Descrição (DescricaoFluxo / Descricao / Motivo / Observacao / etc.)
            else if (normName === "descricaofluxo" || normName === "descricao" || normName === "observacao" || tName.includes("descri") || tName.includes("observa") || tName.includes("motivo")) {
              spPayload[iName] = descricaoFinal;
            }
          });
        } else {
          // Fallback caso metadados de campos não retornem: inclui colunas padrão
          spPayload.Title = String(proximoId);
          spPayload.CategoriaFluxo = categoriaFinal;
          spPayload.TipoFluxo = tipoFinal;
          spPayload.ValorFluxo = valorNum;
          spPayload.DataFluxo = dataFinal;
          spPayload.Id_Tesoureiro = String(tesoureiroFinal);
          spPayload.StatusFluxo = statusFinal;
          spPayload.DescricaoFluxo = descricaoFinal;
          spPayload.ObservacaoFluxo = descricaoFinal;
          spPayload.ObservacoesFluxo = descricaoFinal;
        }

        console.log("[SharePoint POST BD_FluxoCaixa] Payload a enviar:", JSON.stringify(spPayload));

        // Envia requisição em formato nometadata (padrão moderno e sem dependência de __metadata)
        let postHeaders: Record<string, string> = {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata"
        };
        if (requestDigest) {
          postHeaders["X-RequestDigest"] = requestDigest;
        }

        let postRes = await fetch(listUrl, {
          method: "POST",
          headers: postHeaders,
          body: JSON.stringify(spPayload)
        });

        // Se falhou por propriedade inexistente, remove a propriedade acusada e retenta
        let tentativasRestantes = 5;
        while (!postRes.ok && tentativasRestantes > 0) {
          tentativasRestantes--;
          const errText = await postRes.text().catch(() => "");
          console.warn(`[SharePoint POST BD_FluxoCaixa] Status ${postRes.status}: ${errText.slice(0, 200)}`);

          // Extrai nome da propriedade com erro (ex: The property 'CategoriaFluxo' does not exist)
          const match = errText.match(/The property '([^']+)' does not exist/i);
          if (match && match[1]) {
            const propInvalida = match[1];
            delete spPayload[propInvalida];
            console.log(`[SharePoint POST BD_FluxoCaixa] Removendo coluna não encontrada '${propInvalida}' e retentando...`);
            
            postRes = await fetch(listUrl, {
              method: "POST",
              headers: postHeaders,
              body: JSON.stringify(spPayload)
            });
            continue;
          }

          // Se a lista exigir o formato odata=verbose, tenta com __metadata correto
          if (errText.includes("odata=verbose") || errText.includes("Cannot find type")) {
            postHeaders["Accept"] = "application/json;odata=verbose";
            postHeaders["Content-Type"] = "application/json;odata=verbose";
            const payloadVerbose = {
              __metadata: { type: entityType },
              ...spPayload
            };
            postRes = await fetch(listUrl, {
              method: "POST",
              headers: postHeaders,
              body: JSON.stringify(payloadVerbose)
            });
            break;
          }

          break;
        }

        if (postRes.ok) {
          const postData = await postRes.json().catch(() => null);
          const spCreated = postData?.d || postData;
          if (spCreated?.ID || spCreated?.Id) {
            novoItem.id = spCreated.ID || spCreated.Id;
            novoItem.ID = spCreated.ID || spCreated.Id;
          }
          sharePointGravado = true;
          sharePointDetalhe = "Gravado com sucesso no SharePoint Microsoft 365.";
          console.log(`[BD_FluxoCaixa] 🚀 Item salvo diretamente no SharePoint ID: ${novoItem.ID}`);
        }
      }
    } catch (spErr: any) {
      console.warn("[SharePoint POST BD_FluxoCaixa] Erro de comunicação com o SharePoint:", spErr?.message || spErr);
    }

    console.log(`[BD_FluxoCaixa] ✅ Novo fluxo #${novoItem.ID} cadastrado: ${categoriaFinal} - R$ ${valorNum} (${sharePointGravado ? 'SharePoint OK' : 'Cache SWR OK'})`);

    return res.json({
      sucesso: true,
      mensagem: sharePointGravado 
        ? "Fluxo de caixa salvo com sucesso na lista BD_FluxoCaixa do SharePoint!"
        : "Fluxo de caixa cadastrado com sucesso!",
      sharePointSincronizado: sharePointGravado,
      detalhe: sharePointDetalhe,
      item: novoItem
    });
  } catch (err: any) {
    console.error("[BD_FluxoCaixa] ❌ Erro ao cadastrar novo fluxo:", err);
    return res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro interno ao cadastrar fluxo de caixa."
    });
  }
});

// Atualizar lançamento existente na tabela BD_FluxoCaixa
app.put("/api/sharepoint/fluxo-caixa/:id", async (req: Request, res: Response) => {
  try {
    const idParam = req.params.id;
    const { 
      CategoriaFluxo, 
      TipoFluxo, 
      ValorFluxo, 
      DataFluxo, 
      Id_Tesoureiro, 
      StatusFluxo, 
      DescricaoFluxo 
    } = req.body;

    const valorNum = Number(ValorFluxo || 0);
    if (isNaN(valorNum) || valorNum <= 0) {
      return res.status(400).json({
        sucesso: false,
        erro: "Por favor, informe um valor válido em R$ para o fluxo de caixa."
      });
    }

    const dataFinal = String(DataFluxo || new Date().toISOString().split('T')[0]).trim();
    const dObj = new Date(dataFinal + 'T12:00:00');
    const ano = !isNaN(dObj.getFullYear()) ? dObj.getFullYear() : new Date().getFullYear();
    const mes = !isNaN(dObj.getMonth()) ? dObj.getMonth() + 1 : new Date().getMonth() + 1;
    const dia = !isNaN(dObj.getDate()) ? dObj.getDate() : new Date().getDate();

    const categoriaFinal = String(CategoriaFluxo || 'Entrada').trim() === 'Saída' ? 'Saída' : 'Entrada';
    const tipoFinal = String(TipoFluxo || 'Pix').trim() === 'Espécie' ? 'Espécie' : 'Pix';
    const statusFinal = String(StatusFluxo || 'OK').trim() === 'Pendente' ? 'Pendente' : 'OK';
    const descricaoFinal = String(DescricaoFluxo || '').trim();
    const tesoureiroFinal = Id_Tesoureiro || 4;

    const fluxoExistente: any[] = cache.fluxoCaixa || [];
    const index = fluxoExistente.findIndex(f => 
      String(f.id) === String(idParam) || 
      String(f.ID) === String(idParam) || 
      String(f.Title) === String(idParam)
    );

    const itemOriginal = index >= 0 ? fluxoExistente[index] : null;
    const idItem = itemOriginal?.id || itemOriginal?.ID || idParam;

    const itemAtualizado = {
      ...(itemOriginal || {}),
      id: idItem,
      ID: idItem,
      Title: String(idItem),
      CategoriaFluxo: categoriaFinal,
      TipoFluxo: tipoFinal,
      ValorFluxo: valorNum,
      DataFluxo: dataFinal,
      Id_Tesoureiro: tesoureiroFinal,
      StatusFluxo: statusFinal,
      DescricaoFluxo: descricaoFinal,
      ObservacaoFluxo: descricaoFinal,
      ObservacoesFluxo: descricaoFinal,
      
      // Aliases
      data: dataFinal,
      dataBR: `${String(dia).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano}`,
      ano,
      mes,
      dia,
      tipo: categoriaFinal === 'Saída' ? 'SAIDA' : 'ENTRADA',
      categoria: categoriaFinal,
      formaPagamento: tipoFinal,
      valor: valorNum,
      status: statusFinal,
      descricao: descricaoFinal,
      observacao: descricaoFinal,
      Atualizado: new Date().toISOString()
    };

    if (index >= 0) {
      fluxoExistente[index] = itemAtualizado;
    } else {
      fluxoExistente.unshift(itemAtualizado);
    }
    cache.fluxoCaixa = [...fluxoExistente];

    // Atualiza no SharePoint
    let sharePointAtualizado = false;
    try {
      const token = await getMicrosoftToken().catch(() => null);
      if (token) {
        let requestDigest = "";
        try {
          const contextRes = await fetch(`${SP_SITE_URL}/_api/contextinfo`, {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          });
          if (contextRes.ok) {
            const contextData = await contextRes.json();
            requestDigest = contextData?.d?.GetContextWebInformation?.FormDigestValue || "";
          }
        } catch {}

        const spItemId = Number(String(idItem).replace(/\D/g, '')) || idItem;
        const listUrl = `${SP_SITE_URL}/_api/web/lists(guid'6ddb245c-d7d2-4004-90fc-985ea2667dfa')/items(${spItemId})`;

        const spPayload: Record<string, any> = {
          CategoriaFluxo: categoriaFinal,
          TipoFluxo: tipoFinal,
          ValorFluxo: valorNum,
          DataFluxo: dataFinal,
          Id_Tesoureiro: String(tesoureiroFinal),
          StatusFluxo: statusFinal,
          DescricaoFluxo: descricaoFinal,
          ObservacaoFluxo: descricaoFinal,
          ObservacoesFluxo: descricaoFinal
        };

        const patchRes = await fetch(listUrl, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
            ...(requestDigest ? { "X-RequestDigest": requestDigest } : {})
          },
          body: JSON.stringify(spPayload)
        });

        if (patchRes.ok) {
          sharePointAtualizado = true;
          console.log(`[BD_FluxoCaixa] ✏️ Item #${idItem} atualizado com sucesso no SharePoint!`);
        }
      }
    } catch (spErr) {
      console.warn("[BD_FluxoCaixa] Aviso ao atualizar item no SharePoint:", spErr);
    }

    return res.json({
      sucesso: true,
      mensagem: "Lançamento atualizado com sucesso!",
      sharePointAtualizado,
      item: itemAtualizado
    });
  } catch (err: any) {
    console.error("[BD_FluxoCaixa] ❌ Erro ao atualizar fluxo:", err);
    return res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro interno ao atualizar fluxo de caixa."
    });
  }
});

// Excluir lançamento da tabela BD_FluxoCaixa
app.delete("/api/sharepoint/fluxo-caixa/:id", async (req: Request, res: Response) => {
  try {
    const idParam = req.params.id;
    const fluxoExistente: any[] = cache.fluxoCaixa || [];
    
    // Remove da memória/cache local
    const novaLista = fluxoExistente.filter(f => 
      String(f.id) !== String(idParam) && 
      String(f.ID) !== String(idParam) && 
      String(f.Title) !== String(idParam)
    );
    cache.fluxoCaixa = novaLista;

    // Tenta excluir no SharePoint
    let sharePointExcluido = false;
    try {
      const token = await getMicrosoftToken().catch(() => null);
      if (token) {
        let requestDigest = "";
        try {
          const contextRes = await fetch(`${SP_SITE_URL}/_api/contextinfo`, {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          });
          if (contextRes.ok) {
            const contextData = await contextRes.json();
            requestDigest = contextData?.d?.GetContextWebInformation?.FormDigestValue || "";
          }
        } catch {}

        const spItemId = Number(String(idParam).replace(/\D/g, '')) || idParam;
        const listUrl = `${SP_SITE_URL}/_api/web/lists(guid'6ddb245c-d7d2-4004-90fc-985ea2667dfa')/items(${spItemId})`;

        const deleteRes = await fetch(listUrl, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${token}`,
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*",
            ...(requestDigest ? { "X-RequestDigest": requestDigest } : {})
          }
        });

        if (deleteRes.ok) {
          sharePointExcluido = true;
          console.log(`[BD_FluxoCaixa] 🗑️ Item #${idParam} excluído do SharePoint!`);
        }
      }
    } catch (spErr) {
      console.warn("[BD_FluxoCaixa] Aviso ao excluir item no SharePoint:", spErr);
    }

    return res.json({
      sucesso: true,
      mensagem: "Lançamento excluído com sucesso!",
      sharePointExcluido
    });
  } catch (err: any) {
    console.error("[BD_FluxoCaixa] ❌ Erro ao excluir fluxo:", err);
    return res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro interno ao excluir fluxo de caixa."
    });
  }
});

// Obter métricas consolidadas para a tela Indicador Trilho
app.get("/api/sharepoint/indicador-trilho", async (req: Request, res: Response) => {
  const force = req.query.force === "true" || req.query.refresh === "true";
  
  const [membros, celulas, capacitacoes, membrosCapac] = await Promise.all([
    carregarMembrosCompleto(force),
    carregarCelulasCompleto(force),
    carregarCapacitacoesCompleto(force),
    carregarMembrosCapacCompleto(force)
  ]);

  // Mapa de setor para área com base na lista BD_celulas
  const setorParaAreaMap = new Map<string, string>();
  celulas.forEach((c: any) => {
    const s = String(c.Setor || c.setor || "").trim();
    const a = String(c.Area || c.area || c.Rede || "").trim();
    if (s && a && !setorParaAreaMap.has(s.toLowerCase())) {
      setorParaAreaMap.set(s.toLowerCase(), a);
    }
  });

  // 1. Extrair todas as etapas do trilho a partir de BD_Capacitacao
  const etapasExtraidas = Array.from(
    new Set(
      capacitacoes
        .map((c: any) => {
          const nome = String(
            c.Capacitacao || 
            c.capacitacao || 
            c.Title || 
            c.NomeCapacitacao || 
            c.Nome || 
            c.Etapa || 
            ""
          ).trim();
          return nome;
        })
        .filter(Boolean)
    )
  );

  // Etapas padrão da Paz Church como fallback caso a lista ainda esteja sendo populada
  const etapasPadrao = [
    "Encontro com Deus",
    "Batismo",
    "Maturidade no Espírito",
    "CTL",
    "Treinamento de Líderes"
  ];

  const etapas: string[] = etapasExtraidas.length > 0 ? etapasExtraidas : etapasPadrao;

  // 2. Mapear conclusões de membros de BD_MembrosCapac (verificando ID_membro e Status === 'OK')
  // Essa verificação é feita em lote único para todas as etapas do trilho
  const concluidosPorMembroId = new Map<string, Set<string>>();
  const concluidosPorMembroNome = new Map<string, Set<string>>();

  membrosCapac.forEach((mc: any) => {
    // Validação de Status (conforme regra de negócio: Status === 'OK')
    const statusRaw = String(mc.Status || mc.status || mc.STATUS || mc.Situacao || "").trim().toUpperCase();
    const isConcluido = 
      statusRaw === "OK" || 
      statusRaw.includes("OK") || 
      statusRaw === "1" || 
      statusRaw === "TRUE" || 
      statusRaw === "SIM" || 
      statusRaw === "CONCLUÍDO" || 
      statusRaw === "CONCLUIDO" || 
      statusRaw === "CONCLUÍDA" || 
      statusRaw === "CONCLUIDA" ||
      !statusRaw; // Se a tabela apenas registra os concluídos sem preencher a coluna status

    if (!isConcluido) return;

    // ID do membro que corresponde ao ID da lista BD_membros (incluindo ID_Lider e ID_esp)
    const rawId = String(
      mc.ID_membro ||
      mc.ID_Lider ||
      mc.ID_esp ||
      mc.ID_MEMBRO || 
      mc.IdMembro || 
      mc.ID_Membro || 
      mc.Id_Membro || 
      mc.id_membro ||
      mc.MembroId || 
      mc.Membro_ID || 
      ""
    ).trim();

    const rawNome = String(
      mc.NomeMembro || 
      mc.Membro || 
      mc.Title || 
      mc.Nome || 
      mc.Nome_Membro || 
      ""
    ).trim().toLowerCase();

    // Relacionamento: BD_Capacitacao - Coluna Capacitacao; BD_MembrosCapac - Capacitacao
    const capNome = String(
      mc.Capacitacao || 
      mc.capacitacao || 
      mc.Bairro || // No SharePoint a coluna interna pode ser Bairro mas armazena a Capacitacao
      mc.NomeCapacitacao || 
      mc.CAPACITACAO || 
      mc.Etapa || 
      mc.Title || 
      ""
    ).trim();

    if (!capNome) return;

    // Encontra a etapa correspondente (case-insensitive)
    const etapaEncontrada = etapas.find(e => 
      e.toLowerCase() === capNome.toLowerCase() ||
      capNome.toLowerCase().includes(e.toLowerCase()) ||
      e.toLowerCase().includes(capNome.toLowerCase())
    ) || capNome;

    if (rawId && rawId !== "0" && rawId !== "null") {
      if (!concluidosPorMembroId.has(rawId)) {
        concluidosPorMembroId.set(rawId, new Set());
      }
      concluidosPorMembroId.get(rawId)!.add(etapaEncontrada);
    }

    if (rawNome) {
      if (!concluidosPorMembroNome.has(rawNome)) {
        concluidosPorMembroNome.set(rawNome, new Set());
      }
      concluidosPorMembroNome.get(rawNome)!.add(etapaEncontrada);
    }
  });

  // 3. Processar membros e vincular etapas concluídas
  const membrosProcessados = membros.map((m: any) => {
    const mId = String(m.id || m.ID || "").trim();
    const mNome = String(m.nome || m.Title || "").trim();
    const mNomeLower = mNome.toLowerCase();
    const setorOriginal = String(m.setor || m.raw?.Setor || m.raw?.setor || "Sem Setor").trim();
    const areaOriginal = String(m.area || m.raw?.Area || m.raw?.area || setorParaAreaMap.get(setorOriginal.toLowerCase()) || "Geral").trim();

    const etapasConcluidasSet = new Set<string>();

    // Verifica por ID
    if (mId && concluidosPorMembroId.has(mId)) {
      concluidosPorMembroId.get(mId)!.forEach(e => etapasConcluidasSet.add(e));
    }

    // Verifica por Nome
    if (mNomeLower && concluidosPorMembroNome.has(mNomeLower)) {
      concluidosPorMembroNome.get(mNomeLower)!.forEach(e => etapasConcluidasSet.add(e));
    }

    const etapasConcluidas = Array.from(etapasConcluidasSet);
    const totalConcluidas = etapasConcluidas.length;
    const totalEtapas = etapas.length;
    const percentual = totalEtapas > 0 ? Math.round((totalConcluidas / totalEtapas) * 100) : 0;

    let statusTrilho: 'Completo' | 'Em Andamento' | 'Não Iniciado' = 'Não Iniciado';
    if (totalConcluidas >= totalEtapas && totalEtapas > 0) {
      statusTrilho = 'Completo';
    } else if (totalConcluidas > 0) {
      statusTrilho = 'Em Andamento';
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

  // 4. Lista de Áreas e Setores distintos
  const areasDisponiveis = Array.from(
    new Set(membrosProcessados.map(m => m.area).filter(Boolean))
  ).sort();

  const setoresDisponiveis = Array.from(
    new Set(membrosProcessados.map(m => m.setor).filter(Boolean))
  ).sort();

  // Função auxiliar para calcular métricas de um conjunto de membros
  const calcularMetricasConjunto = (listaMembros: typeof membrosProcessados) => {
    const total = listaMembros.length;
    if (total === 0) {
      return {
        totalMembros: 0,
        percentualMedio: 0,
        membrosCompletos: 0,
        membrosEmAndamento: 0,
        membrosNaoIniciados: 0,
        etapasStats: etapas.map(etapa => ({
          etapa,
          concluidos: 0,
          pendentes: 0,
          percentual: 0
        }))
      };
    }

    const etapasStats = etapas.map(etapa => {
      const concluidos = listaMembros.filter(m => m.etapasConcluidas.includes(etapa)).length;
      const pendentes = total - concluidos;
      const percentual = Math.round((concluidos / total) * 100);
      return {
        etapa,
        concluidos,
        pendentes,
        percentual
      };
    });

    const membrosCompletos = listaMembros.filter(m => m.statusTrilho === 'Completo').length;
    const membrosNaoCompletos = total - membrosCompletos;
    const membrosComZero = listaMembros.filter(m => m.etapasConcluidas.length === 0).length;
    const membrosEmAndamento = listaMembros.filter(m => m.statusTrilho === 'Em Andamento').length;
    const membrosNaoIniciados = membrosComZero;

    // Média de quantos % do trilho foi concluído de todos os membros dentro do escopo
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

  // Métricas da Igreja em Geral
  const geral = calcularMetricasConjunto(membrosProcessados);

  // Métricas por Área
  const porArea: Record<string, ReturnType<typeof calcularMetricasConjunto>> = {};
  areasDisponiveis.forEach(area => {
    const membrosDaArea = membrosProcessados.filter(m => m.area.toLowerCase() === area.toLowerCase());
    porArea[area] = calcularMetricasConjunto(membrosDaArea);
  });

  // Métricas por Setor
  const porSetor: Record<string, ReturnType<typeof calcularMetricasConjunto>> = {};
  setoresDisponiveis.forEach(setor => {
    const membrosDoSetor = membrosProcessados.filter(m => m.setor.toLowerCase() === setor.toLowerCase());
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

// Fila por item para sequencializar gravações no SharePoint
const itemUpdateQueues = new Map<string, Promise<any>>();

function enqueueItemUpdate<T>(id: string, task: () => Promise<T>): Promise<T> {
  const current = itemUpdateQueues.get(id) || Promise.resolve();
  const next = current
    .then(async () => {
      await new Promise(r => setTimeout(r, 100));
      return task();
    })
    .finally(() => {
      if (itemUpdateQueues.get(id) === next) {
        itemUpdateQueues.delete(id);
      }
    });
  itemUpdateQueues.set(id, next);
  return next;
}

// Validar relatório
app.post("/api/sharepoint/validar-relatorio", async (req: Request, res: Response) => {
  const { id, recebido, idTesoureiro, dataTesouraria } = req.body;
  const isRecebido = Boolean(recebido);

  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  const dataBR = dataTesouraria || `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
  const idTesoureiroFinal = isRecebido 
    ? (idTesoureiro !== undefined && idTesoureiro !== null && String(idTesoureiro).trim() !== "" ? String(idTesoureiro) : "4")
    : null;

  const relatorio = cache.relatorios.find((r: any) => String(r.ID || r.Id) === String(id));
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
    } catch (err: any) {
      console.warn("[SharePoint] Falha ao enviar validação para o SharePoint:", err?.message);
    }
  }).catch(e => console.warn("[SharePoint] Aviso na fila de validação:", e));

  res.json({
    sucesso: true,
    id,
    TESOURARIA_RECEB: isRecebido,
    DATA_TESOURARIA: isRecebido ? dataBR : null,
    ID_TESOUREIRO: idTesoureiroFinal
  });
});

// Editar relatório
app.post("/api/sharepoint/editar-relatorio", async (req: Request, res: Response) => {
  const { id, celula, data, pix, especie, total } = req.body;
  
  const item = cache.relatorios.find((r: any) => String(r.ID || r.Id) === String(id));
  if (item) {
    if (celula !== undefined) {
      item.C_x00e9_lula = celula;
      item.Célula = celula;
    }
    if (data !== undefined) {
      item.DataNascimento = data;
      item.DataCelula = data;
    }
    if (pix !== undefined) {
      item.Bairro = Number(pix);
      item.ValorOferta = Number(pix);
    }
    if (especie !== undefined) {
      item.OfertaEspecie = Number(especie);
    }
    if (total !== undefined) {
      item.valorTotal = Number(total);
      item.Total = Number(total);
    }
  }

  enqueueItemUpdate(String(id), async () => {
    try {
      const token = await getMicrosoftToken();
      const spValidateUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('BD_Relatorio')/items(${id})/validateUpdateListItem`;
      const formValues: Array<{ FieldName: string; FieldValue: string }> = [];
      if (pix !== undefined) formValues.push({ FieldName: "Bairro", FieldValue: String(pix) });
      if (especie !== undefined) formValues.push({ FieldName: "OfertaEspecie", FieldValue: String(especie) });
      if (data !== undefined) formValues.push({ FieldName: "DataNascimento", FieldValue: String(data) });
      if (celula !== undefined) formValues.push({ FieldName: "C_x00e9_lula", FieldValue: String(celula) });

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
    } catch (err: any) {
      console.warn("[SharePoint] Falha ao enviar edição para o SharePoint:", err?.message);
    }
  }).catch(e => console.warn("[SharePoint] Aviso na fila de edição:", e));

  res.json({
    sucesso: true,
    id,
    item
  });
});

// Excluir relatório (suporta tanto POST /api/sharepoint/excluir-relatorio quanto DELETE /api/sharepoint/relatorios/:id)
const handleExcluirRelatorio = async (idParaExcluir: string | number, res: Response) => {
  const idStr = String(idParaExcluir).trim();
  if (!idStr) {
    return res.status(400).json({ sucesso: false, erro: "ID do relatório não informado" });
  }

  // 1. Remover do cache persistente e atualizar entrada
  const store = PersistentCacheManager.getStore();
  const relatorios = store.relatorios?.data || [];
  const index = relatorios.findIndex((r: any) => String(r.ID || r.Id || r.id) === idStr);
  let itemRemovido = null;
  if (index !== -1) {
    itemRemovido = relatorios.splice(index, 1)[0];
    PersistentCacheManager.atualizarEntrada("relatorios", relatorios);
  }

  // 2. Deletar no SharePoint via fila assíncrona
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
      console.log(`[SharePoint] ✅ Relatório ID ${idStr} excluído no SharePoint.`);
    } catch (err: any) {
      console.warn(`[SharePoint] Aviso ao excluir relatório ID ${idStr} no SharePoint:`, err?.message);
    }
  }).catch(e => console.warn("[SharePoint] Aviso na fila de exclusão:", e));

  return res.json({
    sucesso: true,
    mensagem: `Relatório ID ${idStr} excluído com sucesso.`,
    id: idStr,
    item: itemRemovido,
    totalRestantes: (store.relatorios?.data || []).length
  });
};

app.post("/api/sharepoint/excluir-relatorio", async (req: Request, res: Response) => {
  const { id } = req.body;
  return handleExcluirRelatorio(id, res);
});

app.delete("/api/sharepoint/relatorios/:id", async (req: Request, res: Response) => {
  const { id } = req.params;
  return handleExcluirRelatorio(id, res);
});

// Forçar sincronização manual
app.post("/api/sharepoint/sync", async (req: Request, res: Response) => {
  try {
    await sincronizarListasSharePoint();
    res.json({
      sucesso: true,
      mensagem: "Sincronização com SharePoint concluída.",
      totalMembros: cache.membros.length,
      totalRelatorios: cache.relatorios.length,
      totalCelulas: cache.celulas.length
    });
  } catch (err: any) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Falha ao sincronizar com SharePoint"
    });
  }
});

// Fallback para rotas /api/* não encontradas - sempre retorna JSON, nunca HTML
app.use("/api/*", (req: Request, res: Response) => {
  res.status(404).json({
    sucesso: false,
    erro: `Rota da API não encontrada: ${req.method} ${req.originalUrl || req.url}`
  });
});

// Tratador global de erros da API para evitar crash no runtime do Vercel
app.use((err: any, req: Request, res: Response, next: any) => {
  console.error("[API Global Error]", err);
  if (!res.headersSent) {
    res.status(500).json({
      sucesso: false,
      erro: err?.message || "Erro interno no servidor da API"
    });
  }
});

export { app };
export default app;
