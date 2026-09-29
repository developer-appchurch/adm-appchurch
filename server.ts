import express, { Request, Response } from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

const app = express();
const PORT = 3000;

app.use(express.json());

// SharePoint Microsoft 365 Config
const SP_USER = process.env.SHAREPOINT_USER || "midia.sobral@paz.church";
const SP_PASS = process.env.SHAREPOINT_PASS || "Pazsobral23";

function sanitizeSharePointSiteUrl(rawUrl?: string): string {
  let url = String(rawUrl || "https://pazchurch.sharepoint.com/sites/PazSobral").trim();
  // Remove /Lists/ ou /Lists no final se vier de variáveis de ambiente
  url = url.replace(/\/Lists\/?$/i, "");
  // Remove barras extras no final
  url = url.replace(/\/+$/, "");
  return url || "https://pazchurch.sharepoint.com/sites/PazSobral";
}

const SP_SITE_URL = sanitizeSharePointSiteUrl(process.env.SHAREPOINT_SITE_URL);
const SP_RESOURCE = "https://pazchurch.sharepoint.com";
const MS_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office Public Client

interface CachedData {
  token: string | null;
  tokenExpiresAt: number;
  membros: any[];
  relatorios: any[];
  celulas: any[];
  lastSync: string | null;
  status: "CONECTADO" | "CONECTANDO" | "ERRO";
  erro: string | null;
}

const cache: CachedData = {
  token: null,
  tokenExpiresAt: 0,
  membros: [],
  relatorios: [],
  celulas: [],
  lastSync: null,
  status: "CONECTANDO",
  erro: null
};

// Obter token OAuth da Microsoft para SharePoint
async function getMicrosoftToken(): Promise<string> {
  const now = Date.now();
  if (cache.token && cache.tokenExpiresAt > now + 60000) {
    return cache.token;
  }

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
    const errText = await res.text();
    throw new Error(`Falha na autenticação Microsoft: ${res.status} - ${errText}`);
  }

  const json: any = await res.json();
  cache.token = json.access_token;
  // Expiração com margem de segurança
  cache.tokenExpiresAt = Date.now() + (Number(json.expires_in || 3600) * 1000);
  return cache.token!;
}

// Mapeamento de GUIDs canônicos das listas do SharePoint (imunes a erros 404 e alterações de título)
const KNOWN_LIST_GUIDS: Record<string, string> = {
  "BD_membros": "0f0da17f-880f-4391-9521-1e41636cfecc",
  "BD_Relatorio": "82228774-77ec-4574-9b57-19272c6910af",
  "BD_celulas": "dfa7d45a-9023-4c35-a7f3-1c976360ffe0",
  "BD_PerfilPermissao": "4bafa1ae-bb82-4084-9aff-dd8ec5a6a8ab",
  "BD_Bairros": "bd3a1a3f-b775-421d-8269-1d7acdcacfba"
};

// Obter URL base segura para consulta de itens da lista (prioriza GUID para evitar 404)
async function getSharePointListUrl(listTitle: string, token: string, orderByIdDesc: boolean = false): Promise<string> {
  const orderParam = orderByIdDesc ? "&$orderby=Id%20desc" : "";
  const guid = KNOWN_LIST_GUIDS[listTitle];

  if (guid) {
    return `${SP_SITE_URL}/_api/web/lists(guid'${guid}')/items?$top=5000${orderParam}`;
  }

  return `${SP_SITE_URL}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000${orderParam}`;
}

// Buscar itens de uma lista do SharePoint com paginação completa e fallback automático
async function fetchSharePointList(listTitle: string, maxItems: number = 25000, orderByIdDesc: boolean = false): Promise<any[]> {
  const token = await getMicrosoftToken();
  let items: any[] = [];
  
  let nextUrl: string | null = await getSharePointListUrl(listTitle, token, orderByIdDesc);
  let attemptFallback = true;
  
  while (nextUrl && items.length < maxItems) {
    // Normaliza URL caso venha relativa da paginação do SharePoint
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
      // Se retornar 404 na primeira tentativa por título, tenta descobrir o GUID dinamicamente
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
              continue; // Re-executa com o GUID encontrado
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

// Normaliza string para comparação sem acento e case-insensitive
function normalizar(texto: any): string {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

// Sincronização automática de todas as listas
async function sincronizarListasSharePoint(): Promise<void> {
  console.log("[SharePoint] Iniciando sincronização automática com Microsoft 365...");
  cache.status = "CONECTANDO";
  let erros: string[] = [];

  // 1. BD_membros
  try {
    const membrosBrutos = await fetchSharePointList("BD_membros", 5000, false);
    if (Array.isArray(membrosBrutos) && membrosBrutos.length > 0) {
      cache.membros = membrosBrutos.map((m: any) => ({
        id: m.ID || m.Id,
        ID: m.ID || m.Id,
        Title: m.Title || "",
        nome: m.Nome || m.NomeCompleto || m.Title || "Membro",
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
      console.log(`[SharePoint] BD_membros carregado: ${cache.membros.length} membros.`);
    }
  } catch (eMem: any) {
    console.warn("[SharePoint] Erro ao carregar BD_membros:", eMem?.message || eMem);
    erros.push(`BD_membros: ${eMem?.message || eMem}`);
  }

  // 2. BD_Relatorio (carrega com Id desc para priorizar registros recentes e cobrir todas as páginas)
  try {
    const relatoriosBrutos = await fetchSharePointList("BD_Relatorio", 25000, true);
    if (Array.isArray(relatoriosBrutos) && relatoriosBrutos.length > 0) {
      cache.relatorios = relatoriosBrutos;
      console.log(`[SharePoint] BD_Relatorio carregado: ${cache.relatorios.length} itens.`);
    }
  } catch (eRel: any) {
    console.warn("[SharePoint] Erro ao carregar BD_Relatorio:", eRel?.message || eRel);
    erros.push(`BD_Relatorio: ${eRel?.message || eRel}`);
  }

  // 3. BD_celulas
  try {
    const celulasBrutas = await fetchSharePointList("BD_celulas", 1000, false);
    if (Array.isArray(celulasBrutas) && celulasBrutas.length > 0) {
      cache.celulas = celulasBrutas;
      console.log(`[SharePoint] BD_celulas carregado: ${cache.celulas.length} células.`);
    }
  } catch (eCel: any) {
    console.warn("[SharePoint] Erro ao carregar BD_celulas:", eCel?.message || eCel);
    erros.push(`BD_celulas: ${eCel?.message || eCel}`);
  }

  if (cache.membros.length > 0 || cache.relatorios.length > 0 || cache.celulas.length > 0) {
    cache.status = "CONECTADO";
    cache.erro = erros.length > 0 ? erros.join("; ") : null;
    cache.lastSync = new Date().toISOString();
    console.log("[SharePoint] Conexão e sincronização ativas com sucesso!");
  } else {
    cache.status = "ERRO";
    cache.erro = erros.join("; ") || "Nenhuma lista pôde ser consultada";
    console.error("[SharePoint] Falha ao sincronizar listas:", cache.erro);
  }
}

// Inicia sincronização ao ligar o servidor (se não estiver em ambiente serverless)
if (!process.env.VERCEL) {
  sincronizarListasSharePoint().catch(console.error);

  // Agendar renovação periódica a cada 15 minutos em servidores dedicados
  setInterval(() => {
    sincronizarListasSharePoint().catch(console.error);
  }, 15 * 60 * 1000);
}

// Helper para garantir dados sincronizados em cold starts (Vercel Serverless)
async function garantirDadosSincronizados(): Promise<void> {
  if (cache.membros.length === 0 || cache.relatorios.length === 0 || cache.celulas.length === 0) {
    await sincronizarListasSharePoint().catch((err) => {
      console.warn("[SharePoint] Aviso na sincronização sob demanda:", err?.message || err);
    });
  }
}

// --- ROTAS DA API ---

// Status da conexão
app.get("/api/sharepoint/status", async (req: Request, res: Response) => {
  if (cache.membros.length === 0 && cache.relatorios.length === 0) {
    await garantirDadosSincronizados();
  }

  res.json({
    status: cache.status,
    conta: SP_USER,
    siteUrl: SP_SITE_URL,
    totalMembros: cache.membros.length,
    totalRelatorios: cache.relatorios.length,
    totalCelulas: cache.celulas.length,
    ultimoSync: cache.lastSync,
    erro: cache.erro
  });
});

// Listar membros sincronizados
app.get("/api/sharepoint/membros", async (req: Request, res: Response) => {
  if (cache.membros.length === 0) {
    await garantirDadosSincronizados();
  }

  res.json({
    sucesso: true,
    total: cache.membros.length,
    membros: cache.membros.map(m => ({
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

// Autenticação de Usuário contra a tabela BD_membros do SharePoint + Master Accounts
app.post("/api/sharepoint/auth-membro", async (req: Request, res: Response) => {
  const { login, senha } = req.body;
  const termo = String(login || "").trim();
  const senhaDigitada = String(senha || "").trim();

  // Ambos login e senha são estritamente obrigatórios
  if (!termo || !senhaDigitada) {
    return res.status(401).json({
      sucesso: false,
      erro: "Por favor, preencha o login e a senha."
    });
  }

  const termoNorm = normalizar(termo);
  const senhaDigitadaNorm = senhaDigitada.toLowerCase();

  // Senhas padrão aceitas pela igreja para contingência / master
  const senhasValidasPadrao = ["pazsobral23", "12345", "123456", "admin", "teste", "pazsobral", "sobral23", "admin123"];

  // 1. Contas Master / Administrativas (Developer, Mídia SharePoint, Admin Geral, Tesouraria)
  const isMasterDeveloper = 
    termoNorm === "developer.appchurch@gmail.com" || 
    termoNorm === "developer.appchurch" || 
    termoNorm === "developer" ||
    termoNorm === "appchurch";

  const isMasterMidia = 
    termoNorm === "midia.sobral@paz.church" || 
    termoNorm === "midia.sobral" || 
    termoNorm === "midia";

  const isMasterAdmin = 
    termoNorm === "admin" || 
    termoNorm === "tesouraria" || 
    termoNorm === "adm" ||
    termoNorm === "pazchurch";

  if (isMasterDeveloper || isMasterMidia || isMasterAdmin) {
    if (senhasValidasPadrao.includes(senhaDigitadaNorm) || senhaDigitada === SP_PASS || senhaDigitada === "12345") {
      console.log(`[Auth] Login administrativo bem-sucedido: ${termo}`);
      return res.json({
        sucesso: true,
        membro: {
          id: 4,
          ID: 4,
          nome: isMasterDeveloper ? "Developer AppChurch" : (isMasterMidia ? "Mídia Paz Church" : "Junio Fonteles"),
          login: termo,
          email: isMasterDeveloper ? "developer.appchurch@gmail.com" : (isMasterMidia ? "midia.sobral@paz.church" : "tesouraria@pazchurch.com"),
          cargo: "Tesouraria",
          celula: "Central",
          setor: "Safira",
          area: "Área Central",
          telefone: "(88) 99999-0000",
          status: "Ativo"
        }
      });
    }
  }

  // 2. Se a lista de membros ainda estiver vazia, tenta sincronizar imediatamente
  if (cache.membros.length === 0) {
    await sincronizarListasSharePoint().catch(() => {});
  }

  // 3. Busca na tabela BD_membros: por Login, Nome, Email, Title ou ID
  let membro = cache.membros.find((m) => {
    const loginNorm = normalizar(m.login);
    const nomeNorm = normalizar(m.nome);
    const titleNorm = normalizar(m.Title);
    const emailNorm = normalizar(m.email);
    const emailUserNorm = normalizar(m.email ? m.email.split("@")[0] : "");
    const idNorm = String(m.id || m.ID || "");

    return (
      loginNorm === termoNorm ||
      nomeNorm === termoNorm ||
      titleNorm === termoNorm ||
      emailNorm === termoNorm ||
      emailUserNorm === termoNorm ||
      idNorm === termoNorm
    );
  });

  // Se não encontrou no cache mas o cache estiver vazio, tenta sincronizar uma vez
  if (!membro && cache.membros.length === 0) {
    try {
      const direct = await fetchSharePointList("BD_membros", 5000, false);
      if (Array.isArray(direct) && direct.length > 0) {
        cache.membros = direct.map((m: any) => ({
          id: m.ID || m.Id,
          ID: m.ID || m.Id,
          Title: m.Title || "",
          nome: m.Nome || m.NomeCompleto || m.Title || "Membro",
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

        membro = cache.membros.find((m) => {
          const loginNorm = normalizar(m.login);
          const nomeNorm = normalizar(m.nome);
          const titleNorm = normalizar(m.Title);
          const emailNorm = normalizar(m.email);
          return (
            loginNorm === termoNorm ||
            nomeNorm === termoNorm ||
            titleNorm === termoNorm ||
            emailNorm === termoNorm
          );
        });
      }
    } catch (e: any) {
      console.warn("[Auth] Aviso ao sincronizar membros sob demanda:", e?.message || e);
    }
  }

  // Se o usuário não for localizado
  if (!membro) {
    // Se o termo for "jfonteles" ou "junio fonteles" (conhecido líder/tesoureiro)
    if (termoNorm.includes("fonteles") || termoNorm === "jfonteles" || termoNorm === "junio") {
      if (senhasValidasPadrao.includes(senhaDigitadaNorm) || senhaDigitada === "12345") {
        return res.json({
          sucesso: true,
          membro: {
            id: 4,
            ID: 4,
            nome: "Junio Fonteles",
            login: "Jfonteles",
            email: "juniosina@hotmail.com",
            cargo: "Líder de Setor",
            celula: "Central",
            setor: "Safira",
            area: "Área Central",
            telefone: "(88) 99999-0004",
            status: "Ativo"
          }
        });
      }
    }

    return res.status(401).json({
      sucesso: false,
      erro: "Login não encontrado no cadastro do SharePoint."
    });
  }

  // Validação da senha:
  // 1) Se a senha cadastrada no SharePoint bater (case-insensitive ou exata)
  // 2) OU se for uma das senhas padrão do sistema (12345, 123456, Pazsobral23, teste, etc.)
  // 3) Se o membro não tiver senha cadastrada no SharePoint, permite com as senhas padrão
  const senhaCadastrada = String(membro.senha || "").trim();
  const senhaCorreta = 
    (senhaCadastrada && (senhaDigitada === senhaCadastrada || senhaDigitadaNorm === senhaCadastrada.toLowerCase())) ||
    senhasValidasPadrao.includes(senhaDigitadaNorm) ||
    senhaDigitada === SP_PASS;

  if (!senhaCorreta) {
    return res.status(401).json({
      sucesso: false,
      erro: "Senha incorreta para este usuário."
    });
  }

  // Login e senha validados com sucesso
  console.log(`[Auth] Usuário autenticado com sucesso: ${membro.nome} (${membro.login})`);
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

// Obter relatórios da lista BD_Relatorio
app.get("/api/sharepoint/relatorios", async (req: Request, res: Response) => {
  if (cache.relatorios.length === 0) {
    await garantirDadosSincronizados();
  }

  res.json({
    sucesso: true,
    total: cache.relatorios.length,
    relatorios: cache.relatorios
  });
});

// Obter células da lista BD_celulas
app.get("/api/sharepoint/celulas", async (req: Request, res: Response) => {
  if (cache.celulas.length === 0) {
    await garantirDadosSincronizados();
  }

  res.json({
    sucesso: true,
    total: cache.celulas.length,
    celulas: cache.celulas
  });
});

// Obter setores distintos da tabela BD_celulas
app.get("/api/sharepoint/setores", async (req: Request, res: Response) => {
  if (cache.celulas.length === 0) {
    await garantirDadosSincronizados();
  }

  const setoresUnicos = Array.from(
    new Set(
      cache.celulas
        .map((c: any) => String(c.Setor || c.setor || "").trim())
        .filter(Boolean)
    )
  ).sort();

  res.json({
    sucesso: true,
    total: setoresUnicos.length,
    setores: setoresUnicos
  });
});

// Fila por item para sequencializar gravações e evitar colisões de concorrência (SPException) no SharePoint
const itemUpdateQueues = new Map<string, Promise<any>>();

function enqueueItemUpdate<T>(id: string, task: () => Promise<T>): Promise<T> {
  const current = itemUpdateQueues.get(id) || Promise.resolve();
  const next = current
    .then(async () => {
      // Espaçamento de 100ms para liberação de bloqueio do item no SharePoint
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

// Validar ou atualizar status TESOURARIA_RECEB de um relatório
app.post("/api/sharepoint/validar-relatorio", async (req: Request, res: Response) => {
  const { id, recebido, idTesoureiro, dataTesouraria } = req.body;
  const isRecebido = Boolean(recebido);

  // Formatar data em dd/MM/yyyy conforme padrão pt-BR do SharePoint
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  const dataBR = dataTesouraria || `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
  
  // ID do tesoureiro logado (ex: Junio Fonteles - ID: 4)
  const idTesoureiroFinal = isRecebido 
    ? (idTesoureiro !== undefined && idTesoureiro !== null && String(idTesoureiro).trim() !== "" ? String(idTesoureiro) : "4")
    : null;

  // 1. Atualizar em memória cache local imediatamente
  const relatorio = cache.relatorios.find((r: any) => String(r.ID || r.Id) === String(id));
  if (relatorio) {
    relatorio.TESOURARIA_RECEB = isRecebido;
    relatorio.DATA_TESOURARIA = isRecebido ? dataBR : null;
    relatorio.ID_TESOUREIRO = idTesoureiroFinal;
  }

  // 2. Persistir alteração no SharePoint via validateUpdateListItem em fila por item
  enqueueItemUpdate(String(id), async () => {
    try {
      const token = await getMicrosoftToken();
      const spValidateUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('BD_Relatorio')/items(${id})/validateUpdateListItem`;

      const formValues = [
        { FieldName: "TESOURARIA_RECEB", FieldValue: isRecebido ? "1" : "0" },
        { FieldName: "DATA_TESOURARIA", FieldValue: isRecebido ? dataBR : "" },
        { FieldName: "ID_TESOUREIRO", FieldValue: isRecebido ? String(idTesoureiroFinal || "4") : "" }
      ];

      const spRes = await fetch(spValidateUrl, {
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

      if (spRes.ok) {
        const json = await spRes.json().catch(() => null);
        const results = json?.d?.ValidateUpdateListItem?.results || [];
        const hasError = results.some((r: any) => r.HasException);
        if (!hasError) {
          console.log(`[SharePoint] Relatório ID ${id} validado com sucesso via validateUpdateListItem (TESOURARIA_RECEB=${isRecebido})`);
          return;
        }
        console.warn(`[SharePoint] Erro nos campos ao validar ID ${id}:`, results);
      } else {
        const txt = await spRes.text().catch(() => "");
        console.warn(`[SharePoint] Resposta ao validar ID ${id}: ${spRes.status} ${txt.slice(0, 150)}`);
      }
    } catch (err: any) {
      console.warn("[SharePoint] Falha ao enviar atualização para o SharePoint:", err?.message);
    }
  }).catch(e => console.warn("[SharePoint] Aviso na fila de atualização:", e));

  res.json({
    sucesso: true,
    id,
    TESOURARIA_RECEB: isRecebido,
    DATA_TESOURARIA: isRecebido ? dataBR : null,
    ID_TESOUREIRO: idTesoureiroFinal
  });
});

// Atualizar dados de um relatório (Célula, Data, PIX, Espécie, etc.)
app.post("/api/sharepoint/editar-relatorio", async (req: Request, res: Response) => {
  const { id, celula, data, pix, especie, total } = req.body;
  
  // 1. Atualiza no cache em memória
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

  // 2. Persistir no SharePoint via validateUpdateListItem em fila por item
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
        const spRes = await fetch(spValidateUrl, {
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

        if (spRes.ok) {
          console.log(`[SharePoint] Relatório ID ${id} editado com sucesso via validateUpdateListItem`);
        } else {
          const txt = await spRes.text().catch(() => "");
          console.warn(`[SharePoint] Resposta ao editar ID ${id}: ${spRes.status} ${txt.slice(0, 150)}`);
        }
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

// --- VITE MIDDLEWARE & SERVING ---
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req: Request, res: Response) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`[Server] Servidor rodando em http://0.0.0.0:${PORT}`);
    console.log(`[Server] SharePoint configurado: ${SP_USER} -> ${SP_SITE_URL}`);
  });
}

// Inicia servidor somente se não estiver rodando no Vercel Serverless
if (!process.env.VERCEL) {
  startServer();
}

export { app };
export default app;
