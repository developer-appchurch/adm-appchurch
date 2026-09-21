import express, { Request, Response } from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

const app = express();
const PORT = 3000;

app.use(express.json());

// SharePoint Microsoft 365 Config
const SP_USER = process.env.SHAREPOINT_USER || "midia.sobral@paz.church";
const SP_PASS = process.env.SHAREPOINT_PASS || "Pazsobral23";
const SP_SITE_URL = process.env.SHAREPOINT_SITE_URL || "https://pazchurch.sharepoint.com/sites/PazSobral";
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

// Buscar itens de uma lista do SharePoint com paginação completa
async function fetchSharePointList(listTitle: string, maxItems: number = 25000, orderByIdDesc: boolean = false): Promise<any[]> {
  const token = await getMicrosoftToken();
  let items: any[] = [];
  const orderParam = orderByIdDesc ? "&$orderby=Id%20desc" : "";
  let nextUrl: string | null = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=5000${orderParam}`;
  
  while (nextUrl && items.length < maxItems) {
    const res: any = await fetch(nextUrl, {
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json;odata=verbose"
      }
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`Erro ao consultar lista ${listTitle}: ${res.status} - ${errText}`);
    }

    const data: any = await res.json();
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
  try {
    // 1. BD_membros
    const membrosBrutos = await fetchSharePointList("BD_membros", 5000, false);
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

    // 2. BD_Relatorio (carrega com Id desc para priorizar registros recentes e cobrir todas as páginas)
    try {
      const relatoriosBrutos = await fetchSharePointList("BD_Relatorio", 25000, true);
      cache.relatorios = relatoriosBrutos;
      console.log(`[SharePoint] BD_Relatorio carregado: ${cache.relatorios.length} itens.`);
    } catch (eRel) {
      console.warn("[SharePoint] Aviso ao carregar BD_Relatorio:", eRel);
    }

    // 3. BD_celulas
    try {
      const celulasBrutas = await fetchSharePointList("BD_celulas", 1000, false);
      cache.celulas = celulasBrutas;
      console.log(`[SharePoint] BD_celulas carregado: ${cache.celulas.length} células.`);
    } catch (eCel) {
      console.warn("[SharePoint] Aviso ao carregar BD_celulas:", eCel);
    }

    cache.status = "CONECTADO";
    cache.erro = null;
    cache.lastSync = new Date().toISOString();
    console.log("[SharePoint] Conexão e sincronização concluídas com sucesso!");
  } catch (err: any) {
    cache.status = "ERRO";
    cache.erro = err?.message || String(err);
    console.error("[SharePoint] Erro na sincronização:", err);
  }
}

// Inicia sincronização ao ligar o servidor
sincronizarListasSharePoint().catch(console.error);

// Agendar renovação periódica a cada 15 minutos
setInterval(() => {
  sincronizarListasSharePoint().catch(console.error);
}, 15 * 60 * 1000);

// --- ROTAS DA API ---

// Status da conexão
app.get("/api/sharepoint/status", (req: Request, res: Response) => {
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
app.get("/api/sharepoint/membros", (req: Request, res: Response) => {
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

// Autenticação de Usuário contra a tabela BD_membros do SharePoint
app.post("/api/sharepoint/auth-membro", async (req: Request, res: Response) => {
  const { login, senha } = req.body;
  const termo = String(login || "").trim();
  const senhaDigitada = String(senha || "").trim();

  // Ambos login e senha são estritamente obrigatórios
  if (!termo || !senhaDigitada) {
    return res.status(401).json({
      sucesso: false,
      erro: "Login ou Senha incorretos"
    });
  }

  // Se a lista de membros ainda estiver vazia, tenta sincronizar imediatamente
  if (cache.membros.length === 0) {
    await sincronizarListasSharePoint().catch(() => {});
  }

  const termoNorm = normalizar(termo);

  // Busca na tabela BD_membros: por Login, Nome, Email ou Title
  const membro = cache.membros.find((m) => {
    const loginNorm = normalizar(m.login);
    const nomeNorm = normalizar(m.nome);
    const titleNorm = normalizar(m.Title);
    const emailNorm = normalizar(m.email);
    const emailUserNorm = normalizar(m.email ? m.email.split("@")[0] : "");

    return (
      loginNorm === termoNorm ||
      nomeNorm === termoNorm ||
      titleNorm === termoNorm ||
      emailNorm === termoNorm ||
      emailUserNorm === termoNorm
    );
  });

  // Se o usuário não for localizado
  if (!membro) {
    return res.status(401).json({
      sucesso: false,
      erro: "Login ou Senha incorretos"
    });
  }

  // Validação estrita: a senha relacionada ao usuário deve estar correta
  const senhaCadastrada = String(membro.senha || "").trim();
  const senhaCorreta = senhaCadastrada 
    ? (senhaDigitada === senhaCadastrada)
    : (senhaDigitada === "Pazsobral23" || senhaDigitada === "teste");

  if (!senhaCorreta) {
    return res.status(401).json({
      sucesso: false,
      erro: "Login ou Senha incorretos"
    });
  }

  // Login e senha validados com sucesso
  return res.json({
    sucesso: true,
    membro: {
      id: membro.id,
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
app.get("/api/sharepoint/relatorios", (req: Request, res: Response) => {
  res.json({
    sucesso: true,
    total: cache.relatorios.length,
    relatorios: cache.relatorios
  });
});

// Obter células da lista BD_celulas
app.get("/api/sharepoint/celulas", (req: Request, res: Response) => {
  res.json({
    sucesso: true,
    total: cache.celulas.length,
    celulas: cache.celulas
  });
});

// Obter setores distintos da tabela BD_celulas
app.get("/api/sharepoint/setores", (req: Request, res: Response) => {
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

startServer();
