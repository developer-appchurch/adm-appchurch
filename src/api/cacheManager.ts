import fs from 'fs';
import path from 'path';

// Configuração de TTL (Time To Live)
export const FRESH_TTL_MS = 5 * 60 * 1000; // 5 minutos (fresco - 0 chamadas de rede)
export const STALE_TTL_MS = 24 * 60 * 60 * 1000; // 24 horas (stale - serve instantâneo e revalida em background)
export const TOKEN_EXPIRY_BUFFER_MS = 2 * 60 * 1000; // 2 minutos de margem de segurança para o token

interface CacheEntry<T = any> {
  data: T;
  updatedAt: number;
  expiresAt?: number;
  itemCount?: number;
}

interface CacheStore {
  token: { access_token: string; expiresAt: number } | null;
  membros: CacheEntry<any[]> | null;
  relatorios: CacheEntry<any[]> | null;
  celulas: CacheEntry<any[]> | null;
  capacitacoes: CacheEntry<any[]> | null;
  membrosCapac: CacheEntry<any[]> | null;
  fluxoCaixa: CacheEntry<any[]> | null;
  status: "CONECTADO" | "CONECTANDO" | "ERRO";
  erro: string | null;
  lastSync: string | null;
  metrics: {
    hits: number;
    staleHits: number;
    misses: number;
    backgroundRevalidations: number;
    lastRevalidationAt: string | null;
  };
}

// Arquivos persistentes no diretório /tmp (disponível no ambiente Vercel Serverless / AWS Lambda)
const TMP_DIR = process.platform === 'win32' ? path.join(process.cwd(), '.tmp') : '/tmp';
const CACHE_FILE_PATH = path.join(TMP_DIR, 'appchurch_sp_cache.json');
const TOKEN_FILE_PATH = path.join(TMP_DIR, 'appchurch_sp_token.json');
const CREDS_FILE_PATH = path.join(TMP_DIR, 'appchurch_sp_credentials.json');

export interface StoredCredentials {
  user: string;
  pass: string;
  siteUrl?: string;
  clientId?: string;
  updatedAt?: string;
}

// Garante que o diretório /tmp exista
try {
  if (!fs.existsSync(TMP_DIR)) {
    fs.mkdirSync(TMP_DIR, { recursive: true });
  }
} catch {}

// Variável global no runtime Node.js que sobrevive a múltiplas invocações do container Serverless
declare global {
  // eslint-disable-next-line no-var
  var __SP_GLOBAL_STORE__: CacheStore | undefined;
}

function inicializarStore(): CacheStore {
  if (globalThis.__SP_GLOBAL_STORE__) {
    return globalThis.__SP_GLOBAL_STORE__;
  }

  // Tenta restaurar do arquivo de cache em /tmp
  let storeRestaurada: Partial<CacheStore> | null = null;
  try {
    if (fs.existsSync(CACHE_FILE_PATH)) {
      const conteudo = fs.readFileSync(CACHE_FILE_PATH, 'utf-8');
      storeRestaurada = JSON.parse(conteudo);
      console.log('[CacheManager] ✅ Cache persistente restaurado com sucesso de /tmp!');
    }
  } catch (e) {
    console.warn('[CacheManager] Aviso ao ler cache persistente de /tmp:', e);
  }

  const novaStore: CacheStore = {
    token: storeRestaurada?.token || null,
    membros: storeRestaurada?.membros || null,
    relatorios: storeRestaurada?.relatorios || null,
    celulas: storeRestaurada?.celulas || null,
    capacitacoes: storeRestaurada?.capacitacoes || null,
    membrosCapac: storeRestaurada?.membrosCapac || null,
    fluxoCaixa: storeRestaurada?.fluxoCaixa || null,
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

export class PersistentCacheManager {
  private static revalidacoesEmAndamento = new Map<string, Promise<any>>();

  public static getStore(): CacheStore {
    return inicializarStore();
  }

  /**
   * Salva o estado atual do cache em /tmp para persistência entre invocações frias/quentes
   */
  public static async salvarEmDisco(): Promise<void> {
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
      await fs.promises.writeFile(CACHE_FILE_PATH, payload, 'utf-8');
    } catch (err) {
      console.warn('[CacheManager] Não foi possível persistir cache em disco (/tmp):', err);
    }
  }

  /**
   * Obtém token Microsoft do cache (em memória ou /tmp)
   */
  public static getMicrosoftToken(): string | null {
    const store = this.getStore();
    const now = Date.now();
    
    if (store.token && store.token.expiresAt > now + TOKEN_EXPIRY_BUFFER_MS) {
      return store.token.access_token;
    }

    // Tenta ler do arquivo de token separado
    try {
      if (fs.existsSync(TOKEN_FILE_PATH)) {
        const raw = fs.readFileSync(TOKEN_FILE_PATH, 'utf-8');
        const tokenObj = JSON.parse(raw);
        if (tokenObj && tokenObj.access_token && tokenObj.expiresAt > now + TOKEN_EXPIRY_BUFFER_MS) {
          store.token = tokenObj;
          return tokenObj.access_token;
        }
      }
    } catch {}

    return null;
  }

  /**
   * Salva o token Microsoft no cache global e em /tmp
   */
  public static async setMicrosoftToken(token: string, expiresInSeconds: number): Promise<void> {
    const store = this.getStore();
    const expiresAt = Date.now() + (Number(expiresInSeconds || 3600) * 1000);
    const tokenObj = { access_token: token, expiresAt };
    
    store.token = tokenObj;

    try {
      await fs.promises.writeFile(TOKEN_FILE_PATH, JSON.stringify(tokenObj), 'utf-8');
    } catch {}
  }

  /**
   * Obtém credenciais Microsoft salvas no disco (/tmp) para persistência em ambiente serverless
   */
  public static getSavedCredentials(): StoredCredentials | null {
    try {
      if (fs.existsSync(CREDS_FILE_PATH)) {
        const raw = fs.readFileSync(CREDS_FILE_PATH, 'utf-8');
        const creds = JSON.parse(raw);
        if (creds && creds.user && creds.pass) {
          return creds;
        }
      }
    } catch {}
    return null;
  }

  /**
   * Salva credenciais customizadas Microsoft no disco (/tmp)
   */
  public static async saveCredentials(creds: StoredCredentials): Promise<void> {
    try {
      await fs.promises.writeFile(CREDS_FILE_PATH, JSON.stringify(creds), 'utf-8');
    } catch (e) {
      console.warn('[CacheManager] Erro ao persistir credenciais em /tmp:', e);
    }
  }

  /**
   * Executa a estratégia Stale-While-Revalidate para listas do SharePoint
   */
  public static async getWithSWR<T extends any[]>(
    chave: 'membros' | 'relatorios' | 'celulas' | 'capacitacoes' | 'membrosCapac' | 'fluxoCaixa',
    revalidador: () => Promise<T>,
    options?: {
      freshTtlMs?: number;
      staleTtlMs?: number;
      forceRefresh?: boolean;
    }
  ): Promise<{
    data: T;
    cacheStatus: 'HIT' | 'STALE' | 'MISS';
    source: 'memory' | 'disk' | 'network';
    itemCount: number;
    updatedAt: number;
  }> {
    const store = this.getStore();
    const freshTtl = options?.freshTtlMs ?? FRESH_TTL_MS;
    const staleTtl = options?.staleTtlMs ?? STALE_TTL_MS;
    const now = Date.now();

    const entrada = store[chave] as CacheEntry<T> | null;

    // Se forceRefresh não foi solicitado e temos uma entrada válida
    if (!options?.forceRefresh && entrada && Array.isArray(entrada.data) && entrada.data.length > 0) {
      const idade = now - entrada.updatedAt;

      // 1. FRESH HIT (Menos de freshTtl): Retorna imediatamente sem nenhuma chamada à Microsoft
      if (idade < freshTtl) {
        store.metrics.hits++;
        return {
          data: entrada.data,
          cacheStatus: 'HIT',
          source: 'memory',
          itemCount: entrada.data.length,
          updatedAt: entrada.updatedAt
        };
      }

      // 2. STALE HIT (Entre freshTtl e staleTtl): Retorna imediatamente ao usuário e revalida em background!
      if (idade < staleTtl) {
        store.metrics.staleHits++;
        
        // Dispara revalidação em background se não houver outra em andamento
        this.dispararRevalidacaoEmBackground(chave, revalidador);

        return {
          data: entrada.data,
          cacheStatus: 'STALE',
          source: 'memory',
          itemCount: entrada.data.length,
          updatedAt: entrada.updatedAt
        };
      }
    }

    // 3. CACHE MISS (Primeiro acesso ou cache totalmente expirado): Executa a busca e popula o cache
    store.metrics.misses++;
    console.log(`[CacheManager] 🌐 [${chave}] Cache Miss ou Forçado - buscando da API da Microsoft...`);

    const novosDados = await revalidador();
    this.atualizarEntrada(chave, novosDados);

    return {
      data: novosDados,
      cacheStatus: 'MISS',
      source: 'network',
      itemCount: novosDados.length,
      updatedAt: Date.now()
    };
  }

  /**
   * Dispara a busca em background sem travar a resposta HTTP do usuário
   */
  private static dispararRevalidacaoEmBackground<T extends any[]>(
    chave: 'membros' | 'relatorios' | 'celulas' | 'capacitacoes' | 'membrosCapac' | 'fluxoCaixa',
    revalidador: () => Promise<T>
  ): void {
    if (this.revalidacoesEmAndamento.has(chave)) {
      return; // Já existe uma revalidação rodando em segundo plano
    }

    console.log(`[CacheManager] 🔄 [${chave}] SWR: Iniciando revalidação assíncrona em background...`);
    const store = this.getStore();
    store.metrics.backgroundRevalidations++;

    const promessa = (async () => {
      try {
        const novosDados = await revalidador();
        if (Array.isArray(novosDados) && novosDados.length > 0) {
          this.atualizarEntrada(chave, novosDados);
          store.metrics.lastRevalidationAt = new Date().toISOString();
          console.log(`[CacheManager] ⚡ [${chave}] Revalidação em background concluída com sucesso! (${novosDados.length} itens)`);
        }
      } catch (err) {
        console.warn(`[CacheManager] ⚠️ [${chave}] Aviso na revalidação em background:`, err);
      } finally {
        this.revalidacoesEmAndamento.delete(chave);
      }
    })();

    this.revalidacoesEmAndamento.set(chave, promessa);
  }

  /**
   * Atualiza a entrada de cache e salva em disco
   */
  public static atualizarEntrada<T extends any[]>(
    chave: 'membros' | 'relatorios' | 'celulas' | 'capacitacoes' | 'membrosCapac' | 'fluxoCaixa',
    dados: T
  ): void {
    const store = this.getStore();
    store[chave] = {
      data: dados,
      updatedAt: Date.now(),
      itemCount: dados.length
    };
    store.lastSync = new Date().toISOString();
    store.status = "CONECTADO";
    store.erro = null;

    // Salva em disco de forma não bloqueante
    this.salvarEmDisco().catch(() => {});
  }

  /**
   * Limpa todo o cache
   */
  public static limparCache(): void {
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
    } catch {}
  }

  /**
   * Retorna estatísticas completas para a tela de diagnóstico e logs
   */
  public static getEstatisticas() {
    const store = this.getStore();
    const now = Date.now();

    const formatarIdade = (ts?: number) => {
      if (!ts) return null;
      const seg = Math.floor((now - ts) / 1000);
      if (seg < 60) return `${seg}s atrás`;
      const min = Math.floor(seg / 60);
      return `${min}m ${seg % 60}s atrás`;
    };

    return {
      statusGeral: store.status,
      ultimoSync: store.lastSync,
      tokenValido: !!(store.token && store.token.expiresAt > now),
      tokenExpiraEmSegundos: store.token ? Math.max(0, Math.floor((store.token.expiresAt - now) / 1000)) : 0,
      metricas: store.metrics,
      revalidacoesAtivas: Array.from(this.revalidacoesEmAndamento.keys()),
      listas: {
        membros: {
          carregado: !!store.membros,
          total: store.membros?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.membros?.updatedAt),
          isStale: store.membros ? (now - store.membros.updatedAt > FRESH_TTL_MS) : true
        },
        relatorios: {
          carregado: !!store.relatorios,
          total: store.relatorios?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.relatorios?.updatedAt),
          isStale: store.relatorios ? (now - store.relatorios.updatedAt > FRESH_TTL_MS) : true
        },
        celulas: {
          carregado: !!store.celulas,
          total: store.celulas?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.celulas?.updatedAt),
          isStale: store.celulas ? (now - store.celulas.updatedAt > FRESH_TTL_MS) : true
        },
        capacitacoes: {
          carregado: !!store.capacitacoes,
          total: store.capacitacoes?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.capacitacoes?.updatedAt),
          isStale: store.capacitacoes ? (now - store.capacitacoes.updatedAt > FRESH_TTL_MS) : true
        },
        membrosCapac: {
          carregado: !!store.membrosCapac,
          total: store.membrosCapac?.itemCount ?? 0,
          idadeTexto: formatarIdade(store.membrosCapac?.updatedAt),
          isStale: store.membrosCapac ? (now - store.membrosCapac.updatedAt > FRESH_TTL_MS) : true
        }
      },
      configuracoes: {
        freshTtlMinutos: FRESH_TTL_MS / 60000,
        staleTtlHoras: STALE_TTL_MS / 3600000,
        caminhoCacheDisco: CACHE_FILE_PATH
      }
    };
  }
}
