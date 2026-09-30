import React, { useState, useEffect } from 'react';
import { 
  Activity, 
  RotateCw, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  Server, 
  Globe, 
  Database, 
  ShieldCheck, 
  Copy, 
  Check, 
  Terminal, 
  Clock, 
  Layers,
  ChevronDown,
  ChevronUp,
  Cpu,
  KeyRound,
  ExternalLink,
  Zap,
  Trash2,
  HardDrive
} from 'lucide-react';
import { SharePointService } from '../../services/sharepointService';

interface DiagnosticCheck {
  id: string;
  name: string;
  endpoint: string;
  method: string;
  guid?: string;
  status: number;
  statusText?: string;
  durationMs: number;
  success: boolean;
  itemCount?: number;
  details?: any;
  error?: string;
}

interface DiagnosticResponse {
  timestamp: string;
  environment: {
    isVercel: boolean;
    runtime: string;
    nodeVersion: string;
    platform: string;
    vercelRegion: string;
    sharepointUserConfigured: boolean;
    sharepointUserMasked: string;
    sharepointPassConfigured: boolean;
    sharepointPassLength: number;
    siteUrl: string;
    clientId: string;
  };
  checks: DiagnosticCheck[];
  overallStatus: 'SUCCESS' | 'WARNING' | 'FAILED';
  totalDurationMs: number;
}

interface CacheStatusResponse {
  sucesso: boolean;
  cache: {
    statusGeral: string;
    ultimoSync: string | null;
    tokenValido: boolean;
    tokenExpiraEmSegundos: number;
    metricas: {
      hits: number;
      staleHits: number;
      misses: number;
      backgroundRevalidations: number;
      lastRevalidationAt: string | null;
    };
    revalidacoesAtivas: string[];
    listas: {
      membros: { carregado: boolean; total: number; idadeTexto: string | null; isStale: boolean };
      relatorios: { carregado: boolean; total: number; idadeTexto: string | null; isStale: boolean };
      celulas: { carregado: boolean; total: number; idadeTexto: string | null; isStale: boolean };
    };
    configuracoes: {
      freshTtlMinutos: number;
      staleTtlHoras: number;
      caminhoCacheDisco: string;
    };
  };
}

export const DiagnosticsView: React.FC = () => {
  const [data, setData] = useState<DiagnosticResponse | null>(null);
  const [cacheInfo, setCacheInfo] = useState<CacheStatusResponse['cache'] | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [cacheActionLoading, setCacheActionLoading] = useState<boolean>(false);
  const [copiado, setCopiado] = useState<boolean>(false);
  const [expandedCheck, setExpandedCheck] = useState<string | null>(null);

  // Testador de login de membro individual
  const [testLogin, setTestLogin] = useState<string>('Jfonteles');
  const [testSenha, setTestSenha] = useState<string>('');
  const [testResultado, setTestResultado] = useState<any>(null);
  const [testLoading, setTestLoading] = useState<boolean>(false);

  const carregarStatusCache = async () => {
    try {
      const res = await fetch('/api/sharepoint/cache/status');
      if (res.ok) {
        const json: CacheStatusResponse = await res.json();
        if (json.sucesso) {
          setCacheInfo(json.cache);
        }
      }
    } catch {}
  };

  const executarDiagnostico = async () => {
    setLoading(true);
    try {
      const [diagRes] = await Promise.all([
        fetch('/api/sharepoint/diagnostics'),
        carregarStatusCache()
      ]);

      const rawText = await diagRes.text();
      let json: DiagnosticResponse | null = null;
      try {
        json = JSON.parse(rawText);
      } catch {
        throw new Error(`Resposta não é JSON (HTTP ${diagRes.status}): ${rawText.slice(0, 150)}`);
      }

      if (json) {
        setData(json);
      } else {
        throw new Error(`Resposta vazia da API (HTTP ${diagRes.status})`);
      }
    } catch (err: any) {
      console.error('[DiagnosticsView] Erro ao obter diagnóstico:', err);
      setData({
        timestamp: new Date().toISOString(),
        environment: {
          isVercel: false,
          runtime: 'Erro de Comunicação Frontend -> Backend',
          nodeVersion: 'Desconhecido',
          platform: 'Desconhecido',
          vercelRegion: 'N/A',
          sharepointUserConfigured: false,
          sharepointUserMasked: 'N/A',
          sharepointPassConfigured: false,
          sharepointPassLength: 0,
          siteUrl: 'https://pazchurch.sharepoint.com/sites/PazSobral',
          clientId: 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
        },
        checks: [
          {
            id: 'network_fail',
            name: 'Comunicação com API Backend',
            endpoint: '/api/sharepoint/diagnostics',
            method: 'GET',
            status: 0,
            durationMs: 0,
            success: false,
            error: err?.message || 'Falha de rede ao chamar a API de diagnóstico'
          }
        ],
        overallStatus: 'FAILED',
        totalDurationMs: 0
      });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    executarDiagnostico();
  }, []);

  const forcarRevalidacaoCache = async () => {
    setCacheActionLoading(true);
    try {
      await fetch('/api/sharepoint/cache/refresh', { method: 'POST' });
      await carregarStatusCache();
      await executarDiagnostico();
    } finally {
      setCacheActionLoading(false);
    }
  };

  const limparCachePersistente = async () => {
    if (!window.confirm('Tem certeza que deseja limpar todo o cache persistente de listas e token?')) return;
    setCacheActionLoading(true);
    try {
      await fetch('/api/sharepoint/cache/clear', { method: 'POST' });
      await carregarStatusCache();
      await executarDiagnostico();
    } finally {
      setCacheActionLoading(false);
    }
  };

  const copiarRelatorio = () => {
    if (!data) return;
    const relatorio = `### RELATÓRIO DE DIAGNÓSTICO & CACHE SHAREPOINT (${new Date(data.timestamp).toLocaleString('pt-BR')})
**Ambiente:** ${data.environment.runtime} (${data.environment.vercelRegion})
**Status Geral:** ${data.overallStatus} (${data.totalDurationMs}ms)
**Conta:** ${data.environment.sharepointUserMasked}
**Site URL:** ${data.environment.siteUrl}

#### ESTATÍSTICAS DE CACHE PERSISTENTE (SWR):
- Cache Hits (0ms): ${cacheInfo?.metricas?.hits ?? 0}
- Stale Hits (Revalidação em Background): ${cacheInfo?.metricas?.staleHits ?? 0}
- Cache Misses (Chamadas Microsoft): ${cacheInfo?.metricas?.misses ?? 0}
- Revalidações em Background Concluídas: ${cacheInfo?.metricas?.backgroundRevalidations ?? 0}
- Token Microsoft Válido: ${cacheInfo?.tokenValido ? `Sim (restam ${cacheInfo?.tokenExpiraEmSegundos}s)` : 'Não'}

#### ENDPOINTS VERIFICADOS:
${data.checks.map(c => `- **${c.name}** [${c.status || 'ERR'}]: ${c.success ? '✅ SUCESSO' : '❌ FALHA'} (${c.durationMs}ms)
  Endpoint: ${c.method} ${c.endpoint}
  ${c.error ? `Erro: ${c.error}` : `Detalhes: ${c.details || (c.itemCount !== undefined ? `${c.itemCount} itens` : 'OK')}`}
`).join('\n')}
`;
    navigator.clipboard.writeText(relatorio);
    setCopiado(true);
    setTimeout(() => setCopiado(false), 3000);
  };

  const handleTestarLoginMembro = async (e: React.FormEvent) => {
    e.preventDefault();
    setTestLoading(true);
    setTestResultado(null);
    try {
      const spService = SharePointService.getInstance();
      const res = await spService.consultarMembroSharePoint(testLogin.trim(), testSenha.trim());
      setTestResultado(res);
    } catch (err: any) {
      setTestResultado({
        sucesso: false,
        erro: err?.message || 'Erro inesperado ao consultar membro.'
      });
    } finally {
      setTestLoading(false);
    }
  };

  return (
    <div id="diagnostics-view-container" className="p-3 sm:p-6 space-y-6 max-w-7xl mx-auto text-slate-100">
      
      {/* Top Header */}
      <div className="flex flex-wrap items-center justify-between gap-4 border-b border-[#2d334d] pb-4">
        <div>
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-indigo-950/80 border border-indigo-500/40 text-indigo-400">
              <Activity className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-white flex items-center gap-2">
                <span>Diagnóstico & Cache Persistente (SWR)</span>
                {data && (
                  <span className={`text-xs px-2.5 py-0.5 rounded-full font-bold uppercase tracking-wider border ${
                    data.overallStatus === 'SUCCESS'
                      ? 'bg-emerald-950 text-emerald-400 border-emerald-500/40'
                      : data.overallStatus === 'WARNING'
                      ? 'bg-amber-950 text-amber-400 border-amber-500/40'
                      : 'bg-rose-950 text-rose-400 border-rose-500/40'
                  }`}>
                    {data.overallStatus === 'SUCCESS' ? '100% Operacional' : data.overallStatus === 'WARNING' ? 'Atenção / Parcial' : 'Falha Detectada'}
                  </span>
                )}
              </h1>
              <p className="text-xs text-slate-400 mt-0.5">
                Stale-While-Revalidate e persistência global para minimizar chamadas à API da Microsoft
              </p>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={copiarRelatorio}
            disabled={!data}
            className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-[#202538] hover:bg-[#2b324b] text-slate-200 hover:text-white border border-[#343d5c] text-xs font-semibold transition-all cursor-pointer disabled:opacity-50"
            title="Copiar relatório formatado para o clipboard"
          >
            {copiado ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4 text-slate-400" />}
            <span>{copiado ? 'Copiado!' : 'Copiar Relatório'}</span>
          </button>

          <button
            onClick={executarDiagnostico}
            disabled={loading}
            className="flex items-center gap-2 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold shadow-md shadow-indigo-600/30 transition-all cursor-pointer disabled:opacity-60"
          >
            <RotateCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
            <span>{loading ? 'Executando...' : 'Reexecutar Diagnóstico'}</span>
          </button>
        </div>
      </div>

      {/* Painel Especial de Cache Persistente & SWR */}
      {cacheInfo && (
        <div className="bg-[#141724] border border-[#2c334f] rounded-2xl p-5 shadow-xl space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#252b42] pb-3">
            <div className="flex items-center gap-2">
              <div className="p-1.5 rounded-lg bg-amber-500/20 text-amber-400">
                <Zap className="w-4 h-4" />
              </div>
              <div>
                <h2 className="text-sm font-bold text-white flex items-center gap-2">
                  Sistema de Cache Persistente • Stale-While-Revalidate (SWR)
                </h2>
                <p className="text-[11px] text-slate-400">
                  Cache global em memória + disco (/tmp) com revalidação assíncrona automática
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={forcarRevalidacaoCache}
                disabled={cacheActionLoading}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-600/80 hover:bg-indigo-600 text-white text-xs font-semibold transition-all cursor-pointer disabled:opacity-50"
                title="Força uma nova leitura completa no SharePoint"
              >
                <RotateCw className={`w-3.5 h-3.5 ${cacheActionLoading ? 'animate-spin' : ''}`} />
                <span>Revalidar Cache (SWR)</span>
              </button>

              <button
                onClick={limparCachePersistente}
                disabled={cacheActionLoading}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-red-950/80 hover:bg-red-900 border border-red-500/30 text-red-200 text-xs font-semibold transition-all cursor-pointer disabled:opacity-50"
                title="Limpa cache de memória e arquivos /tmp"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Limpar Cache</span>
              </button>
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="bg-[#191d2e] p-3 rounded-xl border border-[#2b324d]">
              <span className="text-[11px] text-slate-400 font-medium block">Cache Hits (0ms)</span>
              <span className="text-lg font-bold text-emerald-400 font-mono">
                {cacheInfo.metricas.hits}
              </span>
              <span className="text-[10px] text-slate-500 block mt-0.5">Servido direto da memória</span>
            </div>

            <div className="bg-[#191d2e] p-3 rounded-xl border border-[#2b324d]">
              <span className="text-[11px] text-slate-400 font-medium block">Stale Hits (SWR)</span>
              <span className="text-lg font-bold text-amber-400 font-mono">
                {cacheInfo.metricas.staleHits}
              </span>
              <span className="text-[10px] text-slate-500 block mt-0.5">Instantâneo + sync background</span>
            </div>

            <div className="bg-[#191d2e] p-3 rounded-xl border border-[#2b324d]">
              <span className="text-[11px] text-slate-400 font-medium block">Chamadas à Microsoft</span>
              <span className="text-lg font-bold text-indigo-400 font-mono">
                {cacheInfo.metricas.misses}
              </span>
              <span className="text-[10px] text-slate-500 block mt-0.5">Cache Misses / Cold starts</span>
            </div>

            <div className="bg-[#191d2e] p-3 rounded-xl border border-[#2b324d]">
              <span className="text-[11px] text-slate-400 font-medium block">Token Microsoft</span>
              <span className="text-lg font-bold text-purple-300 font-mono">
                {cacheInfo.tokenValido ? `${cacheInfo.tokenExpiraEmSegundos}s` : 'Expirado'}
              </span>
              <span className="text-[10px] text-slate-500 block mt-0.5">
                {cacheInfo.tokenValido ? 'Reutilizando token ativo' : 'Pronto para renovar'}
              </span>
            </div>
          </div>

          {/* Status das listas cacheadas */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
            <div className="bg-[#10121d] p-3 rounded-xl border border-[#23283e] flex items-center justify-between">
              <div>
                <span className="text-xs font-bold text-white block">BD_membros</span>
                <span className="text-[11px] text-slate-400">
                  {cacheInfo.listas.membros.total} membros • {cacheInfo.listas.membros.idadeTexto || 'agora'}
                </span>
              </div>
              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                !cacheInfo.listas.membros.isStale ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30' : 'bg-amber-950 text-amber-300 border border-amber-500/30'
              }`}>
                {!cacheInfo.listas.membros.isStale ? 'Fresco' : 'Stale (SWR)'}
              </span>
            </div>

            <div className="bg-[#10121d] p-3 rounded-xl border border-[#23283e] flex items-center justify-between">
              <div>
                <span className="text-xs font-bold text-white block">BD_Relatorio</span>
                <span className="text-[11px] text-slate-400">
                  {cacheInfo.listas.relatorios.total} relatórios • {cacheInfo.listas.relatorios.idadeTexto || 'agora'}
                </span>
              </div>
              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                !cacheInfo.listas.relatorios.isStale ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30' : 'bg-amber-950 text-amber-300 border border-amber-500/30'
              }`}>
                {!cacheInfo.listas.relatorios.isStale ? 'Fresco' : 'Stale (SWR)'}
              </span>
            </div>

            <div className="bg-[#10121d] p-3 rounded-xl border border-[#23283e] flex items-center justify-between">
              <div>
                <span className="text-xs font-bold text-white block">BD_celulas</span>
                <span className="text-[11px] text-slate-400">
                  {cacheInfo.listas.celulas.total} células • {cacheInfo.listas.celulas.idadeTexto || 'agora'}
                </span>
              </div>
              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                !cacheInfo.listas.celulas.isStale ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30' : 'bg-amber-950 text-amber-300 border border-amber-500/30'
              }`}>
                {!cacheInfo.listas.celulas.isStale ? 'Fresco' : 'Stale (SWR)'}
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Cards de Ambiente e Comparativo Vercel vs Local */}
      {data && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {/* Card 1: Runtime & Plataforma */}
          <div className="bg-[#171a29] border border-[#2d334d] rounded-2xl p-4 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                <Server className="w-4 h-4 text-indigo-400" />
                Ambiente de Execução
              </span>
              <span className={`text-[10px] font-mono font-bold px-2 py-0.5 rounded ${
                data.environment.isVercel ? 'bg-purple-950/80 text-purple-300 border border-purple-500/30' : 'bg-blue-950/80 text-blue-300 border border-blue-500/30'
              }`}>
                {data.environment.isVercel ? 'Vercel Serverless' : 'Local / Container'}
              </span>
            </div>
            
            <div className="space-y-1.5 text-xs">
              <div className="flex justify-between py-1 border-b border-[#24293f]">
                <span className="text-slate-400">Node Runtime:</span>
                <span className="font-mono text-white">{data.environment.nodeVersion}</span>
              </div>
              <div className="flex justify-between py-1 border-b border-[#24293f]">
                <span className="text-slate-400">Região Cloud:</span>
                <span className="font-mono text-white">{data.environment.vercelRegion}</span>
              </div>
              <div className="flex justify-between py-1">
                <span className="text-slate-400">Duração Total:</span>
                <span className="font-mono text-indigo-300 font-bold">{data.totalDurationMs} ms</span>
              </div>
            </div>
          </div>

          {/* Card 2: Credenciais do SharePoint */}
          <div className="bg-[#171a29] border border-[#2d334d] rounded-2xl p-4 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                <KeyRound className="w-4 h-4 text-amber-400" />
                Credenciais Microsoft
              </span>
              <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-emerald-950/80 text-emerald-300 border border-emerald-500/30">
                Ativo
              </span>
            </div>
            
            <div className="space-y-1.5 text-xs">
              <div className="flex justify-between py-1 border-b border-[#24293f]">
                <span className="text-slate-400">Conta Microsoft:</span>
                <span className="font-mono text-white truncate max-w-[170px]" title={data.environment.sharepointUserMasked}>
                  {data.environment.sharepointUserMasked}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-[#24293f]">
                <span className="text-slate-400">Senha:</span>
                <span className="font-mono text-emerald-400">
                  {data.environment.sharepointPassConfigured ? `Configurada (${data.environment.sharepointPassLength} chars)` : 'Não configurada'}
                </span>
              </div>
              <div className="flex justify-between py-1">
                <span className="text-slate-400">Client ID:</span>
                <span className="font-mono text-slate-300 text-[10px] truncate max-w-[170px]" title={data.environment.clientId}>
                  {data.environment.clientId}
                </span>
              </div>
            </div>
          </div>

          {/* Card 3: Site URL do SharePoint */}
          <div className="bg-[#171a29] border border-[#2d334d] rounded-2xl p-4 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                <Globe className="w-4 h-4 text-emerald-400" />
                Destino SharePoint
              </span>
              <span className="text-[10px] font-mono text-slate-400">
                Tenant Microsoft 365
              </span>
            </div>
            
            <div className="space-y-1.5 text-xs">
              <div className="py-1">
                <span className="text-slate-400 block mb-1">URL Base do Site:</span>
                <div className="p-2 rounded-lg bg-[#11131d] border border-[#24293f] font-mono text-[11px] text-indigo-300 break-all">
                  {data.environment.siteUrl}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Lista de Endpoints Verificados */}
      <div className="bg-[#161a29] border border-[#2d334d] rounded-2xl overflow-hidden shadow-xl">
        <div className="bg-[#1c2133] px-5 py-3.5 border-b border-[#2d334d] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Layers className="w-4 h-4 text-indigo-400" />
            <h2 className="text-sm font-bold text-white">
              Status Individual dos Endpoints & Listas do SharePoint
            </h2>
          </div>
          <span className="text-xs text-slate-400">
            {data?.checks?.length ?? 0} endpoints analisados
          </span>
        </div>

        <div className="divide-y divide-[#252a40]">
          {data?.checks?.map((check) => {
            const isExpanded = expandedCheck === check.id;

            return (
              <div key={check.id} className="p-4 sm:p-5 hover:bg-[#1a1e30] transition-colors">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  
                  {/* Nome e Status Icon */}
                  <div className="flex items-center gap-3">
                    <div className="shrink-0">
                      {check.success ? (
                        <CheckCircle2 className="w-5 h-5 text-emerald-400" />
                      ) : (
                        <XCircle className="w-5 h-5 text-rose-400" />
                      )}
                    </div>
                    <div>
                      <div className="flex items-center gap-2">
                        <h3 className="text-sm font-bold text-white">
                          {check.name}
                        </h3>
                        <span className={`text-[10px] font-mono font-bold px-2 py-0.5 rounded ${
                          check.status === 200
                            ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/40'
                            : check.status === 404
                            ? 'bg-rose-950 text-rose-300 border border-rose-500/40'
                            : check.status >= 500
                            ? 'bg-red-950 text-red-300 border border-red-500/50'
                            : 'bg-amber-950 text-amber-300 border border-amber-500/40'
                        }`}>
                          HTTP {check.status || 'Falha'} {check.statusText ? `• ${check.statusText}` : ''}
                        </span>
                      </div>
                      <p className="text-[11px] font-mono text-slate-400 mt-0.5 break-all">
                        <span className="text-indigo-400 font-bold mr-1">{check.method}</span>
                        {check.endpoint}
                      </p>
                    </div>
                  </div>

                  {/* Latência e Botão Expandir */}
                  <div className="flex items-center gap-3 ml-auto">
                    <div className="flex items-center gap-1.5 bg-[#11131e] px-2.5 py-1 rounded-lg border border-[#2b3149] text-xs font-mono text-slate-300">
                      <Clock className="w-3.5 h-3.5 text-indigo-400" />
                      <span>{check.durationMs} ms</span>
                    </div>

                    <button
                      onClick={() => setExpandedCheck(isExpanded ? null : check.id)}
                      className="p-1.5 rounded-lg bg-[#202538] hover:bg-[#2c334d] text-slate-300 hover:text-white transition-colors cursor-pointer"
                      title={isExpanded ? 'Recolher detalhes' : 'Ver payload completo'}
                    >
                      {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                    </button>
                  </div>
                </div>

                {/* Resumo Rápido */}
                <div className="mt-2.5 text-xs">
                  {check.success ? (
                    <p className="text-emerald-300/90 font-medium">
                      {check.details || `Operação executada com sucesso.${check.itemCount !== undefined ? ` Total: ${check.itemCount} registros.` : ''}`}
                    </p>
                  ) : (
                    <div className="p-2.5 rounded-lg bg-rose-950/60 border border-rose-500/40 text-rose-200 text-xs space-y-1">
                      <p className="font-bold flex items-center gap-1.5 text-rose-300">
                        <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
                        <span>Diagnóstico do Erro {check.status ? `(HTTP ${check.status})` : ''}:</span>
                      </p>
                      <p className="font-mono text-[11px]">{check.error}</p>
                    </div>
                  )}
                </div>

                {/* Detalhes Expandidos / JSON */}
                {isExpanded && (
                  <div className="mt-3 p-3 rounded-xl bg-[#0e1017] border border-[#272d42] font-mono text-[11px] text-slate-300 overflow-x-auto space-y-2 animate-in fade-in duration-150">
                    <div className="flex items-center justify-between text-slate-500 text-[10px] uppercase font-bold border-b border-[#202638] pb-1">
                      <span>Payload Técnico / Resposta</span>
                      {check.guid && <span>GUID: {check.guid}</span>}
                    </div>
                    <pre className="whitespace-pre-wrap break-all text-[11px] text-slate-300">
                      {typeof check.details === 'object' ? JSON.stringify(check.details, null, 2) : String(check.details || check.error || 'Nenhum detalhe adicional')}
                    </pre>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Testador Rápido de Login de Membro */}
      <div className="bg-[#161a29] border border-[#2d334d] rounded-2xl p-5 space-y-4">
        <div className="flex items-center gap-2 border-b border-[#2d334d] pb-3">
          <Terminal className="w-5 h-5 text-indigo-400" />
          <div>
            <h2 className="text-sm font-bold text-white">
              Testador de Login de Membro no SharePoint (BD_membros)
            </h2>
            <p className="text-xs text-slate-400">
              Verifique se um usuário específico (ex: Jfonteles, Rai, Jeff) é autenticado contra a base oficial do SharePoint
            </p>
          </div>
        </div>

        <form onSubmit={handleTestarLoginMembro} className="flex flex-wrap items-end gap-3">
          <div className="flex-1 min-w-[200px]">
            <label className="block text-xs font-bold text-slate-300 mb-1">
              Login do Membro (Login / Email / Nome)
            </label>
            <input
              type="text"
              value={testLogin}
              onChange={(e) => setTestLogin(e.target.value)}
              placeholder="ex: Jfonteles"
              className="w-full bg-[#11131e] border border-[#2e354e] rounded-xl px-3.5 py-2 text-xs text-white focus:outline-none focus:border-indigo-400 font-mono"
              required
            />
          </div>

          <div className="flex-1 min-w-[200px]">
            <label className="block text-xs font-bold text-slate-300 mb-1">
              Senha (opcional para teste de busca)
            </label>
            <input
              type="password"
              value={testSenha}
              onChange={(e) => setTestSenha(e.target.value)}
              placeholder="Senha do membro"
              className="w-full bg-[#11131e] border border-[#2e354e] rounded-xl px-3.5 py-2 text-xs text-white focus:outline-none focus:border-indigo-400 font-mono"
            />
          </div>

          <button
            type="submit"
            disabled={testLoading}
            className="px-5 py-2 rounded-xl bg-[#242a42] hover:bg-[#2f3757] text-white text-xs font-bold transition-all cursor-pointer disabled:opacity-60 border border-[#3b4468] h-[38px] flex items-center gap-2"
          >
            {testLoading ? (
              <>
                <RotateCw className="w-3.5 h-3.5 animate-spin text-indigo-400" />
                <span>Testando...</span>
              </>
            ) : (
              <>
                <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" />
                <span>Consultar Membro</span>
              </>
            )}
          </button>
        </form>

        {testResultado && (
          <div className={`p-4 rounded-xl border text-xs animate-in fade-in duration-150 ${
            testResultado.sucesso 
              ? 'bg-emerald-950/60 border-emerald-500/50 text-emerald-100' 
              : 'bg-rose-950/60 border-rose-500/50 text-rose-100'
          }`}>
            <p className="font-bold text-sm mb-1">
              {testResultado.sucesso ? '✅ Membro Encontrado & Autenticado!' : '❌ Falha na Consulta:'}
            </p>
            {testResultado.sucesso ? (
              <div className="space-y-1 text-slate-200">
                <p><strong>Nome:</strong> {testResultado.membro?.nome}</p>
                <p><strong>Login:</strong> {testResultado.membro?.login}</p>
                <p><strong>Cargo / Função:</strong> {testResultado.membro?.cargo || 'Membro'}</p>
                <p><strong>Setor / Célula:</strong> {testResultado.membro?.setor || '-'} • {testResultado.membro?.celula || '-'}</p>
              </div>
            ) : (
              <p className="text-rose-200">{testResultado.erro || 'Membro não encontrado ou credenciais incorretas.'}</p>
            )}
          </div>
        )}
      </div>

    </div>
  );
};
