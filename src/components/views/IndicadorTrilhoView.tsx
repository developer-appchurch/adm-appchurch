import React, { useState, useEffect, useMemo } from 'react';
import {
  GraduationCap,
  TrendingUp,
  Users,
  CheckCircle2,
  Clock,
  AlertCircle,
  Search,
  Filter,
  RefreshCw,
  Layers,
  ChevronRight,
  Sparkles,
  BarChart3,
  PieChart as PieChartIcon,
  Award,
  ArrowUpRight,
  Download,
  Building2,
  Compass
} from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  CartesianGrid,
  PieChart,
  Pie,
  Cell,
  LabelList,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar
} from 'recharts';
import { IndicadorTrilhoResponse, MembroTrilhoItem, MetricasTrilhoConjunto } from '../../types';

type FiltroVisao = 'geral' | 'area' | 'setor';
type AbaAtiva = 'graficos' | 'membros' | 'comparativo';

// Cores temáticas para gráficos no Dark Theme
const PALETA_CORES = [
  '#6366f1', // Indigo
  '#06b6d4', // Cyan
  '#10b981', // Emerald
  '#f59e0b', // Amber
  '#ec4899', // Pink
  '#8b5cf6', // Violet
  '#3b82f6', // Blue
  '#14b8a6', // Teal
];

const CORES_STATUS = {
  completo: '#10b981',
  emAndamento: '#f59e0b',
  naoIniciado: '#64748b'
};

export const IndicadorTrilhoView: React.FC = () => {
  const [data, setData] = useState<IndicadorTrilhoResponse | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [refreshing, setRefreshing] = useState<boolean>(false);
  const [erro, setErro] = useState<string | null>(null);

  // Filtros principais
  const [visao, setVisao] = useState<FiltroVisao>('geral');
  const [areaSelecionada, setAreaSelecionada] = useState<string>('todos');
  const [setorSelecionado, setSetorSelecionado] = useState<string>('todos');
  const [abaAtiva, setAbaAtiva] = useState<AbaAtiva>('graficos');

  // Filtros de busca de membros
  const [buscaMembro, setBuscaMembro] = useState<string>('');
  const [filtroStatusMembro, setFiltroStatusMembro] = useState<string>('todos');
  const [filtroEtapaPendente, setFiltroEtapaPendente] = useState<string>('todos');

  // Carrega os dados da API
  const carregarDados = async (forcar: boolean = false) => {
    if (forcar) setRefreshing(true);
    else setLoading(true);
    setErro(null);

    try {
      const url = `/api/sharepoint/indicador-trilho${forcar ? '?refresh=true' : ''}`;
      const res = await fetch(url);
      if (!res.ok) {
        throw new Error(`Erro HTTP ${res.status} ao carregar Indicador do Trilho`);
      }
      const json: IndicadorTrilhoResponse = await res.json();
      if (json.sucesso) {
        setData(json);
      } else {
        throw new Error('Falha ao processar métricas do trilho.');
      }
    } catch (err: any) {
      console.error('[IndicadorTrilho] Erro:', err);
      setErro(err?.message || 'Não foi possível carregar as métricas do trilho.');
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  useEffect(() => {
    carregarDados();
  }, []);

  // Metadados e opções derivadas
  const etapas = useMemo(() => data?.etapas || [], [data]);
  const areasDisponiveis = useMemo(() => data?.areasDisponiveis || [], [data]);
  const setoresDisponiveis = useMemo(() => data?.setoresDisponiveis || [], [data]);

  // Lista de membros filtrada com base na visão atual (Geral, Área, Setor)
  const membrosFiltradosPorVisao = useMemo(() => {
    if (!data?.membros) return [];
    let lista = data.membros;

    if (visao === 'area' && areaSelecionada !== 'todos') {
      lista = lista.filter(m => m.area.toLowerCase() === areaSelecionada.toLowerCase());
    } else if (visao === 'setor' && setorSelecionado !== 'todos') {
      lista = lista.filter(m => m.setor.toLowerCase() === setorSelecionado.toLowerCase());
    }

    return lista;
  }, [data, visao, areaSelecionada, setorSelecionado]);

  // Métricas calculadas para a visão atual selecionada
  const metricasAtuais = useMemo(() => {
    const total = membrosFiltradosPorVisao.length;
    if (total === 0) {
      return {
        totalMembros: 0,
        percentualMedio: 0,
        membrosCompletos: 0,
        pctCompletos: '0.0',
        membrosNaoCompletos: 0,
        pctNaoCompletos: '0.0',
        membrosComZero: 0,
        pctComZero: '0.0',
        membrosEmAndamento: 0,
        pctEmAndamento: '0.0',
        membrosNaoIniciados: 0,
        etapasStats: etapas.map(etapa => ({
          etapa,
          concluidos: 0,
          pendentes: 0,
          percentual: 0
        })),
        maxPercentual: 0,
        etapasDestaque: [] as { etapa: string; concluidos: number; pendentes: number; percentual: number }[],
        minPercentual: 0,
        etapasGargalo: [] as { etapa: string; concluidos: number; pendentes: number; percentual: number }[]
      };
    }

    const etapasStats = etapas.map(etapa => {
      const concluidos = membrosFiltradosPorVisao.filter(m => m.etapasConcluidas.includes(etapa)).length;
      const pendentes = total - concluidos;
      const percentual = parseFloat(((concluidos / total) * 100).toFixed(1));
      return {
        etapa,
        concluidos,
        pendentes,
        percentual
      };
    });

    // Quantidade de membros que completaram 100% do trilho
    const membrosCompletos = membrosFiltradosPorVisao.filter(m => 
      m.statusTrilho === 'Completo' || (etapas.length > 0 && m.etapasConcluidas.length >= etapas.length)
    ).length;
    const pctCompletos = ((membrosCompletos / total) * 100).toFixed(1);

    // Membros que ainda não completaram o trilho (< 100%)
    const membrosNaoCompletos = total - membrosCompletos;
    const pctNaoCompletos = ((membrosNaoCompletos / total) * 100).toFixed(1);

    // Membros com 0% do trilho (não iniciados)
    const membrosComZero = membrosFiltradosPorVisao.filter(m => m.etapasConcluidas.length === 0).length;
    const pctComZero = ((membrosComZero / total) * 100).toFixed(1);

    // Membros em andamento (> 0% e < 100%)
    const membrosEmAndamento = total - membrosCompletos - membrosComZero;
    const pctEmAndamento = ((membrosEmAndamento / total) * 100).toFixed(1);

    // Conclusão Média do Trilho: Média de quantos % do trilho foi concluído de todos os membros dentro do filtro selecionado
    const somaPercentuais = membrosFiltradosPorVisao.reduce((acc, m) => {
      const pctMembro = etapas.length > 0 ? (m.etapasConcluidas.length / etapas.length) * 100 : 0;
      return acc + pctMembro;
    }, 0);
    const percentualMedio = parseFloat((somaPercentuais / total).toFixed(1));

    // Destaque (Maior porcentagem)
    const maxPercentual = etapasStats.length > 0 ? Math.max(...etapasStats.map(e => e.percentual)) : 0;
    const etapasDestaque = etapasStats.filter(e => e.percentual === maxPercentual);

    // Gargalo (Menor porcentagem - com detecção de empate)
    const minPercentual = etapasStats.length > 0 ? Math.min(...etapasStats.map(e => e.percentual)) : 0;
    const etapasGargalo = etapasStats.filter(e => e.percentual === minPercentual);

    return {
      totalMembros: total,
      percentualMedio,
      membrosCompletos,
      pctCompletos,
      membrosNaoCompletos,
      pctNaoCompletos,
      membrosComZero,
      pctComZero,
      membrosEmAndamento,
      pctEmAndamento,
      membrosNaoIniciados: membrosComZero,
      etapasStats,
      maxPercentual,
      etapasDestaque,
      minPercentual,
      etapasGargalo
    };
  }, [membrosFiltradosPorVisao, etapas]);

  // Dados para o Gráfico de Barras de Etapas
  const dadosGraficoEtapas = useMemo(() => {
    return metricasAtuais.etapasStats.map(item => ({
      etapa: item.etapa,
      percentual: item.percentual,
      concluidos: item.concluidos,
      pendentes: item.pendentes
    }));
  }, [metricasAtuais]);

  // Dados para o Gráfico de Pizza de Status
  const dadosGraficoStatus = useMemo(() => {
    return [
      { name: 'Completo (100%)', value: metricasAtuais.membrosCompletos, color: CORES_STATUS.completo },
      { name: 'Em Andamento', value: metricasAtuais.membrosEmAndamento, color: CORES_STATUS.emAndamento },
      { name: 'Não Iniciado', value: metricasAtuais.membrosNaoIniciados, color: CORES_STATUS.naoIniciado }
    ].filter(item => item.value > 0);
  }, [metricasAtuais]);

  // Dados para o Comparativo entre Áreas (visão Área)
  const dadosComparativoAreas = useMemo(() => {
    if (!data?.porArea) return [];
    return Object.entries(data.porArea).map(([nomeArea, stats]: [string, MetricasTrilhoConjunto]) => ({
      area: nomeArea,
      percentualMedio: stats.percentualMedio,
      totalMembros: stats.totalMembros,
      completos: stats.membrosCompletos
    })).sort((a, b) => b.percentualMedio - a.percentualMedio);
  }, [data]);

  // Dados para o Comparativo entre Setores (visão Setor)
  const dadosComparativoSetores = useMemo(() => {
    if (!data?.porSetor) return [];
    return Object.entries(data.porSetor).map(([nomeSetor, stats]: [string, MetricasTrilhoConjunto]) => ({
      setor: nomeSetor,
      percentualMedio: stats.percentualMedio,
      totalMembros: stats.totalMembros,
      completos: stats.membrosCompletos
    })).sort((a, b) => b.percentualMedio - a.percentualMedio);
  }, [data]);

  // Lista de Membros filtrada para a Tabela
  const membrosTabela = useMemo(() => {
    return membrosFiltradosPorVisao.filter(m => {
      // Filtro de busca textual
      const matchBusca = buscaMembro.trim() === '' ||
        m.nome.toLowerCase().includes(buscaMembro.toLowerCase()) ||
        m.login.toLowerCase().includes(buscaMembro.toLowerCase()) ||
        m.setor.toLowerCase().includes(buscaMembro.toLowerCase()) ||
        m.cargo.toLowerCase().includes(buscaMembro.toLowerCase()) ||
        m.celula.toLowerCase().includes(buscaMembro.toLowerCase());

      // Filtro de Status
      const matchStatus = filtroStatusMembro === 'todos' || m.statusTrilho === filtroStatusMembro;

      // Filtro de Etapa Pendente
      const matchEtapaPendente = filtroEtapaPendente === 'todos' || !m.etapasConcluidas.includes(filtroEtapaPendente);

      return matchBusca && matchStatus && matchEtapaPendente;
    });
  }, [membrosFiltradosPorVisao, buscaMembro, filtroStatusMembro, filtroEtapaPendente]);

  // Função para exportar CSV da lista filtrada de membros
  const exportarCSV = () => {
    if (membrosTabela.length === 0) return;
    const header = ['Nome', 'Login', 'Setor', 'Área', 'Cargo', 'Célula', 'Progresso %', 'Status', ...etapas];
    const rows = membrosTabela.map(m => [
      `"${m.nome}"`,
      `"${m.login}"`,
      `"${m.setor}"`,
      `"${m.area}"`,
      `"${m.cargo}"`,
      `"${m.celula}"`,
      `${m.percentualConclusao}%`,
      `"${m.statusTrilho}"`,
      ...etapas.map(e => m.etapasConcluidas.includes(e) ? '"Concluído"' : '"Pendente"')
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [header.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `indicador_trilho_${visao}_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (loading) {
    return (
      <div id="indicador-trilho-loading" className="p-4 sm:p-6 lg:p-8 space-y-6 max-w-7xl mx-auto animate-fadeIn">
        {/* Banner de carregamento com progresso animado */}
        <div className="bg-[#141724] border border-indigo-500/30 p-6 sm:p-8 rounded-2xl shadow-2xl relative overflow-hidden">
          <div className="absolute inset-0 bg-gradient-to-r from-indigo-500/10 via-purple-500/10 to-transparent animate-pulse" />
          
          <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="flex items-center gap-4">
              <div className="w-14 h-14 sm:w-16 sm:h-16 rounded-2xl bg-indigo-500/20 border border-indigo-500/40 flex items-center justify-center shrink-0">
                <GraduationCap className="w-8 h-8 sm:w-9 sm:h-9 text-indigo-400 animate-bounce" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-lg sm:text-2xl font-black text-white tracking-tight">
                    Consultando Indicador do Trilho...
                  </h2>
                  <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 animate-pulse">
                    SharePoint Online
                  </span>
                </div>
                <p className="text-xs sm:text-sm text-slate-300 mt-1">
                  Carregando BD_Capacitacao, cruzando com mais de 6.600 registros de BD_MembrosCapac e vinculando aos membros.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-3 bg-[#1c2032] px-4 py-2.5 rounded-xl border border-indigo-500/30 shrink-0">
              <RefreshCw className="w-4 h-4 text-indigo-400 animate-spin" />
              <div className="text-xs text-slate-200">
                <p className="font-bold">Processando Métricas</p>
                <p className="text-[10px] text-slate-400">Por favor, aguarde...</p>
              </div>
            </div>
          </div>

          {/* Barra de progresso animada */}
          <div className="relative z-10 mt-6 w-full bg-[#1b1f30] rounded-full h-2 overflow-hidden border border-[#2c334d]">
            <div className="bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-400 h-full rounded-full animate-pulse w-3/4 transition-all duration-1000" />
          </div>
        </div>

        {/* Skeleton Grid de KPI Cards */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {[
            { label: 'Total de Membros', cor: 'indigo' },
            { label: 'Conclusão Média do Trilho', cor: 'emerald' },
            { label: 'Trilho Completo (100%)', cor: 'purple' },
            { label: 'Destaque & Gargalo', cor: 'amber' }
          ].map((card, i) => (
            <div key={i} className="bg-[#161a29] border border-[#272d42] rounded-2xl p-5 shadow-lg space-y-3 animate-pulse">
              <div className="flex items-center justify-between">
                <div className="h-3 bg-slate-700/60 rounded w-24" />
                <div className="w-7 h-7 rounded-lg bg-slate-700/50" />
              </div>
              <div className="h-8 bg-slate-700/80 rounded w-20" />
              <div className="h-2 bg-slate-700/40 rounded w-full" />
              <div className="h-3 bg-slate-700/40 rounded w-32" />
            </div>
          ))}
        </div>

        {/* Skeleton do Gráfico Principal */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-6 shadow-xl space-y-4 animate-pulse">
          <div className="flex items-center justify-between border-b border-[#23283c] pb-3">
            <div className="h-4 bg-slate-700/60 rounded w-48" />
            <div className="h-3 bg-slate-700/40 rounded w-28" />
          </div>
          <div className="h-64 bg-[#10121d] rounded-xl flex items-end justify-between p-6 gap-3">
            {[45, 60, 30, 80, 55, 40, 70, 90, 65, 50, 75, 85].map((h, idx) => (
              <div key={idx} className="flex-1 bg-indigo-500/20 rounded-t-lg transition-all" style={{ height: `${h}%` }} />
            ))}
          </div>
        </div>
      </div>
    );
  }

  if (erro) {
    return (
      <div className="p-6 max-w-2xl mx-auto my-12 bg-rose-950/40 border border-rose-500/50 rounded-2xl text-center space-y-4">
        <AlertCircle className="w-10 h-10 text-rose-400 mx-auto" />
        <h3 className="text-lg font-bold text-white">Erro ao carregar dados do Trilho</h3>
        <p className="text-xs text-rose-200">{erro}</p>
        <button
          onClick={() => carregarDados(true)}
          className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold transition-all cursor-pointer inline-flex items-center gap-2"
        >
          <RefreshCw className="w-3.5 h-3.5" />
          <span>Tentar Novamente</span>
        </button>
      </div>
    );
  }

  return (
    <div id="indicador-trilho-container" className="p-4 sm:p-6 lg:p-8 space-y-6 max-w-7xl mx-auto">
      {/* Header Principal */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-[#141724] border border-[#272d42] p-5 rounded-2xl shadow-xl">
        <div className="flex items-center gap-3.5">
          <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-indigo-600 to-indigo-400 flex items-center justify-center shadow-lg shadow-indigo-600/30 text-white shrink-0">
            <GraduationCap className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h1 className="text-xl sm:text-2xl font-black text-white tracking-tight">
                Indicador do Trilho de Capacitação
              </h1>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
                Paz Church Sobral
              </span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Acompanhamento de avanço nas etapas de formação ministerial e liderança
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            onClick={() => carregarDados(true)}
            disabled={refreshing}
            className="flex items-center gap-2 px-3.5 py-2 rounded-xl bg-[#202538] hover:bg-[#2b324d] text-slate-200 hover:text-white text-xs font-semibold border border-[#313a57] transition-all cursor-pointer disabled:opacity-50"
            title="Recarregar dados do SharePoint"
          >
            <RefreshCw className={`w-3.5 h-3.5 text-indigo-400 ${refreshing ? 'animate-spin' : ''}`} />
            <span>{refreshing ? 'Sincronizando...' : 'Atualizar Dados'}</span>
          </button>

          <button
            onClick={exportarCSV}
            className="flex items-center gap-2 px-3.5 py-2 rounded-xl bg-emerald-950/80 hover:bg-emerald-900 text-emerald-200 text-xs font-semibold border border-emerald-500/30 transition-all cursor-pointer"
            title="Exportar dados filtrados em CSV"
          >
            <Download className="w-3.5 h-3.5 text-emerald-400" />
            <span>Exportar CSV</span>
          </button>
        </div>
      </div>

      {/* Barra de Filtros de Visão (Igreja em Geral | Por Área | Por Setor) */}
      <div className="bg-[#181c2b] border border-[#2b324d] rounded-2xl p-4 shadow-lg flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
        {/* Seletor de Escopo de Visão */}
        <div className="flex items-center gap-1.5 p-1 bg-[#10121d] rounded-xl border border-[#242a3e] w-full sm:w-auto">
          <button
            onClick={() => setVisao('geral')}
            className={`flex-1 sm:flex-initial px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center justify-center gap-2 ${
              visao === 'geral'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-[#1a1e30]'
            }`}
          >
            <Building2 className="w-3.5 h-3.5" />
            <span>Igreja em Geral</span>
          </button>

          <button
            onClick={() => setVisao('area')}
            className={`flex-1 sm:flex-initial px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center justify-center gap-2 ${
              visao === 'area'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-[#1a1e30]'
            }`}
          >
            <Compass className="w-3.5 h-3.5" />
            <span>Por Área</span>
          </button>

          <button
            onClick={() => setVisao('setor')}
            className={`flex-1 sm:flex-initial px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center justify-center gap-2 ${
              visao === 'setor'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-[#1a1e30]'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>Por Setor</span>
          </button>
        </div>

        {/* Dropdowns de Filtro Específico (quando em Área ou Setor) */}
        <div className="flex items-center gap-3 w-full lg:w-auto flex-wrap">
          {visao === 'area' && (
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <label className="text-xs font-bold text-slate-300 shrink-0">Selecionar Área:</label>
              <select
                value={areaSelecionada}
                onChange={(e) => setAreaSelecionada(e.target.value)}
                className="bg-[#10121d] border border-[#2f3754] text-white text-xs rounded-xl px-3.5 py-2 font-medium focus:outline-none focus:border-indigo-400 cursor-pointer w-full sm:w-auto"
              >
                <option value="todos">Todas as Áreas (Consolidado)</option>
                {areasDisponiveis.map(area => (
                  <option key={area} value={area}>{area}</option>
                ))}
              </select>
            </div>
          )}

          {visao === 'setor' && (
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <label className="text-xs font-bold text-slate-300 shrink-0">Selecionar Setor:</label>
              <select
                value={setorSelecionado}
                onChange={(e) => setSetorSelecionado(e.target.value)}
                className="bg-[#10121d] border border-[#2f3754] text-white text-xs rounded-xl px-3.5 py-2 font-medium focus:outline-none focus:border-indigo-400 cursor-pointer w-full sm:w-auto"
              >
                <option value="todos">Todos os Setores (Consolidado)</option>
                {setoresDisponiveis.map(setor => (
                  <option key={setor} value={setor}>{setor}</option>
                ))}
              </select>
            </div>
          )}

          {/* Abas de visualização interna (Gráficos vs Lista de Membros vs Comparativo) */}
          <div className="flex items-center gap-1 p-1 bg-[#10121d] rounded-xl border border-[#242a3e] ml-auto">
            <button
              onClick={() => setAbaAtiva('graficos')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                abaAtiva === 'graficos'
                  ? 'bg-[#252b42] text-indigo-300 border border-indigo-500/40'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span>Gráficos</span>
            </button>

            <button
              onClick={() => setAbaAtiva('comparativo')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                abaAtiva === 'comparativo'
                  ? 'bg-[#252b42] text-indigo-300 border border-indigo-500/40'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <TrendingUp className="w-3.5 h-3.5" />
              <span>Comparativo</span>
            </button>

            <button
              onClick={() => setAbaAtiva('membros')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                abaAtiva === 'membros'
                  ? 'bg-[#252b42] text-indigo-300 border border-indigo-500/40'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <Users className="w-3.5 h-3.5" />
              <span>Membros ({membrosFiltradosPorVisao.length})</span>
            </button>
          </div>
        </div>
      </div>

      {/* KPI Cards de Resumo */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Total de Membros */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden group hover:border-indigo-500/40 transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Total de Membros
              </span>
              <div className="p-2 rounded-xl bg-indigo-500/15 text-indigo-400">
                <Users className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl sm:text-3xl font-black text-white font-mono">
                {metricasAtuais.totalMembros}
              </span>
              <span className="text-xs text-slate-400 font-medium">
                {visao === 'geral' ? 'na Igreja' : visao === 'area' ? `na Área ${areaSelecionada}` : `no Setor ${setorSelecionado}`}
              </span>
            </div>
            <div className="mt-3 text-[11px] text-slate-400 space-y-1">
              <div className="flex items-center justify-between">
                <span>Com capacitações:</span>
                <span className="text-emerald-400 font-bold font-mono">
                  {metricasAtuais.totalMembros - metricasAtuais.membrosComZero}
                </span>
              </div>
              <div className="flex items-center justify-between">
                <span>Sem capacitações (0%):</span>
                <span className="text-rose-400 font-bold font-mono">
                  {metricasAtuais.membrosComZero}
                </span>
              </div>
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-[#23283c] text-[10px] text-slate-400 flex items-center justify-between">
            <span className="text-slate-500">Relacionamento:</span>
            <span className="text-indigo-400 font-medium truncate" title="BD_Capacitacao (Capacitacao) ↔ BD_MembrosCapac (Capacitacao)">
              BD_Capacitacao ↔ BD_MembrosCapac
            </span>
          </div>
        </div>

        {/* Card 2: Conclusão Média do Trilho */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden group hover:border-indigo-500/40 transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Conclusão Média do Trilho
              </span>
              <div className="p-2 rounded-xl bg-emerald-500/15 text-emerald-400">
                <TrendingUp className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl sm:text-3xl font-black text-emerald-400 font-mono">
                {metricasAtuais.percentualMedio}%
              </span>
              <span className="text-xs text-slate-400 font-medium">média dos membros</span>
            </div>
            <div className="mt-3 w-full bg-[#202538] rounded-full h-2.5 overflow-hidden">
              <div
                className="bg-gradient-to-r from-emerald-500 to-teal-400 h-full rounded-full transition-all duration-500"
                style={{ width: `${Math.min(100, metricasAtuais.percentualMedio)}%` }}
              />
            </div>
            <p className="mt-2.5 text-[11px] text-slate-400 leading-snug">
              Média de quantos % do trilho foi concluído de todos os {metricasAtuais.totalMembros} membros dentro do filtro selecionado.
            </p>
          </div>
          <div className="mt-3 pt-2 border-t border-[#23283c] text-[10px] text-slate-500 flex items-center justify-between">
            <span>Escopo: {visao === 'geral' ? 'Toda Igreja' : visao === 'area' ? `Área ${areaSelecionada}` : `Setor ${setorSelecionado}`}</span>
            <span className="text-emerald-400 font-medium">{etapas.length} etapas base</span>
          </div>
        </div>

        {/* Card 3: Trilho Completo (100%) */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden group hover:border-indigo-500/40 transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Trilho Completo (100%)
              </span>
              <div className="p-2 rounded-xl bg-purple-500/15 text-purple-400">
                <Award className="w-4 h-4" />
              </div>
            </div>

            {/* Quantidade de membros com 100% e ao lado a % em relação ao total */}
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl sm:text-3xl font-black text-purple-300 font-mono">
                {metricasAtuais.membrosCompletos}
              </span>
              <span className="text-base font-bold text-purple-400 font-mono">
                ({metricasAtuais.pctCompletos}%)
              </span>
              <span className="text-xs text-slate-400 font-medium">do total</span>
            </div>

            {/* Contagem de quantos ainda não completaram e quantos com 0% */}
            <div className="mt-3 pt-2.5 border-t border-[#23283c] space-y-2 text-xs">
              <div className="flex items-center justify-between">
                <span className="text-slate-400">Ainda não concluíram:</span>
                <span className="font-semibold text-amber-300 font-mono">
                  {metricasAtuais.membrosNaoCompletos} <span className="text-[11px] font-normal text-slate-400">({metricasAtuais.pctNaoCompletos}%)</span>
                </span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-slate-400">Com 0% do trilho:</span>
                <span className="font-semibold text-rose-300 font-mono">
                  {metricasAtuais.membrosComZero} <span className="text-[11px] font-normal text-slate-400">({metricasAtuais.pctComZero}%)</span>
                </span>
              </div>
            </div>
          </div>

          <div className="mt-3 pt-2 border-t border-[#23283c] text-[10px] text-slate-500 flex items-center justify-between">
            <span>Em andamento (&gt;0% e &lt;100%):</span>
            <span className="text-amber-400 font-medium font-mono">
              {metricasAtuais.membrosEmAndamento} ({metricasAtuais.pctEmAndamento}%)
            </span>
          </div>
        </div>

        {/* Card 4: Destaque & Gargalo (com tratamento de empate) */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden group hover:border-indigo-500/40 transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Destaque &amp; Gargalo
              </span>
              <div className="p-2 rounded-xl bg-amber-500/15 text-amber-400">
                <Sparkles className="w-4 h-4" />
              </div>
            </div>

            <div className="mt-2.5 space-y-2.5 text-xs">
              {/* Maior Porcentagem (Destaque) */}
              <div className="bg-[#1b2033] border border-emerald-500/20 rounded-xl p-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-emerald-400 font-bold flex items-center gap-1">
                    🏆 Maior % ({metricasAtuais.maxPercentual}%)
                  </span>
                  {metricasAtuais.etapasDestaque.length > 1 && (
                    <span className="text-[10px] bg-emerald-500/20 text-emerald-300 font-semibold px-1.5 py-0.5 rounded">
                      {metricasAtuais.etapasDestaque.length} empatadas
                    </span>
                  )}
                </div>
                <p className="text-slate-200 text-xs font-medium mt-1 truncate" title={metricasAtuais.etapasDestaque.map(e => e.etapa).join(', ')}>
                  {metricasAtuais.etapasDestaque.map(e => e.etapa).join(', ') || '-'}
                </p>
              </div>

              {/* Menor Porcentagem (Gargalo - com tratamento explícito de empate) */}
              <div className="bg-[#1b2033] border border-amber-500/20 rounded-xl p-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-amber-400 font-bold flex items-center gap-1">
                    ⚠️ Menor % ({metricasAtuais.minPercentual}%)
                  </span>
                  {metricasAtuais.etapasGargalo.length > 1 ? (
                    <span className="text-[10px] bg-amber-500/20 text-amber-300 font-semibold px-1.5 py-0.5 rounded">
                      Empate ({metricasAtuais.etapasGargalo.length} etapas)
                    </span>
                  ) : (
                    <span className="text-[10px] text-slate-400">Etapa crítica</span>
                  )}
                </div>

                {metricasAtuais.etapasGargalo.length > 1 ? (
                  <div className="mt-1.5 space-y-1">
                    <span className="text-[11px] text-slate-400 block font-normal">
                      Etapas empatadas com menor índice:
                    </span>
                    <div className="flex flex-wrap gap-1 max-h-16 overflow-y-auto pr-1">
                      {metricasAtuais.etapasGargalo.map(eg => (
                        <span
                          key={eg.etapa}
                          className="px-1.5 py-0.5 bg-amber-500/10 border border-amber-500/30 text-amber-200 text-[10px] rounded font-medium truncate max-w-[140px]"
                          title={`${eg.etapa} (${eg.percentual}%)`}
                        >
                          {eg.etapa} ({eg.percentual}%)
                        </span>
                      ))}
                    </div>
                  </div>
                ) : (
                  <p className="text-slate-200 text-xs font-medium mt-1 truncate" title={metricasAtuais.etapasGargalo[0]?.etapa}>
                    {metricasAtuais.etapasGargalo[0]?.etapa || '-'}
                  </p>
                )}
              </div>
            </div>
          </div>

          <div className="mt-3 pt-2 border-t border-[#23283c] text-[10px] text-slate-500 flex items-center justify-between">
            <span>Base: {etapas.length} etapas</span>
            <span className="text-indigo-400 font-medium">BD_Capacitacao</span>
          </div>
        </div>
      </div>

      {/* ABA 1: GRÁFICOS */}
      {abaAtiva === 'graficos' && (
        <div className="space-y-6">
          {/* Linha 1: Gráfico Principal de Percentual por Etapa + Donut de Distribuição */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Gráfico de Barras por Etapa */}
            <div className="lg:col-span-2 bg-[#161a29] border border-[#272d42] rounded-2xl p-5 shadow-xl">
              <div className="flex items-center justify-between mb-4 border-b border-[#23283c] pb-3">
                <div className="flex items-center gap-2">
                  <BarChart3 className="w-4 h-4 text-indigo-400" />
                  <h2 className="text-sm font-bold text-white">
                    Percentual de Conclusão por Etapa do Trilho
                  </h2>
                </div>
                <span className="text-xs text-slate-400">
                  {visao === 'geral' ? 'Toda a Igreja' : visao === 'area' ? `Área: ${areaSelecionada}` : `Setor: ${setorSelecionado}`}
                </span>
              </div>

              <div 
                className="w-full"
                style={{ height: `${Math.max(280, dadosGraficoEtapas.length * 52)}px` }}
              >
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    layout="vertical"
                    data={dadosGraficoEtapas}
                    margin={{ top: 10, right: 45, left: 15, bottom: 10 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="#252b40" horizontal={false} />
                    <XAxis
                      type="number"
                      domain={[0, 100]}
                      stroke="#94a3b8"
                      fontSize={11}
                      unit="%"
                      tickLine={false}
                    />
                    <YAxis
                      type="category"
                      dataKey="etapa"
                      stroke="#94a3b8"
                      fontSize={12}
                      tickLine={false}
                      width={160}
                      tick={{ fill: '#e2e8f0', fontWeight: 600 }}
                    />
                    <Tooltip
                      content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          const item = payload[0].payload;
                          return (
                            <div className="bg-[#0f111a] border border-[#2e3650] p-3 rounded-xl shadow-2xl text-xs space-y-1">
                              <p className="font-bold text-white text-sm">{item.etapa}</p>
                              <p className="text-emerald-400 font-bold">
                                Conclusão: {item.percentual}% ({item.concluidos} membros)
                              </p>
                              <p className="text-slate-400">
                                Pendente: {100 - item.percentual}% ({item.pendentes} membros)
                              </p>
                            </div>
                          );
                        }
                        return null;
                      }}
                    />
                    <Bar 
                      dataKey="percentual" 
                      radius={[0, 8, 8, 0]} 
                      barSize={24}
                    >
                      <LabelList
                        dataKey="percentual"
                        position="right"
                        formatter={(val: any) => `${val}%`}
                        fill="#f8fafc"
                        fontSize={12}
                        fontWeight={800}
                        offset={10}
                      />
                      {dadosGraficoEtapas.map((entry, index) => (
                        <Cell
                          key={`cell-${index}`}
                          fill={PALETA_CORES[index % PALETA_CORES.length]}
                        />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Badges de Resumo Rápido das Etapas */}
              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2 mt-4 pt-3 border-t border-[#23283c]">
                {metricasAtuais.etapasStats.map((item, idx) => (
                  <div
                    key={item.etapa}
                    className="p-2.5 rounded-xl bg-[#10121d] border border-[#23283c] text-center"
                  >
                    <span className="text-[10px] text-slate-400 truncate block font-medium">
                      {item.etapa}
                    </span>
                    <span
                      className="text-sm font-bold font-mono block mt-0.5"
                      style={{ color: PALETA_CORES[idx % PALETA_CORES.length] }}
                    >
                      {item.percentual}%
                    </span>
                    <span className="text-[10px] text-slate-500 block">
                      {item.concluidos}/{metricasAtuais.totalMembros}
                    </span>
                  </div>
                ))}
              </div>
            </div>

            {/* Gráfico de Pizza / Donut de Distribuição */}
            <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-5 shadow-xl flex flex-col justify-between">
              <div className="flex items-center justify-between mb-2 border-b border-[#23283c] pb-3">
                <div className="flex items-center gap-2">
                  <PieChartIcon className="w-4 h-4 text-purple-400" />
                  <h2 className="text-sm font-bold text-white">
                    Distribuição do Trilho
                  </h2>
                </div>
              </div>

              <div className="h-56 w-full relative flex items-center justify-center">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={dadosGraficoStatus}
                      cx="50%"
                      cy="50%"
                      innerRadius={55}
                      outerRadius={80}
                      paddingAngle={4}
                      dataKey="value"
                    >
                      {dadosGraficoStatus.map((entry, index) => (
                        <Cell key={`cell-status-${index}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip
                      content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          const p = payload[0];
                          const pct = metricasAtuais.totalMembros > 0
                            ? Math.round(((p.value as number) / metricasAtuais.totalMembros) * 100)
                            : 0;
                          return (
                            <div className="bg-[#0f111a] border border-[#2e3650] p-2.5 rounded-xl shadow-xl text-xs">
                              <p className="font-bold text-white">{p.name}</p>
                              <p className="text-indigo-300">{p.value} membros ({pct}%)</p>
                            </div>
                          );
                        }
                        return null;
                      }}
                    />
                  </PieChart>
                </ResponsiveContainer>
                {/* Texto Central no Donut */}
                <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                  <span className="text-xl font-black text-white font-mono">
                    {metricasAtuais.totalMembros}
                  </span>
                  <span className="text-[10px] text-slate-400 font-medium">Membros</span>
                </div>
              </div>

              {/* Legenda do Donut */}
              <div className="space-y-2 pt-2 border-t border-[#23283c]">
                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-3 h-3 rounded-full bg-emerald-500" />
                    <span className="text-slate-300">Completo (100%)</span>
                  </div>
                  <span className="font-bold text-white font-mono">
                    {metricasAtuais.membrosCompletos} ({metricasAtuais.totalMembros > 0 ? Math.round((metricasAtuais.membrosCompletos / metricasAtuais.totalMembros) * 100) : 0}%)
                  </span>
                </div>

                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-3 h-3 rounded-full bg-amber-500" />
                    <span className="text-slate-300">Em Andamento</span>
                  </div>
                  <span className="font-bold text-white font-mono">
                    {metricasAtuais.membrosEmAndamento} ({metricasAtuais.totalMembros > 0 ? Math.round((metricasAtuais.membrosEmAndamento / metricasAtuais.totalMembros) * 100) : 0}%)
                  </span>
                </div>

                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-3 h-3 rounded-full bg-slate-500" />
                    <span className="text-slate-300">Não Iniciado</span>
                  </div>
                  <span className="font-bold text-white font-mono">
                    {metricasAtuais.membrosNaoIniciados} ({metricasAtuais.totalMembros > 0 ? Math.round((metricasAtuais.membrosNaoIniciados / metricasAtuais.totalMembros) * 100) : 0}%)
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ABA 2: COMPARATIVO (POR ÁREA / POR SETOR) */}
      {abaAtiva === 'comparativo' && (
        <div className="space-y-6">
          {/* Comparativo de Áreas */}
          <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-5 shadow-xl">
            <div className="flex items-center justify-between mb-4 border-b border-[#23283c] pb-3">
              <div className="flex items-center gap-2">
                <Compass className="w-4 h-4 text-cyan-400" />
                <h2 className="text-sm font-bold text-white">
                  Ranking & Comparativo de Conclusão por Área
                </h2>
              </div>
              <span className="text-xs text-slate-400">
                {dadosComparativoAreas.length} Áreas cadastradas
              </span>
            </div>

            <div className="h-72 w-full mb-6">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={dadosComparativoAreas}
                  margin={{ top: 15, right: 20, left: -10, bottom: 20 }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke="#252b40" />
                  <XAxis dataKey="area" stroke="#94a3b8" fontSize={11} tickLine={false} />
                  <YAxis stroke="#94a3b8" fontSize={11} unit="%" domain={[0, 100]} tickLine={false} />
                  <Tooltip
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const item = payload[0].payload;
                        return (
                          <div className="bg-[#0f111a] border border-[#2e3650] p-3 rounded-xl shadow-2xl text-xs space-y-1">
                            <p className="font-bold text-white text-sm">{item.area}</p>
                            <p className="text-cyan-400 font-bold">Conclusão Média: {item.percentualMedio}%</p>
                            <p className="text-slate-300">Total de Membros: {item.totalMembros}</p>
                            <p className="text-emerald-400">Trilho Completo: {item.completos}</p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar dataKey="percentualMedio" fill="#06b6d4" radius={[8, 8, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
              {dadosComparativoAreas.map((item, idx) => (
                <div
                  key={item.area}
                  className="bg-[#10121d] border border-[#242a3e] rounded-xl p-3.5 space-y-2 hover:border-cyan-500/40 transition-all cursor-pointer"
                  onClick={() => {
                    setVisao('area');
                    setAreaSelecionada(item.area);
                    setAbaAtiva('graficos');
                  }}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-white truncate">{item.area}</span>
                    <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-cyan-950 text-cyan-300">
                      #{idx + 1}
                    </span>
                  </div>
                  <div className="flex items-baseline justify-between">
                    <span className="text-xl font-black text-cyan-400 font-mono">
                      {item.percentualMedio}%
                    </span>
                    <span className="text-[11px] text-slate-400">
                      {item.totalMembros} membros
                    </span>
                  </div>
                  <div className="w-full bg-[#1b1f30] rounded-full h-1.5 overflow-hidden">
                    <div
                      className="bg-cyan-400 h-full rounded-full"
                      style={{ width: `${item.percentualMedio}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Comparativo de Setores */}
          <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-5 shadow-xl">
            <div className="flex items-center justify-between mb-4 border-b border-[#23283c] pb-3">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-indigo-400" />
                <h2 className="text-sm font-bold text-white">
                  Ranking & Comparativo de Conclusão por Setor
                </h2>
              </div>
              <span className="text-xs text-slate-400">
                {dadosComparativoSetores.length} Setores cadastrados
              </span>
            </div>

            <div className="h-72 w-full mb-6">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={dadosComparativoSetores.slice(0, 15)}
                  margin={{ top: 15, right: 20, left: -10, bottom: 25 }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke="#252b40" />
                  <XAxis
                    dataKey="setor"
                    stroke="#94a3b8"
                    fontSize={11}
                    tickLine={false}
                    angle={-20}
                    textAnchor="end"
                  />
                  <YAxis stroke="#94a3b8" fontSize={11} unit="%" domain={[0, 100]} tickLine={false} />
                  <Tooltip
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const item = payload[0].payload;
                        return (
                          <div className="bg-[#0f111a] border border-[#2e3650] p-3 rounded-xl shadow-2xl text-xs space-y-1">
                            <p className="font-bold text-white text-sm">{item.setor}</p>
                            <p className="text-indigo-400 font-bold">Conclusão Média: {item.percentualMedio}%</p>
                            <p className="text-slate-300">Total de Membros: {item.totalMembros}</p>
                            <p className="text-emerald-400">Trilho Completo: {item.completos}</p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar dataKey="percentualMedio" fill="#6366f1" radius={[8, 8, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-3">
              {dadosComparativoSetores.map((item, idx) => (
                <div
                  key={item.setor}
                  className="bg-[#10121d] border border-[#242a3e] rounded-xl p-3 space-y-1.5 hover:border-indigo-500/40 transition-all cursor-pointer"
                  onClick={() => {
                    setVisao('setor');
                    setSetorSelecionado(item.setor);
                    setAbaAtiva('graficos');
                  }}
                >
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-white truncate">{item.setor}</span>
                    <span className="text-[9px] font-bold px-1 rounded bg-indigo-950 text-indigo-300">
                      #{idx + 1}
                    </span>
                  </div>
                  <div className="flex items-baseline justify-between">
                    <span className="text-lg font-black text-indigo-300 font-mono">
                      {item.percentualMedio}%
                    </span>
                    <span className="text-[10px] text-slate-400">
                      {item.totalMembros}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ABA 3: LISTA & RASTREAMENTO INDIVIDUAL DE MEMBROS */}
      {abaAtiva === 'membros' && (
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl overflow-hidden shadow-xl space-y-4 p-5">
          {/* Barra de Filtros da Tabela */}
          <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 border-b border-[#23283c] pb-4">
            <div className="relative flex-1 min-w-[240px]">
              <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={buscaMembro}
                onChange={(e) => setBuscaMembro(e.target.value)}
                placeholder="Buscar por nome, login, setor, cargo ou célula..."
                className="w-full bg-[#10121d] border border-[#2d354e] rounded-xl pl-10 pr-4 py-2.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-indigo-400 font-medium"
              />
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              {/* Filtro de Status */}
              <select
                value={filtroStatusMembro}
                onChange={(e) => setFiltroStatusMembro(e.target.value)}
                className="bg-[#10121d] border border-[#2d354e] text-white text-xs rounded-xl px-3 py-2.5 font-medium focus:outline-none focus:border-indigo-400 cursor-pointer"
              >
                <option value="todos">Todos os Status</option>
                <option value="Completo">Apenas Trilho Completo (100%)</option>
                <option value="Em Andamento">Em Andamento</option>
                <option value="Não Iniciado">Não Iniciados</option>
              </select>

              {/* Filtro de Quem Falta Fazer Determinada Etapa */}
              <select
                value={filtroEtapaPendente}
                onChange={(e) => setFiltroEtapaPendente(e.target.value)}
                className="bg-[#10121d] border border-[#2d354e] text-white text-xs rounded-xl px-3 py-2.5 font-medium focus:outline-none focus:border-indigo-400 cursor-pointer"
              >
                <option value="todos">Filtro de Pendências (Todas Etapas)</option>
                {etapas.map(e => (
                  <option key={e} value={e}>Quem FALTA fazer: {e}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Tabela de Membros */}
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-300">
              <thead className="bg-[#10121d] text-slate-400 uppercase text-[10px] font-bold border-b border-[#23283c]">
                <tr>
                  <th className="py-3 px-4">Membro</th>
                  <th className="py-3 px-3">Setor / Área</th>
                  <th className="py-3 px-3">Cargo / Célula</th>
                  <th className="py-3 px-3">Progresso</th>
                  {etapas.map(etapa => (
                    <th key={etapa} className="py-3 px-2 text-center truncate max-w-[100px]" title={etapa}>
                      {etapa}
                    </th>
                  ))}
                  <th className="py-3 px-3 text-right">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#202538]">
                {membrosTabela.length === 0 ? (
                  <tr>
                    <td colSpan={5 + etapas.length} className="text-center py-8 text-slate-500">
                      Nenhum membro encontrado com os filtros selecionados.
                    </td>
                  </tr>
                ) : (
                  membrosTabela.slice(0, 100).map((m) => (
                    <tr key={m.id} className="hover:bg-[#1a1e30] transition-colors">
                      <td className="py-3 px-4 font-bold text-white">
                        <div className="flex flex-col">
                          <span>{m.nome}</span>
                          <span className="text-[10px] text-slate-400 font-mono">@{m.login || 'sem-login'}</span>
                        </div>
                      </td>
                      <td className="py-3 px-3">
                        <div className="flex flex-col">
                          <span className="text-slate-200 font-medium">{m.setor}</span>
                          <span className="text-[10px] text-slate-400">{m.area}</span>
                        </div>
                      </td>
                      <td className="py-3 px-3">
                        <div className="flex flex-col">
                          <span className="text-slate-200">{m.cargo}</span>
                          <span className="text-[10px] text-slate-400">{m.celula || '-'}</span>
                        </div>
                      </td>
                      <td className="py-3 px-3 min-w-[120px]">
                        <div className="space-y-1">
                          <div className="flex justify-between text-[11px] font-mono">
                            <span className="text-slate-300">{m.totalConcluidas}/{m.totalEtapas}</span>
                            <span className="font-bold text-indigo-300">{m.percentualConclusao}%</span>
                          </div>
                          <div className="w-full bg-[#11131e] rounded-full h-1.5 overflow-hidden">
                            <div
                              className={`h-full rounded-full ${
                                m.statusTrilho === 'Completo'
                                  ? 'bg-emerald-400'
                                  : m.statusTrilho === 'Em Andamento'
                                  ? 'bg-indigo-500'
                                  : 'bg-slate-600'
                              }`}
                              style={{ width: `${m.percentualConclusao}%` }}
                            />
                          </div>
                        </div>
                      </td>
                      {etapas.map(etapa => {
                        const concluido = m.etapasConcluidas.includes(etapa);
                        return (
                          <td key={etapa} className="py-3 px-2 text-center">
                            {concluido ? (
                              <span
                                className="inline-flex p-1 rounded-full bg-emerald-950 text-emerald-400 border border-emerald-500/40"
                                title={`${m.nome} concluiu ${etapa}`}
                              >
                                <CheckCircle2 className="w-3.5 h-3.5" />
                              </span>
                            ) : (
                              <span
                                className="inline-flex p-1 rounded-full bg-[#11131e] text-slate-600 border border-[#23283c]"
                                title={`${m.nome} ainda não concluiu ${etapa}`}
                              >
                                <Clock className="w-3.5 h-3.5" />
                              </span>
                            )}
                          </td>
                        );
                      })}
                      <td className="py-3 px-3 text-right">
                        <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wider ${
                          m.statusTrilho === 'Completo'
                            ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/40'
                            : m.statusTrilho === 'Em Andamento'
                            ? 'bg-amber-950 text-amber-300 border border-amber-500/40'
                            : 'bg-slate-800 text-slate-400 border border-slate-700'
                        }`}>
                          {m.statusTrilho}
                        </span>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          {membrosTabela.length > 100 && (
            <div className="text-center py-2 text-xs text-slate-400 border-t border-[#23283c]">
              Exibindo os primeiros 100 membros de {membrosTabela.length} filtrados. Use a busca ou exporte o CSV completo.
            </div>
          )}
        </div>
      )}
    </div>
  );
};
