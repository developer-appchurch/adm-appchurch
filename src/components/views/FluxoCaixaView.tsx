import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { 
  TrendingUp, 
  TrendingDown, 
  DollarSign, 
  Calendar, 
  BarChart3, 
  Search, 
  FileSpreadsheet, 
  FileText, 
  ArrowUpRight, 
  ArrowDownRight, 
  RotateCcw,
  Wallet,
  CheckCircle2,
  PlusCircle,
  Plus,
  Edit2,
  Trash2,
  AlertTriangle,
  RefreshCw
} from 'lucide-react';
import { 
  ResponsiveContainer, 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  Tooltip, 
  CartesianGrid 
} from 'recharts';
import { MembroItem, MovimentacaoFluxoCaixa } from '../../types';
import { MOVIMENTACOES_INICIAIS_FLUXO_CAIXA } from '../../data/mockFluxoCaixaData';
import { ExportService } from '../../services/exportService';
import { NovoFluxoModal } from './NovoFluxoModal';

const MESES_NOMES = [
  { valor: 'todos', label: 'Todos os Meses', abrev: 'TODOS' },
  { valor: '1', label: 'Janeiro', abrev: 'JAN' },
  { valor: '2', label: 'Fevereiro', abrev: 'FEV' },
  { valor: '3', label: 'Março', abrev: 'MAR' },
  { valor: '4', label: 'Abril', abrev: 'ABR' },
  { valor: '5', label: 'Maio', abrev: 'MAI' },
  { valor: '6', label: 'Junho', abrev: 'JUN' },
  { valor: '7', label: 'Julho', abrev: 'JUL' },
  { valor: '8', label: 'Agosto', abrev: 'AGO' },
  { valor: '9', label: 'Setembro', abrev: 'SET' },
  { valor: '10', label: 'Outubro', abrev: 'OUT' },
  { valor: '11', label: 'Novembro', abrev: 'NOV' },
  { valor: '12', label: 'Dezembro', abrev: 'DEZ' }
];

interface FluxoCaixaViewProps {
  anoSelecionado?: number;
  onSelectAno?: (ano: number) => void;
  onRefresh?: () => void;
  onAtualizarDados?: () => void;
  usuarioLogado?: MembroItem | null;
}

export const FluxoCaixaView: React.FC<FluxoCaixaViewProps> = ({
  anoSelecionado = 2026,
  onSelectAno,
  onRefresh,
  usuarioLogado
}) => {
  const [movimentacoes, setMovimentacoes] = useState<MovimentacaoFluxoCaixa[]>(MOVIMENTACOES_INICIAIS_FLUXO_CAIXA);
  const [loading, setLoading] = useState<boolean>(false);
  const [isNovoFluxoModalOpen, setIsNovoFluxoModalOpen] = useState<boolean>(false);
  const [itemEmEdicao, setItemEmEdicao] = useState<MovimentacaoFluxoCaixa | null>(null);
  const [itemParaExcluir, setItemParaExcluir] = useState<MovimentacaoFluxoCaixa | null>(null);
  const [isExcluindo, setIsExcluindo] = useState<boolean>(false);
  const [membrosMap, setMembrosMap] = useState<Record<string, string>>({
    '4': 'Junio Fonteles'
  });
  
  // Filtros principais
  const [ano, setAno] = useState<number>(anoSelecionado);
  const [mesSelecionado, setMesSelecionado] = useState<string>('todos');
  const [filtroTipo, setFiltroTipo] = useState<'todos' | 'ENTRADA' | 'SAIDA'>('todos');
  const [filtroCategoria, setFiltroCategoria] = useState<string>('todos');
  const [buscaTexto, setBuscaTexto] = useState<string>('');
  
  // Paginação
  const [paginaAtual, setPaginaAtual] = useState<number>(1);
  const itensPorPagina = 12;

  // Sincroniza ano externo caso fornecido
  useEffect(() => {
    if (anoSelecionado) {
      setAno(anoSelecionado);
    }
  }, [anoSelecionado]);

  // Carrega movimentações reais da tabela BD_FluxoCaixa
  const carregarFluxoCaixa = useCallback(async (force = false) => {
    setLoading(true);
    try {
      const resp = await fetch(`/api/sharepoint/fluxo-caixa?force=${force ? 'true' : 'false'}&_t=${Date.now()}`, {
        headers: { 'Cache-Control': 'no-cache' }
      });
      if (resp.ok) {
        const data = await resp.json();
        if (data && Array.isArray(data.movimentacoes)) {
          setMovimentacoes(data.movimentacoes);
          return;
        }
      }
      setMovimentacoes([]);
    } catch (err) {
      console.warn('[FluxoCaixaView] Aviso ao carregar BD_FluxoCaixa:', err);
      setMovimentacoes([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    carregarFluxoCaixa();
  }, [carregarFluxoCaixa]);

  // Carrega mapa de membros da tabela BD_membros para resolução de nomes
  useEffect(() => {
    fetch('/api/sharepoint/membros')
      .then(r => r.json())
      .then(d => {
        if (d && Array.isArray(d.membros)) {
          const map: Record<string, string> = { '4': 'Junio Fonteles' };
          d.membros.forEach((mb: any) => {
            const id = String(mb.id || mb.ID || mb.Title || '').trim();
            const nome = String(mb.nome || mb.Nome || mb.Title || '').trim();
            if (id && nome && isNaN(Number(nome))) {
              map[id] = nome;
            }
          });
          setMembrosMap(prev => ({ ...prev, ...map }));
        }
      })
      .catch(() => {});
  }, []);

  // Formatação em Reais (BRL)
  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(val || 0);
  };

  // Movimentações do Ano Selecionado
  const movimentacoesAno = useMemo(() => {
    return movimentacoes.filter(m => Number(m.ano) === Number(ano));
  }, [movimentacoes, ano]);

  // Categorias disponíveis para filtro extraídas dos cadastros reais
  const categoriasDisponiveis = useMemo(() => {
    const cats = new Set<string>();
    movimentacoes.forEach(m => {
      const cat = m.CategoriaFluxo || m.categoria;
      if (cat) cats.add(cat);
    });
    return Array.from(cats).sort();
  }, [movimentacoes]);

  // Métricas do Saldo Atual e do Mês exclusivamente calculadas a partir dos fluxos cadastrados
  const metricas = useMemo(() => {
    // 1. Saldo Geral / Saldo Atual (Consolidado de todas as movimentações reais cadastradas)
    const totalEntradasGeral = movimentacoes
      .filter(m => (m.tipo === 'ENTRADA' || m.CategoriaFluxo === 'Entrada'))
      .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

    const totalSaidasGeral = movimentacoes
      .filter(m => (m.tipo === 'SAIDA' || m.CategoriaFluxo === 'Saída'))
      .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

    const saldoGeral = totalEntradasGeral - totalSaidasGeral;

    // 2. Métricas do Mês Selecionado (ou do ano todo se 'todos')
    const numMes = mesSelecionado === 'todos' ? null : Number(mesSelecionado);
    const movsMes = movimentacoesAno.filter(m => numMes === null || Number(m.mes) === numMes);

    const entradasMes = movsMes
      .filter(m => (m.tipo === 'ENTRADA' || m.CategoriaFluxo === 'Entrada'))
      .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

    const saidasMes = movsMes
      .filter(m => (m.tipo === 'SAIDA' || m.CategoriaFluxo === 'Saída'))
      .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

    const saldoMes = entradasMes - saidasMes;

    return {
      saldoGeral,
      totalEntradasGeral,
      totalSaidasGeral,
      entradasMes,
      saidasMes,
      saldoMes,
      qtdMovsMes: movsMes.length
    };
  }, [movimentacoes, movimentacoesAno, mesSelecionado]);

  // Dados para o Gráfico de Barras Mensal (Janeiro a Dezembro) exclusivamente dos cadastros
  const dadosGraficoMensal = useMemo(() => {
    const meses = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const nomesAbrev = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
    const nomesCompletos = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

    return meses.map((mNum, idx) => {
      const movsDoMes = movimentacoesAno.filter(m => Number(m.mes) === mNum);
      const entradas = movsDoMes
        .filter(m => (m.tipo === 'ENTRADA' || m.CategoriaFluxo === 'Entrada'))
        .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

      const saidas = movsDoMes
        .filter(m => (m.tipo === 'SAIDA' || m.CategoriaFluxo === 'Saída'))
        .reduce((acc, m) => acc + (m.ValorFluxo ?? m.valor ?? 0), 0);

      const saldo = entradas - saidas;

      return {
        mesNumero: mNum,
        mesAbrev: nomesAbrev[idx],
        mesNome: nomesCompletos[idx],
        entradas,
        saidas,
        saldo
      };
    });
  }, [movimentacoesAno]);

  // Histórico Filtrado para a Tabela
  const movimentacoesFiltradas = useMemo(() => {
    return movimentacoes.filter(item => {
      // Filtro de ano
      if (item.ano && Number(item.ano) !== Number(ano)) return false;

      // Filtro de mês
      if (mesSelecionado !== 'todos' && Number(item.mes) !== Number(mesSelecionado)) return false;

      // Filtro de tipo (ENTRADA / SAIDA)
      if (filtroTipo !== 'todos') {
        const itemTipo = (item.tipo || (item.CategoriaFluxo === 'Saída' ? 'SAIDA' : 'ENTRADA'));
        if (itemTipo !== filtroTipo) return false;
      }

      // Filtro de categoria
      if (filtroCategoria !== 'todos') {
        const cat = item.CategoriaFluxo || item.categoria;
        if (cat !== filtroCategoria) return false;
      }

      // Busca textual por descrição, tipo, etc.
      if (buscaTexto.trim() !== '') {
        const query = buscaTexto.toLowerCase().trim();
        const desc = (item.ObservacoesFluxo || item.ObservacaoFluxo || item.DescricaoFluxo || item.descricao || item.observacao || '').toLowerCase();
        const cat = (item.CategoriaFluxo || item.categoria || '').toLowerCase();
        const tp = (item.TipoFluxo || item.formaPagamento || '').toLowerCase();
        const status = (item.StatusFluxo || item.status || '').toLowerCase();
        const match = desc.includes(query) || cat.includes(query) || tp.includes(query) || status.includes(query);
        if (!match) return false;
      }

      return true;
    }).sort((a, b) => {
      const dataA = a.DataFluxo || a.data || '';
      const dataB = b.DataFluxo || b.data || '';
      return new Date(dataB).getTime() - new Date(dataA).getTime();
    });
  }, [movimentacoes, ano, mesSelecionado, filtroTipo, filtroCategoria, buscaTexto]);

  // Paginação
  const totalPaginas = Math.ceil(movimentacoesFiltradas.length / itensPorPagina) || 1;
  const movimentacoesPaginadas = useMemo(() => {
    const inicio = (paginaAtual - 1) * itensPorPagina;
    return movimentacoesFiltradas.slice(inicio, inicio + itensPorPagina);
  }, [movimentacoesFiltradas, paginaAtual, itensPorPagina]);

  const handleMudarAno = (novoAno: number) => {
    setAno(novoAno);
    setPaginaAtual(1);
    if (onSelectAno) onSelectAno(novoAno);
  };

  const handleLimparFiltros = () => {
    setFiltroTipo('todos');
    setFiltroCategoria('todos');
    setBuscaTexto('');
    setPaginaAtual(1);
  };

  const handleNovoLancamento = () => {
    setItemEmEdicao(null);
    setIsNovoFluxoModalOpen(true);
  };

  const handleEditar = (item: MovimentacaoFluxoCaixa) => {
    setItemEmEdicao(item);
    setIsNovoFluxoModalOpen(true);
  };

  const handleNovoFluxoSalvo = (fluxoSalvo: MovimentacaoFluxoCaixa) => {
    setMovimentacoes(prev => {
      const idx = prev.findIndex(m => 
        String(m.id) === String(fluxoSalvo.id) || 
        String(m.ID) === String(fluxoSalvo.ID) || 
        String(m.Title) === String(fluxoSalvo.Title)
      );
      if (idx >= 0) {
        const copy = [...prev];
        copy[idx] = fluxoSalvo;
        return copy;
      }
      return [fluxoSalvo, ...prev];
    });
    setItemEmEdicao(null);
    carregarFluxoCaixa(true);
  };

  const handleConfirmarExcluir = async () => {
    if (!itemParaExcluir) return;
    setIsExcluindo(true);
    try {
      const idParaRemover = itemParaExcluir.id || itemParaExcluir.ID;
      const res = await fetch(`/api/sharepoint/fluxo-caixa/${idParaRemover}`, {
        method: 'DELETE'
      });
      if (res.ok) {
        setMovimentacoes(prev => prev.filter(m => 
          String(m.id) !== String(idParaRemover) && 
          String(m.ID) !== String(idParaRemover) && 
          String(m.Title) !== String(idParaRemover)
        ));
        setItemParaExcluir(null);
        carregarFluxoCaixa(true);
      } else {
        const data = await res.json().catch(() => null);
        alert(data?.erro || 'Erro ao excluir o lançamento.');
      }
    } catch (err) {
      console.error('[FluxoCaixaView] Erro ao excluir:', err);
      alert('Falha de conexão ao excluir o lançamento.');
    } finally {
      setIsExcluindo(false);
    }
  };

  const exportarExcel = () => {
    ExportService.exportarFluxoCaixaExcel(movimentacoesFiltradas, { ano, mes: mesSelecionado }, `Fluxo_Caixa_${ano}_Mes_${mesSelecionado}`);
  };

  const exportarPDF = () => {
    ExportService.exportarFluxoCaixaPDF(movimentacoesFiltradas, { ano, mes: mesSelecionado }, `Fluxo_Caixa_${ano}`);
  };

  return (
    <div id="fluxo-caixa-container" className="p-3.5 sm:p-6 space-y-5 max-w-[1600px] mx-auto text-slate-100">
      
      {/* Modal de Cadastro / Edição de Lançamento */}
      <NovoFluxoModal
        isOpen={isNovoFluxoModalOpen}
        onClose={() => {
          setIsNovoFluxoModalOpen(false);
          setItemEmEdicao(null);
        }}
        usuarioLogado={usuarioLogado}
        itemParaEditar={itemEmEdicao}
        onSalvoComSucesso={handleNovoFluxoSalvo}
      />

      {/* Modal de Confirmação de Exclusão */}
      {itemParaExcluir && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-150">
          <div className="bg-[#181c2b] border border-rose-500/40 rounded-2xl w-full max-w-md shadow-2xl overflow-hidden p-5 text-slate-100 animate-in zoom-in-95">
            <div className="flex items-center gap-3 text-rose-400 mb-3">
              <div className="w-10 h-10 rounded-xl bg-rose-950/80 border border-rose-500/50 flex items-center justify-center shrink-0">
                <AlertTriangle className="w-5 h-5 text-rose-400" />
              </div>
              <div>
                <h3 className="text-base font-bold text-white">Excluir Lançamento</h3>
                <p className="text-xs text-slate-400">Esta ação removerá o registro do sistema e do SharePoint.</p>
              </div>
            </div>

            <div className="bg-[#111420] border border-[#272f47] rounded-xl p-3.5 my-4 space-y-1.5 text-xs">
              <div className="flex justify-between">
                <span className="text-slate-400">ID:</span>
                <span className="font-mono font-bold text-indigo-400">#{itemParaExcluir.ID || itemParaExcluir.id}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">Valor:</span>
                <span className="font-mono font-bold text-white">{formatBRL(itemParaExcluir.ValorFluxo ?? itemParaExcluir.valor ?? 0)}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">Categoria:</span>
                <span className="font-bold text-slate-200">{itemParaExcluir.CategoriaFluxo || itemParaExcluir.categoria}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">Descrição:</span>
                <span className="text-slate-300 font-medium truncate max-w-[200px]">{itemParaExcluir.ObservacoesFluxo || itemParaExcluir.ObservacaoFluxo || itemParaExcluir.DescricaoFluxo || itemParaExcluir.descricao}</span>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setItemParaExcluir(null)}
                disabled={isExcluindo}
                className="px-4 py-2 rounded-xl bg-[#202538] hover:bg-[#2c334d] text-slate-300 font-bold text-xs transition-colors cursor-pointer border border-[#303954]"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={handleConfirmarExcluir}
                disabled={isExcluindo}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white font-bold text-xs shadow-lg shadow-rose-600/30 transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-60"
              >
                {isExcluindo ? (
                  <>
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>Excluindo...</span>
                  </>
                ) : (
                  <>
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Confirmar Exclusão</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Barra Superior de Seleção de Período e Botão Novo Lançamento */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-[#161a29] border border-[#272d42] p-3.5 sm:p-4 rounded-2xl shadow-lg">
        <div className="flex items-center gap-2.5">
          <div className="w-10 h-10 rounded-xl bg-indigo-600/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400">
            <Wallet className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-base sm:text-lg font-bold text-white tracking-tight">
              Fluxo de Caixa
            </h1>
            <p className="text-xs text-slate-400">
              Controle de Entradas, Saídas e Demonstrativo Financeiro
            </p>
          </div>
        </div>

        {/* Seletores de Ano e Mês + Botão Novo Lançamento */}
        <div className="flex items-center flex-wrap gap-2 sm:gap-3 ml-auto">
          {/* Botão Novo Lançamento */}
          <button
            id="btn-novo-lancamento"
            onClick={handleNovoLancamento}
            className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs shadow-md shadow-indigo-600/30 transition-all cursor-pointer"
            title="Cadastrar Novo Lançamento"
          >
            <PlusCircle className="w-4 h-4" />
            <span>Novo Lançamento</span>
          </button>

          {/* Seletor de Ano */}
          <div className="flex items-center gap-1.5 bg-[#202538] px-3 py-1.5 rounded-xl border border-[#313955]">
            <span className="text-xs text-slate-400 font-medium">Ano:</span>
            <select
              value={ano}
              onChange={(e) => handleMudarAno(Number(e.target.value))}
              className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
            >
              <option value={2026} className="bg-[#181c2b] text-white">2026</option>
              <option value={2025} className="bg-[#181c2b] text-white">2025</option>
              <option value={2024} className="bg-[#181c2b] text-white">2024</option>
            </select>
          </div>

          {/* Seletor de Mês */}
          <div className="flex items-center gap-1.5 bg-[#202538] px-3 py-1.5 rounded-xl border border-[#313955]">
            <span className="text-xs text-slate-400 font-medium">Mês:</span>
            <select
              value={mesSelecionado}
              onChange={(e) => {
                setMesSelecionado(e.target.value);
                setPaginaAtual(1);
              }}
              className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
            >
              {MESES_NOMES.map(m => (
                <option key={m.valor} value={m.valor} className="bg-[#181c2b] text-white">
                  {m.label}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* CARDS NO TOPO (KPIS): Saldo Atual, Entrada Mês, Saída Mês, Saldo Mês */}
      <div id="fluxo-caixa-kpis" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        {/* Card 1: Saldo Geral / Saldo Atual */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Saldo Atual
              </span>
              <div className="p-2 rounded-xl bg-indigo-500/15 text-indigo-400">
                <DollarSign className="w-4 h-4" />
              </div>
            </div>
            <div className={`mt-2 text-2xl sm:text-3xl font-black font-mono tracking-tight ${
              metricas.saldoGeral >= 0 ? 'text-emerald-400' : 'text-rose-400'
            }`}>
              {formatBRL(metricas.saldoGeral)}
            </div>
          </div>
          <div className="mt-3 pt-2.5 border-t border-[#23283c] flex items-center justify-between text-xs">
            <span className="text-slate-400">Situação:</span>
            <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
              metricas.saldoGeral >= 0 
                ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30' 
                : 'bg-rose-950 text-rose-300 border border-rose-500/30'
            }`}>
              {metricas.saldoGeral >= 0 ? 'SUPERÁVIT ACUMULADO' : 'DÉFICIT ACUMULADO'}
            </span>
          </div>
        </div>

        {/* Card 2: Entradas do Mês */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Entradas do Mês
              </span>
              <div className="p-2 rounded-xl bg-emerald-500/15 text-emerald-400">
                <TrendingUp className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 text-2xl sm:text-3xl font-black font-mono tracking-tight text-emerald-400">
              {formatBRL(metricas.entradasMes)}
            </div>
          </div>
          <div className="mt-3 pt-2.5 border-t border-[#23283c] flex items-center justify-between text-xs">
            <span className="text-slate-400">
              {mesSelecionado === 'todos' ? 'Ano Completo' : MESES_NOMES.find(m => m.valor === mesSelecionado)?.label}:
            </span>
            <span className="text-emerald-300 font-semibold flex items-center gap-0.5">
              <ArrowUpRight className="w-3.5 h-3.5" />
              Entradas
            </span>
          </div>
        </div>

        {/* Card 3: Saídas do Mês */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Saídas do Mês
              </span>
              <div className="p-2 rounded-xl bg-rose-500/15 text-rose-400">
                <TrendingDown className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 text-2xl sm:text-3xl font-black font-mono tracking-tight text-rose-400">
              {formatBRL(metricas.saidasMes)}
            </div>
          </div>
          <div className="mt-3 pt-2.5 border-t border-[#23283c] flex items-center justify-between text-xs">
            <span className="text-slate-400">Comprometimento:</span>
            <span className="text-rose-300 font-semibold">
              {metricas.entradasMes > 0 
                ? `${((metricas.saidasMes / metricas.entradasMes) * 100).toFixed(1)}%` 
                : '0%'}
            </span>
          </div>
        </div>

        {/* Card 4: Saldo do Mês */}
        <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-lg relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                Saldo do Mês
              </span>
              <div className="p-2 rounded-xl bg-purple-500/15 text-purple-400">
                <BarChart3 className="w-4 h-4" />
              </div>
            </div>
            <div className={`mt-2 text-2xl sm:text-3xl font-black font-mono tracking-tight ${
              metricas.saldoMes >= 0 ? 'text-emerald-400' : 'text-rose-400'
            }`}>
              {formatBRL(metricas.saldoMes)}
            </div>
          </div>
          <div className="mt-3 pt-2.5 border-t border-[#23283c] flex items-center justify-between text-xs">
            <span className="text-slate-400">Resultado Líquido:</span>
            <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
              metricas.saldoMes >= 0 
                ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/30' 
                : 'bg-rose-950 text-rose-300 border border-rose-500/30'
            }`}>
              {metricas.saldoMes >= 0 ? 'POSITIVO' : 'NEGATIVO'}
            </span>
          </div>
        </div>

      </div>

      {/* GRÁFICO DE BARRAS (Entradas em Verde e Saídas em Vermelho por Mês) */}
      <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-xl space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-[#23283c] pb-3">
          <div className="flex items-center gap-2">
            <BarChart3 className="w-5 h-5 text-indigo-400" />
            <h2 className="text-sm sm:text-base font-bold text-white">
              Demonstrativo Mensal de Entradas e Saídas ({ano})
            </h2>
          </div>
          <div className="flex items-center gap-4 text-xs font-semibold">
            <span className="flex items-center gap-1.5 text-emerald-400">
              <span className="w-3 h-3 rounded-sm bg-emerald-500" />
              Entradas (+)
            </span>
            <span className="flex items-center gap-1.5 text-rose-400">
              <span className="w-3 h-3 rounded-sm bg-rose-500" />
              Saídas (-)
            </span>
          </div>
        </div>

        {/* Gráfico Recharts */}
        <div className="w-full h-72 sm:h-80">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart 
              data={dadosGraficoMensal}
              margin={{ top: 10, right: 10, left: 0, bottom: 5 }}
            >
              <CartesianGrid strokeDasharray="3 3" stroke="#252c42" vertical={false} />
              <XAxis 
                dataKey="mesAbrev" 
                stroke="#64748b" 
                tick={{ fill: '#94a3b8', fontSize: 11, fontWeight: 'bold' }} 
              />
              <YAxis 
                stroke="#64748b" 
                tick={{ fill: '#94a3b8', fontSize: 10 }}
                tickFormatter={(v) => `R$ ${(v / 1000).toFixed(0)}k`} 
              />
              <Tooltip 
                content={({ active, payload }) => {
                  if (active && payload && payload.length) {
                    const dataObj = payload[0].payload;
                    return (
                      <div className="bg-[#121522] border border-[#2c334d] p-3 rounded-xl shadow-2xl text-xs space-y-1.5">
                        <p className="font-bold text-white border-b border-[#252c42] pb-1">
                          {dataObj.mesNome} / {ano}
                        </p>
                        <div className="flex items-center justify-between gap-4 text-emerald-400">
                          <span>Entradas:</span>
                          <strong className="font-mono">{formatBRL(dataObj.entradas)}</strong>
                        </div>
                        <div className="flex items-center justify-between gap-4 text-rose-400">
                          <span>Saídas:</span>
                          <strong className="font-mono">{formatBRL(dataObj.saidas)}</strong>
                        </div>
                        <div className="flex items-center justify-between gap-4 pt-1 border-t border-[#252c42] text-slate-200">
                          <span>Resultado:</span>
                          <strong className={`font-mono ${dataObj.saldo >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                            {formatBRL(dataObj.saldo)}
                          </strong>
                        </div>
                      </div>
                    );
                  }
                  return null;
                }}
              />
              <Bar dataKey="entradas" fill="#10b981" radius={[4, 4, 0, 0]} name="Entradas" />
              <Bar dataKey="saidas" fill="#ef4444" radius={[4, 4, 0, 0]} name="Saídas" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* HISTÓRICO DE ÚLTIMAS MOVIMENTAÇÕES (EMBAIXO DO GRÁFICO COM FILTROS) */}
      <div className="bg-[#161a29] border border-[#272d42] rounded-2xl p-4 sm:p-5 shadow-xl space-y-4">
        
        {/* Cabeçalho do Histórico + Botões de Exportação e Novo Fluxo */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-[#23283c] pb-3">
          <div>
            <h3 className="text-base font-bold text-white flex items-center gap-2">
              <span>Histórico de Movimentações</span>
              <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                {movimentacoesFiltradas.length}
              </span>
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">
              Movimentações financeiras cadastradas
            </p>
          </div>

          <div className="flex items-center flex-wrap gap-2">
            <button
              onClick={() => setIsNovoFluxoModalOpen(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white shadow-sm transition-all cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Novo Fluxo</span>
            </button>

            <button
              onClick={exportarExcel}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-xl bg-emerald-800/80 hover:bg-emerald-700 text-white shadow-sm border border-emerald-500/40 transition-all cursor-pointer"
              title="Exportar para Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>

            <button
              onClick={exportarPDF}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-xl bg-rose-800/80 hover:bg-rose-700 text-white shadow-sm border border-rose-500/40 transition-all cursor-pointer"
              title="Exportar para PDF"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>PDF</span>
            </button>
          </div>
        </div>

        {/* Barra de Filtros Rápidos Integrada */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 bg-[#111420] p-3 rounded-xl border border-[#23283c] text-xs">
          
          {/* Busca textual */}
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={buscaTexto}
              onChange={(e) => {
                setBuscaTexto(e.target.value);
                setPaginaAtual(1);
              }}
              placeholder="Buscar descrição ou histórico..."
              className="w-full bg-[#181c2b] border border-[#2c334d] rounded-xl pl-9 pr-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-indigo-400"
            />
          </div>

          {/* Filtro Tipo */}
          <div>
            <select
              value={filtroTipo}
              onChange={(e) => {
                setFiltroTipo(e.target.value as any);
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181c2b] border border-[#2c334d] rounded-xl px-3 py-2 text-xs text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todos os Tipos (Entradas &amp; Saídas)</option>
              <option value="ENTRADA">Apenas Entradas (+)</option>
              <option value="SAIDA">Apenas Saídas (-)</option>
            </select>
          </div>

          {/* Filtro Categoria */}
          <div>
            <select
              value={filtroCategoria}
              onChange={(e) => {
                setFiltroCategoria(e.target.value);
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181c2b] border border-[#2c334d] rounded-xl px-3 py-2 text-xs text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todas as Categorias</option>
              {categoriasDisponiveis.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          {/* Botão Limpar Filtros */}
          <div className="flex items-center">
            <button
              onClick={handleLimparFiltros}
              className="w-full flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl bg-[#202538] hover:bg-[#2c334d] text-slate-300 hover:text-white transition-all text-xs font-semibold cursor-pointer border border-[#303852]"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Limpar Filtros
            </button>
          </div>
        </div>

        {/* Tabela de Movimentações */}
        <div className="border border-[#23283c] rounded-xl overflow-hidden overflow-x-auto bg-[#141724]">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="border-b border-[#23283c] bg-[#10131e] text-slate-400 text-[11px] uppercase tracking-wider font-semibold">
                <th className="py-3 px-4">ID</th>
                <th className="py-3 px-4">Data</th>
                <th className="py-3 px-4">Categoria</th>
                <th className="py-3 px-4">Tipo (Pagamento)</th>
                <th className="py-3 px-4">Descrição / Motivo</th>
                <th className="py-3 px-4">Tesoureiro</th>
                <th className="py-3 px-4">Status</th>
                <th className="py-3 px-4 text-right">Valor (R$)</th>
                <th className="py-3 px-4 text-center">Ações</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#1e2335]">
              {movimentacoesPaginadas.length === 0 ? (
                <tr>
                  <td colSpan={9} className="py-12 text-center text-slate-400">
                    <div className="flex flex-col items-center justify-center gap-2">
                      <Wallet className="w-8 h-8 text-slate-500 stroke-1" />
                      <p className="font-semibold text-sm text-slate-300">Nenhum fluxo de caixa registrado ainda.</p>
                      <p className="text-xs text-slate-500 max-w-sm">
                        Clique no botão "Novo Lançamento" para cadastrar sua primeira movimentação financeira.
                      </p>
                      <button
                        onClick={handleNovoLancamento}
                        className="mt-2 inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold transition-all shadow-md cursor-pointer"
                      >
                        <PlusCircle className="w-4 h-4" />
                        <span>Cadastrar Primeiro Lançamento</span>
                      </button>
                    </div>
                  </td>
                </tr>
              ) : (
                movimentacoesPaginadas.map((m, idx) => {
                  const isEntrada = m.CategoriaFluxo === 'Entrada' || m.tipo === 'ENTRADA';
                  const valorExibicao = m.ValorFluxo ?? m.valor ?? 0;
                  const dataExibicao = m.dataBR || (m.DataFluxo ? m.DataFluxo.split('-').reverse().join('/') : m.data);
                  const tipoPagamento = m.TipoFluxo || m.formaPagamento || 'Pix';
                  const descricaoExibicao = m.ObservacoesFluxo || m.ObservacaoFluxo || m.observacao || m.DescricaoFluxo || m.descricao || '-';
                  const statusExibicao = m.StatusFluxo || m.status || 'OK';
                  const idTesoureiroExibicao = String(m.Id_Tesoureiro || '4').trim();
                  const nomeTesoureiroExibicao = 
                    m.NomeTesoureiro || 
                    m.nome_tesoureiro || 
                    membrosMap[idTesoureiroExibicao] || 
                    (idTesoureiroExibicao === '4' ? 'Junio Fonteles' : (usuarioLogado?.nome || 'Junio Fonteles'));

                  return (
                    <tr key={`${m.id}-${idx}`} className="hover:bg-[#1b2031] transition-colors">
                      <td className="py-3 px-4 font-mono font-bold text-indigo-400 whitespace-nowrap">
                        #{m.ID || m.id}
                      </td>
                      <td className="py-3 px-4 font-mono text-slate-300 whitespace-nowrap">
                        {dataExibicao}
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap">
                        <span className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                          isEntrada
                            ? 'bg-emerald-950 text-emerald-300 border border-emerald-500/40'
                            : 'bg-rose-950 text-rose-300 border border-rose-500/40'
                        }`}>
                          {isEntrada ? (
                            <>
                              <ArrowUpRight className="w-3 h-3" />
                              Entrada
                            </>
                          ) : (
                            <>
                              <ArrowDownRight className="w-3 h-3" />
                              Saída
                            </>
                          )}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-slate-300 whitespace-nowrap">
                        <span className="px-2 py-0.5 rounded-lg bg-[#1e2337] border border-[#2e3654] text-[11px] text-slate-300 font-medium">
                          {tipoPagamento}
                        </span>
                      </td>
                      <td className="py-3 px-4 font-medium text-slate-100 min-w-[220px]">
                        <div>{descricaoExibicao}</div>
                      </td>
                      <td className="py-3 px-4 font-medium text-slate-200 whitespace-nowrap">
                        <span className="text-xs text-slate-200">{nomeTesoureiroExibicao}</span>
                      </td>
                      <td className="py-3 px-4 whitespace-nowrap">
                        <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-lg text-[10px] font-bold ${
                          statusExibicao === 'OK'
                            ? 'bg-emerald-900/30 text-emerald-300 border border-emerald-500/30'
                            : 'bg-amber-900/30 text-amber-300 border border-amber-500/30'
                        }`}>
                          {statusExibicao === 'OK' ? <CheckCircle2 className="w-3 h-3" /> : null}
                          {statusExibicao}
                        </span>
                      </td>
                      <td className={`py-3 px-4 text-right font-mono font-bold text-sm whitespace-nowrap ${
                        isEntrada ? 'text-emerald-400' : 'text-rose-400'
                      }`}>
                        {isEntrada ? `+ ${formatBRL(valorExibicao)}` : `- ${formatBRL(valorExibicao)}`}
                      </td>
                      <td className="py-3 px-4 text-center whitespace-nowrap">
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            onClick={() => handleEditar(m)}
                            className="p-1.5 rounded-lg bg-indigo-900/30 hover:bg-indigo-600/40 text-indigo-300 hover:text-white border border-indigo-500/30 transition-all cursor-pointer"
                            title="Editar Lançamento"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => setItemParaExcluir(m)}
                            className="p-1.5 rounded-lg bg-rose-900/30 hover:bg-rose-600/40 text-rose-300 hover:text-white border border-rose-500/30 transition-all cursor-pointer"
                            title="Excluir Lançamento"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Paginação da Tabela */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pt-2 text-xs text-slate-400">
          <div>
            Mostrando {movimentacoesPaginadas.length} de {movimentacoesFiltradas.length} movimentações registradas
          </div>
          {totalPaginas > 1 && (
            <div className="flex items-center gap-1.5">
              <button
                disabled={paginaAtual <= 1}
                onClick={() => setPaginaAtual(p => Math.max(1, p - 1))}
                className="px-3 py-1 rounded-lg bg-[#202538] hover:bg-[#2d344e] disabled:opacity-40 disabled:cursor-not-allowed text-white cursor-pointer border border-[#2f3752]"
              >
                Anterior
              </button>
              <span className="px-2 font-mono font-bold text-white">
                {paginaAtual} / {totalPaginas}
              </span>
              <button
                disabled={paginaAtual >= totalPaginas}
                onClick={() => setPaginaAtual(p => Math.min(totalPaginas, p + 1))}
                className="px-3 py-1 rounded-lg bg-[#202538] hover:bg-[#2d344e] disabled:opacity-40 disabled:cursor-not-allowed text-white cursor-pointer border border-[#2f3752]"
              >
                Próxima
              </button>
            </div>
          )}
        </div>

      </div>

    </div>
  );
};
