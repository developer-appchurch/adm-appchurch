import React, { useState, useMemo } from 'react';
import { 
  FileSpreadsheet, 
  FileText, 
  Filter, 
  TrendingUp, 
  TrendingDown, 
  DollarSign, 
  Calendar, 
  Layers, 
  MapPin, 
  BarChart3, 
  Search, 
  Database,
  Printer,
  ChevronDown,
  RotateCcw
} from 'lucide-react';
import { 
  LancamentoTesouraria, 
  FiltrosFluxoCaixa, 
  VisualizacaoAgrupamento, 
  SetorTipo, 
  AreaTipo 
} from '../../types';
import { 
  SETORES_DISPONIVEIS, 
  AREAS_DISPONIVEIS, 
  CATEGORIAS_ENTRADA, 
  CATEGORIAS_SAIDA 
} from '../../data/mockSharePointData';
import { ExportService } from '../../services/exportService';

interface FluxoCaixaViewProps {
  todosLancamentos?: LancamentoTesouraria[];
  lancamentos?: LancamentoTesouraria[];
  onAtualizarDados?: () => void;
  onRefresh?: () => void;
  anoSelecionado: number;
  onSelectAno?: (ano: number) => void;
}

export const FluxoCaixaView: React.FC<FluxoCaixaViewProps> = ({
  todosLancamentos,
  lancamentos,
  onAtualizarDados,
  onRefresh,
  anoSelecionado,
  onSelectAno
}) => {
  const listaLancamentos = useMemo(() => {
    if (Array.isArray(todosLancamentos) && todosLancamentos.length > 0) return todosLancamentos;
    if (Array.isArray(lancamentos) && lancamentos.length > 0) return lancamentos;
    return [];
  }, [todosLancamentos, lancamentos]);
  // Filtros
  const [filtros, setFiltros] = useState<FiltrosFluxoCaixa>({
    periodoPredefinido: 'todos',
    dataInicio: '',
    dataFim: '',
    ano: anoSelecionado,
    mes: 'todos',
    setor: 'todos',
    area: 'todos',
    celula: 'todos',
    tipo: 'todos',
    categoria: 'todos',
    metodo: 'todos',
    status: 'todos'
  });

  const [buscaTexto, setBuscaTexto] = useState('');
  const [agrupamentoAtivo, setAgrupamentoAtivo] = useState<VisualizacaoAgrupamento>('mensal');
  const [itensPorPagina] = useState(15);
  const [paginaAtual, setPaginaAtual] = useState(1);

  // Formatação monetária
  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(val || 0);
  };

  // Aplicação dos filtros sobre a base do SharePoint
  const lancamentosFiltrados = useMemo(() => {
    return listaLancamentos.filter(item => {
      // Filtro de ano se selecionado
      if (filtros.ano && item.ano !== Number(filtros.ano)) return false;

      // Filtro de mês se selecionado
      if (filtros.mes !== 'todos' && item.mes !== Number(filtros.mes)) return false;

      // Filtro de setor
      if (filtros.setor !== 'todos' && item.setor !== filtros.setor) return false;

      // Filtro de área
      if (filtros.area !== 'todos' && item.area !== filtros.area) return false;

      // Filtro de tipo
      if (filtros.tipo !== 'todos' && item.tipo !== filtros.tipo) return false;

      // Filtro de categoria
      if (filtros.categoria !== 'todos' && item.categoria !== filtros.categoria) return false;

      // Filtro de método
      if (filtros.metodo !== 'todos' && item.metodo !== filtros.metodo) return false;

      // Filtro de datas
      if (filtros.dataInicio && item.data < filtros.dataInicio) return false;
      if (filtros.dataFim && item.data > filtros.dataFim) return false;

      // Filtro de busca textual
      if (buscaTexto.trim() !== '') {
        const query = buscaTexto.toLowerCase();
        const match = 
          item.celulaNome.toLowerCase().includes(query) ||
          item.setor.toLowerCase().includes(query) ||
          item.categoria.toLowerCase().includes(query) ||
          item.descricao.toLowerCase().includes(query);
        if (!match) return false;
      }

      return true;
    });
  }, [listaLancamentos, filtros, buscaTexto]);

  // Cálculos de totais - REGRA: o total a ser mostrado é SOMENTE os validados pela coluna TESOURARIA_RECEB
  const metricas = useMemo(() => {
    // Apenas entradas validadas pela tesouraria (TESOURARIA_RECEB === true)
    const entradasValidadas = lancamentosFiltrados.filter(
      l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true
    );

    const totalEntradas = entradasValidadas.reduce(
      (acc, curr) => acc + (curr.Total ?? curr.valorTotal ?? 0), 0
    );

    const totalSaidas = lancamentosFiltrados
      .filter(l => l.tipo === 'SAIDA')
      .reduce((acc, curr) => acc + curr.valorTotal, 0);

    const totalPix = entradasValidadas.reduce(
      (acc, curr) => acc + (curr.ValorOferta ?? curr.valorPix ?? 0), 0
    );

    const totalEspecie = entradasValidadas.reduce(
      (acc, curr) => acc + (curr.OfertaEspecie ?? curr.valorEspecie ?? 0), 0
    );

    // Relatórios de entrada pendentes de validação pela tesouraria (TESOURARIA_RECEB !== true)
    const entradasPendentes = lancamentosFiltrados.filter(
      l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB !== true
    );
    const totalPendenteValidacao = entradasPendentes.reduce(
      (acc, curr) => acc + (curr.Total ?? curr.valorTotal ?? 0), 0
    );

    const saldoOperacional = totalEntradas - totalSaidas;
    const percPix = totalEntradas > 0 ? (totalPix / totalEntradas) * 100 : 0;
    const percEspecie = totalEntradas > 0 ? (totalEspecie / totalEntradas) * 100 : 0;

    return {
      totalEntradas,
      totalSaidas,
      saldoOperacional,
      totalPix,
      totalEspecie,
      totalPendenteValidacao,
      percPix,
      percEspecie,
      qtdLancamentos: lancamentosFiltrados.length,
      qtdValidados: entradasValidadas.length,
      qtdPendentes: entradasPendentes.length
    };
  }, [lancamentosFiltrados]);

  // 1. Dados para Agrupamento Mensal (Jan a Dez) - considerando entradas validadas por TESOURARIA_RECEB
  const dadosMensais = useMemo(() => {
    const meses = [
      { num: 1, nome: 'Jan' },
      { num: 2, nome: 'Fev' },
      { num: 3, nome: 'Mar' },
      { num: 4, nome: 'Abr' },
      { num: 5, nome: 'Mai' },
      { num: 6, nome: 'Jun' },
      { num: 7, nome: 'Jul' },
      { num: 8, nome: 'Ago' },
      { num: 9, nome: 'Set' },
      { num: 10, nome: 'Out' },
      { num: 11, nome: 'Nov' },
      { num: 12, nome: 'Dez' }
    ];

    return meses.map(m => {
      const lancsDoMes = lancamentosFiltrados.filter(l => l.mes === m.num);
      const entradas = lancsDoMes
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((sum, l) => sum + (l.Total ?? l.valorTotal ?? 0), 0);
      const pix = lancsDoMes
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((sum, l) => sum + (l.ValorOferta ?? l.valorPix ?? 0), 0);
      const especie = lancsDoMes
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((sum, l) => sum + (l.OfertaEspecie ?? l.valorEspecie ?? 0), 0);
      const saidas = lancsDoMes
        .filter(l => l.tipo === 'SAIDA')
        .reduce((sum, l) => sum + l.valorTotal, 0);

      return {
        mesNumero: m.num,
        mesNome: m.nome,
        entradas,
        pix,
        especie,
        saidas,
        saldo: entradas - saidas,
        qtdLancs: lancsDoMes.length
      };
    });
  }, [lancamentosFiltrados]);

  // 2. Dados para Agrupamento por Setor - entradas somente validadas por TESOURARIA_RECEB
  const dadosSetores = useMemo(() => {
    const setorMap = new Map<string, { entradas: number; pix: number; especie: number; saidas: number; qtd: number }>();
    
    SETORES_DISPONIVEIS.forEach(s => {
      setorMap.set(s, { entradas: 0, pix: 0, especie: 0, saidas: 0, qtd: 0 });
    });

    lancamentosFiltrados.forEach(l => {
      const current = setorMap.get(l.setor) || { entradas: 0, pix: 0, especie: 0, saidas: 0, qtd: 0 };
      if (l.tipo === 'ENTRADA') {
        if (l.TESOURARIA_RECEB === true) {
          current.entradas += (l.Total ?? l.valorTotal ?? 0);
          current.pix += (l.ValorOferta ?? l.valorPix ?? 0);
          current.especie += (l.OfertaEspecie ?? l.valorEspecie ?? 0);
        }
      } else {
        current.saidas += l.valorTotal;
      }
      current.qtd += 1;
      setorMap.set(l.setor, current);
    });

    const totalGlobalEntradas = metricas.totalEntradas || 1;

    return Array.from(setorMap.entries())
      .map(([nome, vals]) => ({
        setor: nome,
        ...vals,
        saldo: vals.entradas - vals.saidas,
        percentual: (vals.entradas / totalGlobalEntradas) * 100
      }))
      .sort((a, b) => b.entradas - a.entradas);
  }, [lancamentosFiltrados, metricas.totalEntradas]);

  // 3. Dados para Agrupamento por Área - entradas somente validadas por TESOURARIA_RECEB
  const dadosAreas = useMemo(() => {
    const areaMap = new Map<string, { entradas: number; saidas: number; pix: number; especie: number; qtd: number }>();

    AREAS_DISPONIVEIS.forEach(a => {
      areaMap.set(a, { entradas: 0, saidas: 0, pix: 0, especie: 0, qtd: 0 });
    });

    lancamentosFiltrados.forEach(l => {
      const current = areaMap.get(l.area) || { entradas: 0, saidas: 0, pix: 0, especie: 0, qtd: 0 };
      if (l.tipo === 'ENTRADA') {
        if (l.TESOURARIA_RECEB === true) {
          current.entradas += (l.Total ?? l.valorTotal ?? 0);
          current.pix += (l.ValorOferta ?? l.valorPix ?? 0);
          current.especie += (l.OfertaEspecie ?? l.valorEspecie ?? 0);
        }
      } else {
        current.saidas += l.valorTotal;
      }
      current.qtd += 1;
      areaMap.set(l.area, current);
    });

    return Array.from(areaMap.entries()).map(([area, vals]) => ({
      area,
      ...vals,
      saldo: vals.entradas - vals.saidas
    }));
  }, [lancamentosFiltrados]);

  // 4. Dados para Comparativo Anual (2025 vs 2026) - entradas validadas por TESOURARIA_RECEB
  const dadosAnuais = useMemo(() => {
    const anos = [2025, 2026];
    return anos.map(ano => {
      const lancsAno = listaLancamentos.filter(l => l.ano === ano);
      const entradas = lancsAno
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((s, l) => s + (l.Total ?? l.valorTotal ?? 0), 0);
      const saidas = lancsAno
        .filter(l => l.tipo === 'SAIDA')
        .reduce((s, l) => s + l.valorTotal, 0);
      const pix = lancsAno
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((s, l) => s + (l.ValorOferta ?? l.valorPix ?? 0), 0);
      const especie = lancsAno
        .filter(l => l.tipo === 'ENTRADA' && l.TESOURARIA_RECEB === true)
        .reduce((s, l) => s + (l.OfertaEspecie ?? l.valorEspecie ?? 0), 0);

      return {
        ano,
        entradas,
        saidas,
        saldo: entradas - saidas,
        pix,
        especie,
        totalLancs: lancsAno.length
      };
    });
  }, [listaLancamentos]);

  // 5. Dados para Agrupamento por Categoria
  const dadosCategorias = useMemo(() => {
    const catMap = new Map<string, { tipo: string; total: number; qtd: number }>();
    
    lancamentosFiltrados.forEach(l => {
      const cur = catMap.get(l.categoria) || { tipo: l.tipo, total: 0, qtd: 0 };
      cur.total += l.valorTotal;
      cur.qtd += 1;
      catMap.set(l.categoria, cur);
    });

    return Array.from(catMap.entries())
      .map(([cat, v]) => ({
        categoria: cat,
        tipo: v.tipo,
        total: v.total,
        qtd: v.qtd
      }))
      .sort((a, b) => b.total - a.total);
  }, [lancamentosFiltrados]);

  // Manipulação de exportação
  const handleExportExcel = () => {
    ExportService.exportarParaExcel(
      lancamentosFiltrados,
      filtros,
      `Fluxo_Caixa_SharePoint_${filtros.setor}_${filtros.ano}`
    );
  };

  const handleExportPDF = () => {
    ExportService.exportarParaPDF(
      lancamentosFiltrados,
      filtros,
      `Relatório de Fluxo de Caixa (${filtros.setor === 'todos' ? 'Geral' : filtros.setor})`
    );
  };

  const handleLimparFiltros = () => {
    setFiltros({
      periodoPredefinido: 'todos',
      dataInicio: '',
      dataFim: '',
      ano: anoSelecionado,
      mes: 'todos',
      setor: 'todos',
      area: 'todos',
      celula: 'todos',
      tipo: 'todos',
      categoria: 'todos',
      metodo: 'todos',
      status: 'todos'
    });
    setBuscaTexto('');
    setPaginaAtual(1);
  };

  // Paginação da tabela total
  const totalPaginas = Math.ceil(lancamentosFiltrados.length / itensPorPagina) || 1;
  const lancamentosPaginados = useMemo(() => {
    const start = (paginaAtual - 1) * itensPorPagina;
    return lancamentosFiltrados.slice(start, start + itensPorPagina);
  }, [lancamentosFiltrados, paginaAtual, itensPorPagina]);

  // Altura máxima para gráfico mensal
  const maxValorMensal = Math.max(...dadosMensais.map(d => Math.max(d.entradas, d.saidas)), 5000);

  return (
    <div id="fluxo-caixa-container" className="p-3.5 sm:p-6 space-y-4 sm:space-y-6 max-w-[1600px] mx-auto">
      {/* Barra de Ações Rápidas & Exportações */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-[#23273c] p-3.5 sm:p-4 rounded-xl border border-[#313752]">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-indigo-600/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400 shrink-0">
            <TrendingUp className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-sm sm:text-base font-bold text-white flex items-center gap-2">
              Fluxo de Caixa & Relatório Dinâmico
            </h2>
            <p className="text-[11px] sm:text-xs text-slate-400">
              Leitura em tempo real do banco de dados SharePoint • {listaLancamentos.length} registros
            </p>
          </div>
        </div>

        {/* Action buttons */}
        <div className="flex items-center flex-wrap gap-2 w-full sm:w-auto">
          <button
            id="btn-export-excel-main"
            onClick={handleExportExcel}
            className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold shadow transition-all duration-150 cursor-pointer"
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0" />
            <span>Exportar Excel</span>
          </button>

          <button
            id="btn-export-pdf-main"
            onClick={handleExportPDF}
            className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold shadow transition-all duration-150 cursor-pointer"
          >
            <FileText className="w-4 h-4 shrink-0" />
            <span>Exportar PDF</span>
          </button>

          <button
            id="btn-print-view"
            onClick={() => window.print()}
            className="p-2 rounded-lg bg-[#2c324b] hover:bg-[#373e5e] text-slate-300 hover:text-white border border-[#3e4768] transition-colors cursor-pointer"
            title="Imprimir Relatório"
            aria-label="Imprimir"
          >
            <Printer className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* KPI Cards do Fluxo de Caixa */}
      <div id="fluxo-caixa-kpis" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Entradas Validadas */}
        <div className="bg-[#23273c] p-4 rounded-xl border border-[#313752]">
          <div className="flex items-center justify-between text-xs text-slate-400 mb-1">
            <span className="font-medium text-emerald-400">Entradas Validadas (TESOURARIA_RECEB)</span>
            <span className="p-1 rounded bg-emerald-500/10 text-emerald-400">
              <TrendingUp className="w-4 h-4" />
            </span>
          </div>
          <div className="text-2xl font-black text-white tracking-tight">
            {formatBRL(metricas.totalEntradas)}
          </div>
          <div className="mt-2 flex items-center justify-between text-[11px] text-slate-300 border-t border-[#2e344e] pt-2">
            <span>PIX: <strong>{formatBRL(metricas.totalPix)}</strong></span>
            <span>Espécie: <strong>{formatBRL(metricas.totalEspecie)}</strong></span>
          </div>
        </div>

        {/* Saídas */}
        <div className="bg-[#23273c] p-4 rounded-xl border border-[#313752]">
          <div className="flex items-center justify-between text-xs text-slate-400 mb-1">
            <span className="font-medium">Total de Saídas / Despesas</span>
            <span className="p-1 rounded bg-rose-500/10 text-rose-400">
              <TrendingDown className="w-4 h-4" />
            </span>
          </div>
          <div className="text-2xl font-black text-white tracking-tight">
            {formatBRL(metricas.totalSaidas)}
          </div>
          <div className="mt-2 text-[11px] text-slate-400 border-t border-[#2e344e] pt-2 flex items-center justify-between">
            <span>Custos operacionais e ministeriais</span>
            <span className="text-rose-400 font-semibold">{metricas.totalEntradas > 0 ? ((metricas.totalSaidas / metricas.totalEntradas) * 100).toFixed(0) : 0}% da receita</span>
          </div>
        </div>

        {/* Saldo Líquido */}
        <div className="bg-[#23273c] p-4 rounded-xl border border-[#313752]">
          <div className="flex items-center justify-between text-xs text-slate-400 mb-1">
            <span className="font-medium">Saldo Líquido Operacional</span>
            <span className="p-1 rounded bg-indigo-500/10 text-indigo-400">
              <DollarSign className="w-4 h-4" />
            </span>
          </div>
          <div className={`text-2xl font-black tracking-tight ${metricas.saldoOperacional >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
            {formatBRL(metricas.saldoOperacional)}
          </div>
          <div className="mt-2 text-[11px] text-slate-300 border-t border-[#2e344e] pt-2 flex items-center justify-between">
            <span>Situação Financeira:</span>
            <span className={`font-bold px-1.5 py-0.5 rounded text-[10px] ${
              metricas.saldoOperacional >= 0 ? 'bg-emerald-900/40 text-emerald-300' : 'bg-rose-900/40 text-rose-300'
            }`}>
              {metricas.saldoOperacional >= 0 ? 'SUPERÁVIT' : 'DÉFICIT'}
            </span>
          </div>
        </div>

        {/* Quantidade e Status de Validação */}
        <div className="bg-[#23273c] p-4 rounded-xl border border-[#313752]">
          <div className="flex items-center justify-between text-xs text-slate-400 mb-1">
            <span className="font-medium">Auditoria de Relatórios</span>
            <span className="p-1 rounded bg-amber-500/10 text-amber-400">
              <BarChart3 className="w-4 h-4" />
            </span>
          </div>
          <div className="text-2xl font-black text-white tracking-tight">
            {metricas.qtdValidados} <span className="text-xs font-normal text-emerald-400 font-bold">Validados</span>
            <span className="text-xs font-normal text-slate-400"> / {metricas.qtdLancamentos}</span>
          </div>
          <div className="mt-2 text-[11px] text-slate-400 border-t border-[#2e344e] pt-2 flex items-center justify-between">
            <span>Pendente Validação:</span>
            <span className="text-amber-400 font-medium">{formatBRL(metricas.totalPendenteValidacao)} ({metricas.qtdPendentes} rel.)</span>
          </div>
        </div>
      </div>

      {/* Painel Avançado de Filtros Personalizados */}
      <div id="painel-filtros-fluxo" className="bg-[#23273c] p-4 rounded-xl border border-[#313752] space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#2e344e] pb-3">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-indigo-400" />
            <span className="text-xs font-bold text-white uppercase tracking-wider">
              Filtros Personalizados & Parâmetros de Relatório
            </span>
          </div>
          <button
            id="btn-limpar-filtros"
            onClick={handleLimparFiltros}
            className="flex items-center gap-1 text-xs text-slate-400 hover:text-white transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3 h-3" />
            Limpar Filtros
          </button>
        </div>

        {/* Linha de Filtros */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
          {/* Filtro Setor */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Setor</label>
            <select
              id="filtro-setor"
              value={filtros.setor}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, setor: e.target.value as SetorTipo | 'todos' }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todos os Setores</option>
              {SETORES_DISPONIVEIS.map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          {/* Filtro Área */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Área / Região</label>
            <select
              id="filtro-area"
              value={filtros.area}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, area: e.target.value as AreaTipo | 'todos' }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todas as Áreas</option>
              {AREAS_DISPONIVEIS.map(a => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>

          {/* Filtro Mês */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Mês Competência</label>
            <select
              id="filtro-mes"
              value={filtros.mes}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, mes: e.target.value === 'todos' ? 'todos' : Number(e.target.value) }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todos os Meses</option>
              <option value="1">Janeiro</option>
              <option value="2">Fevereiro</option>
              <option value="3">Março</option>
              <option value="4">Abril</option>
              <option value="5">Maio</option>
              <option value="6">Junho</option>
              <option value="7">Julho</option>
              <option value="8">Agosto</option>
              <option value="9">Setembro</option>
              <option value="10">Outubro</option>
              <option value="11">Novembro</option>
              <option value="12">Dezembro</option>
            </select>
          </div>

          {/* Filtro Tipo */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Tipo de Fluxo</label>
            <select
              id="filtro-tipo"
              value={filtros.tipo}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, tipo: e.target.value as any }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Entradas e Saídas</option>
              <option value="ENTRADA">Apenas Entradas (+)</option>
              <option value="SAIDA">Apenas Saídas (-)</option>
            </select>
          </div>

          {/* Filtro Categoria */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Categoria</label>
            <select
              id="filtro-categoria"
              value={filtros.categoria}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, categoria: e.target.value }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todas as Categorias</option>
              <optgroup label="Entradas">
                {CATEGORIAS_ENTRADA.map(c => <option key={c} value={c}>{c}</option>)}
              </optgroup>
              <optgroup label="Saídas">
                {CATEGORIAS_SAIDA.map(c => <option key={c} value={c}>{c}</option>)}
              </optgroup>
            </select>
          </div>

          {/* Filtro Forma Pagamento */}
          <div>
            <label className="block text-slate-400 font-medium mb-1">Forma Pagamento</label>
            <select
              id="filtro-metodo"
              value={filtros.metodo}
              onChange={(e) => {
                setFiltros(prev => ({ ...prev, metodo: e.target.value as any }));
                setPaginaAtual(1);
              }}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400 cursor-pointer"
            >
              <option value="todos">Todas</option>
              <option value="PIX">PIX</option>
              <option value="ESPECIE">Espécie / Dinheiro</option>
              <option value="TRANSFERENCIA">Transferência</option>
              <option value="BOLETO">Boleto</option>
            </select>
          </div>
        </div>

        {/* Filtro por Intervalo de Datas e Busca de Célula */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-[#2e344e] text-xs">
          <div>
            <label className="block text-slate-400 font-medium mb-1">Data Início</label>
            <input
              type="date"
              id="filtro-data-inicio"
              value={filtros.dataInicio}
              onChange={(e) => setFiltros(prev => ({ ...prev, dataInicio: e.target.value }))}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400"
            />
          </div>
          <div>
            <label className="block text-slate-400 font-medium mb-1">Data Fim</label>
            <input
              type="date"
              id="filtro-data-fim"
              value={filtros.dataFim}
              onChange={(e) => setFiltros(prev => ({ ...prev, dataFim: e.target.value }))}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-2.5 py-1.5 text-white focus:outline-none focus:border-indigo-400"
            />
          </div>
          <div>
            <label className="block text-slate-400 font-medium mb-1">Pesquisa Rápida (Célula / Descrição)</label>
            <div className="relative">
              <input
                type="text"
                id="busca-texto-fluxo"
                value={buscaTexto}
                onChange={(e) => setBuscaTexto(e.target.value)}
                placeholder="Ex: Leão de Judá, Refúgio, Manutenção..."
                className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg pl-8 pr-3 py-1.5 text-white focus:outline-none focus:border-indigo-400"
              />
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2" />
            </div>
          </div>
        </div>
      </div>

      {/* O MAIS IMPORTANTE: Barra de Seleção de Visualizações Dinâmicas */}
      <div id="visualizacoes-tabs" className="flex items-center gap-2 border-b border-[#2e344e] pb-3 overflow-x-auto scrollbar-thin">
        <span className="text-xs font-bold text-slate-400 uppercase tracking-wider mr-2 flex items-center gap-1.5 shrink-0">
          <Layers className="w-3.5 h-3.5 text-indigo-400" />
          Visualização:
        </span>
        
        <button
          id="tab-vis-mensal"
          onClick={() => setAgrupamentoAtivo('mensal')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'mensal'
              ? 'bg-indigo-600 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Visão Mensal (Jan-Dez)
        </button>

        <button
          id="tab-vis-setor"
          onClick={() => setAgrupamentoAtivo('setor')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'setor'
              ? 'bg-indigo-600 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Por Setor (Laranja, Amarelo, Roxo, Verde)
        </button>

        <button
          id="tab-vis-area"
          onClick={() => setAgrupamentoAtivo('area')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'area'
              ? 'bg-indigo-600 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Por Área (Supervisões)
        </button>

        <button
          id="tab-vis-anual"
          onClick={() => setAgrupamentoAtivo('anual')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'anual'
              ? 'bg-indigo-600 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Visão Anual (2025 vs 2026)
        </button>

        <button
          id="tab-vis-categoria"
          onClick={() => setAgrupamentoAtivo('categoria')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'categoria'
              ? 'bg-indigo-600 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Por Categorias
        </button>

        <button
          id="tab-vis-tabela-total"
          onClick={() => setAgrupamentoAtivo('tabela-total')}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer shrink-0 ${
            agrupamentoAtivo === 'tabela-total'
              ? 'bg-emerald-700 text-white shadow-md'
              : 'bg-[#252a3f] text-slate-300 hover:bg-[#2e3550] hover:text-white'
          }`}
        >
          Leitura Total SharePoint (Base Completa)
        </button>
      </div>

      {/* Conteúdo Dinâmico Baseado na Aba Ativa */}

      {/* 1. VISÃO MENSAL (GRÁFICO DINÂMICO + TABELA COMPARATIVA) */}
      {agrupamentoAtivo === 'mensal' && (
        <div className="space-y-6 animate-in fade-in duration-200">
          {/* Gráfico Dinâmico de Barras Mensais */}
          <div className="bg-[#23273c] p-5 rounded-xl border border-[#313752]">
            <div className="flex flex-wrap items-center justify-between gap-2 mb-6">
              <div>
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <BarChart3 className="w-4 h-4 text-indigo-400" />
                  Evolução Mensal do Fluxo de Caixa (Entradas vs Saídas)
                </h3>
                <p className="text-xs text-slate-400">
                  Comparativo de arrecadação de células, dízimos e despesas mês a mês
                </p>
              </div>

              {/* Legenda do gráfico */}
              <div className="flex items-center gap-4 text-xs">
                <div className="flex items-center gap-1.5">
                  <div className="w-3 h-3 rounded-xs bg-emerald-500"></div>
                  <span className="text-slate-300">Entradas (R$)</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <div className="w-3 h-3 rounded-xs bg-rose-500"></div>
                  <span className="text-slate-300">Saídas (R$)</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <div className="w-3 h-3 rounded-xs bg-sky-400"></div>
                  <span className="text-slate-300">Saldo Líquido</span>
                </div>
              </div>
            </div>

            {/* Renderização do Gráfico em Barras */}
            <div className="overflow-x-auto pb-2 scrollbar-thin">
              <div className="min-w-[550px] md:min-w-0 grid grid-cols-12 gap-2 h-64 items-end pt-4 pb-2 border-b border-[#2e344e]">
                {dadosMensais.map(m => {
                  const altEntrada = Math.min(100, Math.max(8, (m.entradas / maxValorMensal) * 100));
                  const altSaida = Math.min(100, Math.max(8, (m.saidas / maxValorMensal) * 100));
                  
                  return (
                    <div key={m.mesNumero} className="flex flex-col items-center h-full justify-end group relative">
                      {/* Tooltip on hover */}
                      <div className="absolute -top-16 bg-[#161826] border border-[#3b4366] text-white p-2 rounded shadow-xl opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-20 text-[10px] whitespace-nowrap min-w-[130px]">
                        <p className="font-bold text-slate-200 border-b border-slate-700 pb-0.5">{m.mesNome} / {filtros.ano}</p>
                        <p className="text-emerald-400">Entradas: {formatBRL(m.entradas)}</p>
                        <p className="text-rose-400">Saídas: {formatBRL(m.saidas)}</p>
                        <p className="text-sky-300 font-semibold">Saldo: {formatBRL(m.saldo)}</p>
                      </div>

                      {/* Barras agrupadas */}
                      <div className="flex items-end gap-1 w-full justify-center h-full pb-1">
                        {/* Barra Entrada */}
                        <div
                          style={{ height: `${altEntrada}%` }}
                          className="w-2.5 sm:w-3.5 bg-emerald-500 hover:bg-emerald-400 rounded-t-sm transition-all duration-300"
                        />
                        {/* Barra Saída */}
                        <div
                          style={{ height: `${altSaida}%` }}
                          className="w-2.5 sm:w-3.5 bg-rose-500 hover:bg-rose-400 rounded-t-sm transition-all duration-300"
                        />
                      </div>
                      {/* Label mês */}
                      <span className="text-[11px] font-semibold text-slate-400 mt-1">
                        {m.mesNome}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Tabela Comparativa Mensal */}
          <div className="bg-[#23273c] rounded-xl border border-[#313752] overflow-hidden">
            <div className="p-4 border-b border-[#313752] flex justify-between items-center">
              <h4 className="text-xs font-bold uppercase text-white tracking-wider">
                Demonstrativo Financeiro Mês a Mês
              </h4>
              <span className="text-xs text-slate-400">Ano de Competência: {filtros.ano}</span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-[#181b2a] text-slate-300 uppercase tracking-wider text-[11px] border-b border-[#2e344e]">
                  <tr>
                    <th className="px-4 py-3">Mês</th>
                    <th className="px-4 py-3 text-right">PIX (R$)</th>
                    <th className="px-4 py-3 text-right">Espécie (R$)</th>
                    <th className="px-4 py-3 text-right">Total Entradas</th>
                    <th className="px-4 py-3 text-right">Total Saídas</th>
                    <th className="px-4 py-3 text-right">Resultado Líquido</th>
                    <th className="px-4 py-3 text-center">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#2a3048]">
                  {dadosMensais.map(m => (
                    <tr key={m.mesNumero} className="hover:bg-[#282d46] transition-colors">
                      <td className="px-4 py-3 font-semibold text-white">
                        {m.mesNome} / {filtros.ano}
                      </td>
                      <td className="px-4 py-3 text-right text-slate-300">
                        {formatBRL(m.pix)}
                      </td>
                      <td className="px-4 py-3 text-right text-slate-300">
                        {formatBRL(m.especie)}
                      </td>
                      <td className="px-4 py-3 text-right font-bold text-emerald-400">
                        {formatBRL(m.entradas)}
                      </td>
                      <td className="px-4 py-3 text-right font-bold text-rose-400">
                        {formatBRL(m.saidas)}
                      </td>
                      <td className={`px-4 py-3 text-right font-black ${m.saldo >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                        {formatBRL(m.saldo)}
                      </td>
                      <td className="px-4 py-3 text-center">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          m.saldo >= 0 ? 'bg-emerald-950/80 text-emerald-300 border border-emerald-500/30' : 'bg-rose-950/80 text-rose-300 border border-rose-500/30'
                        }`}>
                          {m.saldo >= 0 ? 'SUPERÁVIT' : 'DÉFICIT'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* 2. VISÃO POR SETOR (DETALHAMENTO DOS SETORES) */}
      {agrupamentoAtivo === 'setor' && (
        <div className="space-y-6 animate-in fade-in duration-200">
          <div className="bg-[#23273c] rounded-xl border border-[#313752] overflow-hidden">
            <div className="p-4 border-b border-[#313752] flex flex-wrap items-center justify-between gap-2">
              <div>
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <MapPin className="w-4 h-4 text-indigo-400" />
                  Consolidado Financeiro por Setor
                </h3>
                <p className="text-xs text-slate-400">
                  Ranking de contribuição financeira, distribuição PIX vs Espécie e saldo por setor
                </p>
              </div>
              <span className="text-xs font-semibold px-2.5 py-1 rounded bg-[#2e3550] text-slate-200">
                Total de {dadosSetores.length} Setores Mapeados
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-[#181b2a] text-slate-300 uppercase tracking-wider text-[11px] border-b border-[#2e344e]">
                  <tr>
                    <th className="px-4 py-3">Setor</th>
                    <th className="px-4 py-3 text-center">% do Total</th>
                    <th className="px-4 py-3 text-right">PIX (R$)</th>
                    <th className="px-4 py-3 text-right">Espécie (R$)</th>
                    <th className="px-4 py-3 text-right">Total Arrecadado</th>
                    <th className="px-4 py-3 text-right">Despesas Alocadas</th>
                    <th className="px-4 py-3 text-right">Saldo Líquido</th>
                    <th className="px-4 py-3 text-center">Registros</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#2a3048]">
                  {dadosSetores.map(s => (
                    <tr key={s.setor} className="hover:bg-[#282d46] transition-colors">
                      <td className="px-4 py-3 font-bold text-white flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-indigo-400"></span>
                        {s.setor}
                      </td>
                      <td className="px-4 py-3 text-center">
                        <div className="flex items-center justify-center gap-2">
                          <div className="w-16 bg-[#181b2a] rounded-full h-2 overflow-hidden">
                            <div
                              className="bg-indigo-500 h-full rounded-full"
                              style={{ width: `${Math.min(100, s.percentual)}%` }}
                            />
                          </div>
                          <span className="font-semibold text-slate-300 text-[11px]">
                            {s.percentual.toFixed(1)}%
                          </span>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-right text-slate-300">{formatBRL(s.pix)}</td>
                      <td className="px-4 py-3 text-right text-slate-300">{formatBRL(s.especie)}</td>
                      <td className="px-4 py-3 text-right font-bold text-emerald-400">{formatBRL(s.entradas)}</td>
                      <td className="px-4 py-3 text-right font-bold text-rose-400">{formatBRL(s.saidas)}</td>
                      <td className={`px-4 py-3 text-right font-black ${s.saldo >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                        {formatBRL(s.saldo)}
                      </td>
                      <td className="px-4 py-3 text-center font-semibold text-slate-400">{s.qtd}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* 3. VISÃO POR ÁREA (SUPERVISÕES) */}
      {agrupamentoAtivo === 'area' && (
        <div className="space-y-6 animate-in fade-in duration-200">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {dadosAreas.map(a => (
              <div key={a.area} className="bg-[#23273c] p-4 rounded-xl border border-[#313752]">
                <div className="flex items-center justify-between border-b border-[#2e344e] pb-2.5 mb-3">
                  <h4 className="font-bold text-white text-sm flex items-center gap-2">
                    <MapPin className="w-4 h-4 text-emerald-400" />
                    {a.area}
                  </h4>
                  <span className="text-[11px] font-semibold text-slate-400 bg-[#191b29] px-2 py-0.5 rounded">
                    {a.qtd} lançamentos
                  </span>
                </div>

                <div className="space-y-2 text-xs">
                  <div className="flex justify-between text-slate-300">
                    <span>Entradas Totais:</span>
                    <strong className="text-emerald-400">{formatBRL(a.entradas)}</strong>
                  </div>
                  <div className="flex justify-between text-slate-400 text-[11px]">
                    <span>• PIX:</span>
                    <span>{formatBRL(a.pix)}</span>
                  </div>
                  <div className="flex justify-between text-slate-400 text-[11px]">
                    <span>• Espécie / Dinheiro:</span>
                    <span>{formatBRL(a.especie)}</span>
                  </div>
                  <div className="flex justify-between text-slate-300 pt-1 border-t border-[#2e344e]">
                    <span>Saídas / Despesas:</span>
                    <strong className="text-rose-400">{formatBRL(a.saidas)}</strong>
                  </div>
                  <div className="flex justify-between text-slate-200 pt-1 border-t border-[#2e344e] font-bold">
                    <span>Saldo Operacional:</span>
                    <span className={a.saldo >= 0 ? 'text-emerald-300' : 'text-rose-300'}>
                      {formatBRL(a.saldo)}
                    </span>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 4. VISÃO ANUAL (2025 vs 2026) */}
      {agrupamentoAtivo === 'anual' && (
        <div className="space-y-6 animate-in fade-in duration-200">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {dadosAnuais.map(ano => (
              <div key={ano.ano} className="bg-[#23273c] p-6 rounded-xl border border-[#313752] space-y-4">
                <div className="flex items-center justify-between border-b border-[#2e344e] pb-3">
                  <h3 className="text-lg font-black text-white flex items-center gap-2">
                    <Calendar className="w-5 h-5 text-indigo-400" />
                    Exercício Financeiro {ano.ano}
                  </h3>
                  <span className="text-xs font-semibold px-2.5 py-1 rounded bg-[#1c1f30] text-slate-300">
                    {ano.totalLancs} registros
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-3 text-xs">
                  <div className="bg-[#181b2a] p-3 rounded-lg">
                    <span className="text-slate-400 block text-[11px]">Receita Anual Bruta</span>
                    <strong className="text-base text-emerald-400">{formatBRL(ano.entradas)}</strong>
                  </div>
                  <div className="bg-[#181b2a] p-3 rounded-lg">
                    <span className="text-slate-400 block text-[11px]">Despesas Anuais</span>
                    <strong className="text-base text-rose-400">{formatBRL(ano.saidas)}</strong>
                  </div>
                </div>

                <div className="bg-[#181b2a] p-3.5 rounded-lg flex items-center justify-between">
                  <span className="text-xs font-medium text-slate-300">Superávit / Resultado do Ano</span>
                  <span className={`text-base font-black ${ano.saldo >= 0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                    {formatBRL(ano.saldo)}
                  </span>
                </div>

                <div className="text-xs text-slate-400 space-y-1 pt-2">
                  <div className="flex justify-between">
                    <span>Arrecadação Digital (PIX):</span>
                    <strong className="text-slate-200">{formatBRL(ano.pix)}</strong>
                  </div>
                  <div className="flex justify-between">
                    <span>Arrecadação Física (Espécie):</span>
                    <strong className="text-slate-200">{formatBRL(ano.especie)}</strong>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 5. VISÃO POR CATEGORIAS */}
      {agrupamentoAtivo === 'categoria' && (
        <div className="bg-[#23273c] rounded-xl border border-[#313752] overflow-hidden animate-in fade-in duration-200">
          <div className="p-4 border-b border-[#313752] flex justify-between items-center">
            <h3 className="text-sm font-bold text-white flex items-center gap-2">
              <Layers className="w-4 h-4 text-indigo-400" />
              Detalhamento de Fluxo de Caixa por Categorias
            </h3>
            <span className="text-xs text-slate-400">Total de {dadosCategorias.length} categorias</span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left">
              <thead className="bg-[#181b2a] text-slate-300 uppercase tracking-wider text-[11px] border-b border-[#2e344e]">
                <tr>
                  <th className="px-4 py-3">Categoria</th>
                  <th className="px-4 py-3">Tipo</th>
                  <th className="px-4 py-3 text-right">Qtd Registros</th>
                  <th className="px-4 py-3 text-right">Volume Total (R$)</th>
                  <th className="px-4 py-3 text-right">% do Fluxo</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#2a3048]">
                {dadosCategorias.map(c => {
                  const baseTotal = c.tipo === 'ENTRADA' ? metricas.totalEntradas : metricas.totalSaidas;
                  const perc = baseTotal > 0 ? (c.total / baseTotal) * 100 : 0;
                  return (
                    <tr key={c.categoria} className="hover:bg-[#282d46] transition-colors">
                      <td className="px-4 py-3 font-semibold text-white">{c.categoria}</td>
                      <td className="px-4 py-3">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          c.tipo === 'ENTRADA' ? 'bg-emerald-950 text-emerald-300' : 'bg-rose-950 text-rose-300'
                        }`}>
                          {c.tipo === 'ENTRADA' ? '+ Entrada' : '- Saída'}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-right text-slate-300">{c.qtd}</td>
                      <td className={`px-4 py-3 text-right font-black ${
                        c.tipo === 'ENTRADA' ? 'text-emerald-400' : 'text-rose-400'
                      }`}>
                        {formatBRL(c.total)}
                      </td>
                      <td className="px-4 py-3 text-right font-medium text-slate-400">
                        {perc.toFixed(1)}%
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 6. LEITURA TOTAL SHAREPOINT (TABELA COMPLETA COM PAGINAÇÃO) */}
      {agrupamentoAtivo === 'tabela-total' && (
        <div className="bg-[#23273c] rounded-xl border border-[#313752] overflow-hidden animate-in fade-in duration-200">
          <div className="p-4 border-b border-[#313752] flex flex-wrap items-center justify-between gap-3 bg-[#1e2235]">
            <div className="flex items-center gap-2">
              <Database className="w-4 h-4 text-emerald-400" />
              <div>
                <h3 className="text-sm font-bold text-white">
                  Base de Dados Completa do SharePoint (Auditoria Total)
                </h3>
                <p className="text-[11px] text-slate-400">
                  Exibindo {lancamentosFiltrados.length} lançamentos de acordo com os filtros aplicados
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleExportExcel}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-emerald-700/80 hover:bg-emerald-600 text-white text-xs font-semibold"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                Baixar Planilha Completa
              </button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left">
              <thead className="bg-[#181b2a] text-slate-300 uppercase tracking-wider text-[11px] border-b border-[#2e344e]">
                <tr>
                  <th className="px-3.5 py-3">DataCelula</th>
                  <th className="px-3 py-3 text-center">Semana</th>
                  <th className="px-3.5 py-3">Célula</th>
                  <th className="px-3 py-3">Líder</th>
                  <th className="px-3 py-3">Setor</th>
                  <th className="px-3 py-3">Área</th>
                  <th className="px-3 py-3 text-right">ValorOferta (PIX)</th>
                  <th className="px-3 py-3 text-right">OfertaEspecie</th>
                  <th className="px-3.5 py-3 text-right">Total</th>
                  <th className="px-3 py-3 text-center">TESOURARIA_RECEB</th>
                  <th className="px-3 py-3 text-center">Data Tesouraria</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#2a3048]">
                {lancamentosPaginados.map(l => {
                  const pix = l.ValorOferta ?? l.valorPix ?? 0;
                  const esp = l.OfertaEspecie ?? l.valorEspecie ?? 0;
                  const total = l.Total ?? l.valorTotal ?? (pix + esp);
                  const isRecebido = l.TESOURARIA_RECEB === true;

                  return (
                    <tr key={l.id} className="hover:bg-[#282d46] transition-colors">
                      <td className="px-3.5 py-2.5 font-medium text-slate-300">{l.DataCelula || l.data}</td>
                      <td className="px-3 py-2.5 text-slate-400 text-center font-mono">Sem. {l.NumSemana ?? l.semanaNumero}</td>
                      <td className="px-3.5 py-2.5 font-bold text-white">{l.Célula || l.celulaNome}</td>
                      <td className="px-3 py-2.5 text-slate-300">{l.LíderCelula || '-'}</td>
                      <td className="px-3 py-2.5 font-semibold text-indigo-300">{l.Setor || l.setor}</td>
                      <td className="px-3 py-2.5 text-slate-300">{l.Area || l.area}</td>
                      <td className="px-3 py-2.5 text-right text-slate-300 font-mono">
                        {pix > 0 ? formatBRL(pix) : '-'}
                      </td>
                      <td className="px-3 py-2.5 text-right text-slate-300 font-mono">
                        {esp > 0 ? formatBRL(esp) : '-'}
                      </td>
                      <td className={`px-3.5 py-2.5 text-right font-black font-mono ${
                        isRecebido ? 'text-emerald-400' : 'text-amber-400'
                      }`}>
                        {formatBRL(total)}
                      </td>
                      <td className="px-3 py-2.5 text-center">
                        <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold ${
                          isRecebido 
                            ? 'bg-emerald-950/80 text-emerald-300 border border-emerald-500/30' 
                            : 'bg-amber-950/80 text-amber-300 border border-amber-500/30'
                        }`}>
                          <span className={`w-1.5 h-1.5 rounded-full ${isRecebido ? 'bg-emerald-400' : 'bg-amber-400'}`} />
                          {isRecebido ? 'TRUE (Validado)' : 'FALSE (Pendente)'}
                        </span>
                      </td>
                      <td className="px-3 py-2.5 text-center text-[11px] text-slate-400 font-mono">
                        {l.DATA_TESOURARIA || (isRecebido ? l.data : '-')}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Paginação */}
          <div className="p-4 border-t border-[#2e344e] flex items-center justify-between text-xs text-slate-400 bg-[#1a1d2d]">
            <span>
              Página <strong>{paginaAtual}</strong> de <strong>{totalPaginas}</strong> ({lancamentosFiltrados.length} itens)
            </span>
            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setPaginaAtual(p => Math.max(1, p - 1))}
                disabled={paginaAtual === 1}
                className="px-3 py-1 rounded bg-[#252a3f] hover:bg-[#303752] text-white disabled:opacity-40 cursor-pointer"
              >
                Anterior
              </button>
              <button
                onClick={() => setPaginaAtual(p => Math.min(totalPaginas, p + 1))}
                disabled={paginaAtual === totalPaginas}
                className="px-3 py-1 rounded bg-[#252a3f] hover:bg-[#303752] text-white disabled:opacity-40 cursor-pointer"
              >
                Próxima
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
