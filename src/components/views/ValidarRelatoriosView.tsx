import React, { useState, useMemo, useEffect } from 'react';
import { 
  RotateCw, 
  CheckCircle2,
  Lock,
  Edit,
  Check,
  Undo2,
  X,
  FileCheck,
  Search,
  Save,
  CheckSquare,
  Square
} from 'lucide-react';
import { LancamentoTesouraria, MembroItem } from '../../types';
import { SharePointService } from '../../services/sharepointService';

interface ValidarRelatoriosViewProps {
  lancamentos: LancamentoTesouraria[];
  anoSelecionado: number;
  usuarioLogado?: MembroItem | null;
  onAtualizarLancamento?: (id: string, dados: Partial<LancamentoTesouraria>) => void;
  onConfirmarLancamento?: (id: string, idTesoureiro?: string | number, dataTesouraria?: string) => void;
  onDesconfirmarLancamento?: (id: string) => void;
  onRefresh: () => void;
}

export const ValidarRelatoriosView: React.FC<ValidarRelatoriosViewProps> = ({
  lancamentos,
  anoSelecionado,
  usuarioLogado,
  onAtualizarLancamento,
  onConfirmarLancamento,
  onDesconfirmarLancamento,
  onRefresh
}) => {
  const [tabAtiva, setTabAtiva] = useState<'P_VALIDAR' | 'CONFIRMADOS'>('P_VALIDAR');
  const [setorFiltro, setSetorFiltro] = useState<string>('Fire');
  const [buscaTexto, setBuscaTexto] = useState<string>('');
  const [setoresDisponiveis, setSetoresDisponiveis] = useState<string[]>([
    'Amarelo', 'Azul', 'Black', 'Diamante', 'Fire', 'Legacy', 'Onix', 'Safira', 'Titanium', 'White'
  ]);
  
  // Itens selecionados via checkbox para ações em lote
  const [selecionados, setSelecionados] = useState<Set<string>>(new Set());

  // Modal de edição de relatório
  const [modalEditarItem, setModalEditarItem] = useState<LancamentoTesouraria | null>(null);
  const [editCelula, setEditCelula] = useState<string>('');
  const [editData, setEditData] = useState<string>('');
  const [editPix, setEditPix] = useState<number>(0);
  const [editEspecie, setEditEspecie] = useState<number>(0);
  const [salvandoEdicao, setSalvandoEdicao] = useState<boolean>(false);
  const [mensagemSucesso, setMensagemSucesso] = useState<string | null>(null);

  // Carrega setores distintos obtidos da base de dados de células
  useEffect(() => {
    const spService = SharePointService.getInstance();
    const setoresIniciais = spService.getSetores();
    if (setoresIniciais && setoresIniciais.length > 0) {
      setSetoresDisponiveis(setoresIniciais);
    }

    // Busca setores oficiais via API
    fetch('/api/sharepoint/setores')
      .then(res => res.json())
      .then(data => {
        if (data && Array.isArray(data.setores) && data.setores.length > 0) {
          setSetoresDisponiveis(data.setores);
        }
      })
      .catch(() => {
        // Mantém setores padrão
      });
  }, []);

  const exibirAlerta = (msg: string) => {
    setMensagemSucesso(msg);
    setTimeout(() => {
      setMensagemSucesso(null);
    }, 4000);
  };

  // Extrai o ano de um relatório de forma segura, suportando DD/MM/YYYY e YYYY-MM-DD
  const extrairAnoLancamento = (item: LancamentoTesouraria): number => {
    if (item.ano && typeof item.ano === 'number' && item.ano > 2000) {
      return item.ano;
    }
    // Verifica formato brasileiro DD/MM/YYYY
    const dataBR = item.dataBR || (item as any).DataNascimento;
    if (dataBR && typeof dataBR === 'string' && dataBR.includes('/')) {
      const parts = dataBR.trim().split('/');
      if (parts.length === 3) {
        const y = parseInt(parts[2], 10);
        if (!isNaN(y) && y > 2000) return y;
      }
    }
    // Verifica data no formato ISO YYYY-MM-DD ou DD/MM/YYYY
    const dataStr = item.data || item.DataCelula || (item as any).dataCelula || '';
    if (dataStr && typeof dataStr === 'string') {
      if (dataStr.includes('/')) {
        const parts = dataStr.trim().split('/');
        if (parts.length === 3) {
          const y = parseInt(parts[2], 10);
          if (!isNaN(y) && y > 2000) return y;
        }
      }
      if (dataStr.length >= 4) {
        const parsed = parseInt(dataStr.slice(0, 4), 10);
        if (!isNaN(parsed) && parsed > 2000) return parsed;
      }
    }
    // Verifica data de criação
    const criadoStr = (item as any).Criado || (item as any).Created || item.dataSincronizacao || '';
    if (criadoStr && typeof criadoStr === 'string' && criadoStr.length >= 4) {
      const parsed = parseInt(criadoStr.slice(0, 4), 10);
      if (!isNaN(parsed) && parsed > 2000) return parsed;
    }
    return anoSelecionado;
  };

  // 1. Filtragem por Ano Selecionado
  const lancamentosDoAno = useMemo(() => {
    return lancamentos.filter(item => {
      const anoItem = extrairAnoLancamento(item);
      return anoItem === anoSelecionado;
    });
  }, [lancamentos, anoSelecionado]);

  // Lista unificada de todos os setores distintos
  const listaSetores = useMemo(() => {
    const conjunto = new Set<string>();
    setoresDisponiveis.forEach(s => s && conjunto.add(s.trim()));
    lancamentos.forEach(l => {
      const s = (l.Setor || l.setor || '').trim();
      if (s) conjunto.add(s);
    });
    return Array.from(conjunto).sort();
  }, [setoresDisponiveis, lancamentos]);

  // 2. Contagem de relatórios NÃO CONFIRMADOS por setor no Ano selecionado
  // Regra: TESOURARIA_RECEB vazia ou false
  const naoConfirmadosPorSetor = useMemo(() => {
    const mapa: Record<string, number> = {};
    listaSetores.forEach(setor => {
      const pendentesSetor = lancamentosDoAno.filter(l => {
        const itemSetor = (l.Setor || l.setor || '').trim();
        const naoConfirmado = l.TESOURARIA_RECEB !== true;
        return itemSetor.toLowerCase() === setor.toLowerCase() && naoConfirmado;
      });
      mapa[setor] = pendentesSetor.length;
    });
    return mapa;
  }, [listaSetores, lancamentosDoAno]);

  // Contagens gerais para as abas no Ano selecionado
  const totalNaoConfirmadosGeral = useMemo(() => {
    return lancamentosDoAno.filter(l => l.TESOURARIA_RECEB !== true).length;
  }, [lancamentosDoAno]);

  const totalConfirmadosGeral = useMemo(() => {
    return lancamentosDoAno.filter(l => l.TESOURARIA_RECEB === true).length;
  }, [lancamentosDoAno]);

  // 3. Filtragem final dos relatórios para exibição na tabela
  const relatoriosFiltrados = useMemo(() => {
    return lancamentosDoAno.filter(item => {
      const isConfirmado = item.TESOURARIA_RECEB === true;

      // Filtro da aba: P/ Validar vs Confirmados
      if (tabAtiva === 'P_VALIDAR' && isConfirmado) return false;
      if (tabAtiva === 'CONFIRMADOS' && !isConfirmado) return false;

      // Filtro de Setor selecionado
      if (setorFiltro !== 'todos') {
        const itemSetor = (item.Setor || item.setor || '').trim();
        if (itemSetor.toLowerCase() !== setorFiltro.toLowerCase()) {
          return false;
        }
      }

      // Filtro de busca textual (célula, líder, semana)
      if (buscaTexto.trim()) {
        const busca = buscaTexto.toLowerCase().trim();
        const celula = (item.Célula || item.celulaNome || '').toLowerCase();
        const lider = (item.LíderCelula || item.liderCelula || (item as any).Nome || '').toLowerCase();
        const semana = String(item.NumSemana ?? item.semanaNumero ?? '');
        if (!celula.includes(busca) && !lider.includes(busca) && !semana.includes(busca)) {
          return false;
        }
      }

      return true;
    });
  }, [lancamentosDoAno, tabAtiva, setorFiltro, buscaTexto]);

  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(val || 0);
  };

  const getFormatDateBR = (item: LancamentoTesouraria): string => {
    if (item.dataBR) return item.dataBR;
    const nasc = (item as any).DataNascimento;
    if (nasc && typeof nasc === 'string' && nasc.includes('/')) return nasc;
    const iso = item.DataCelula || (item as any).dataCelula || item.data || '';
    if (!iso) return '-';
    if (iso.includes('/')) return iso;
    const clean = iso.includes('T') ? iso.split('T')[0] : iso;
    const parts = clean.split('-');
    if (parts.length === 3) {
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    return iso;
  };

  // Abrir Modal de Edição
  const handleAbrirEditar = (item: LancamentoTesouraria) => {
    setModalEditarItem(item);
    setEditCelula(item.Célula || item.celulaNome || '');
    setEditData(getFormatDateBR(item));
    const pix = item.valorPix ?? (item.ValorOferta ?? 0);
    const esp = item.valorEspecie ?? (item.OfertaEspecie ?? 0);
    setEditPix(pix);
    setEditEspecie(esp);
  };

  // Salvar Edição
  const handleSalvarEdicao = async () => {
    if (!modalEditarItem) return;
    setSalvandoEdicao(true);
    const total = Number((editPix + editEspecie).toFixed(2));

    try {
      if (onAtualizarLancamento) {
        onAtualizarLancamento(modalEditarItem.id, {
          Célula: editCelula,
          celulaNome: editCelula,
          dataBR: editData,
          data: editData,
          valorPix: editPix,
          ValorOferta: editPix,
          valorEspecie: editEspecie,
          OfertaEspecie: editEspecie,
          valorTotal: total,
          Total: total
        });
      }

      // Sincroniza diretamente com o backend
      await fetch('/api/sharepoint/editar-relatorio', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: modalEditarItem.id,
          celula: editCelula,
          data: editData,
          pix: editPix,
          especie: editEspecie,
          total
        })
      });

      exibirAlerta(`Relatório da célula "${editCelula}" atualizado com sucesso.`);
      setModalEditarItem(null);
    } catch (e) {
      console.error('Erro ao salvar edição:', e);
      exibirAlerta('Erro ao salvar alterações no relatório.');
    } finally {
      setSalvandoEdicao(false);
    }
  };

  // Helper para obter a data de hoje formatada em dd/MM/yyyy
  const getHojeDataBR = (): string => {
    const d = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
  };

  // Helper para obter ID do usuário tesoureiro logado (ex: Junio Fonteles - ID: 4)
  const getIdTesoureiroLogado = (): string | number => {
    return usuarioLogado?.id || usuarioLogado?.ID || 4;
  };

  // Confirmar Relatório
  const handleConfirmarItem = async (id: string, celulaNome: string) => {
    const idTesoureiro = getIdTesoureiroLogado();
    const dataHojeBR = getHojeDataBR();

    if (onConfirmarLancamento) {
      onConfirmarLancamento(id, idTesoureiro, dataHojeBR);
    } else {
      try {
        await fetch('/api/sharepoint/validar-relatorio', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            id, 
            recebido: true,
            idTesoureiro,
            dataTesouraria: dataHojeBR
          })
        });
      } catch (e) {
        console.warn('Erro ao sincronizar confirmação no servidor:', e);
      }
    }

    const nomeTesoureiro = usuarioLogado?.nome || 'Junio Fonteles';
    exibirAlerta(`Relatório "${celulaNome}" confirmado por ${nomeTesoureiro} em ${dataHojeBR}.`);
    setSelecionados(prev => {
      const next = new Set(prev);
      next.delete(id);
      return next;
    });
  };

  // Helper para obter o nome do tesoureiro a partir do ID salvo no SharePoint
  const getNomeTesoureiroPorId = (idOuNome: string | number | undefined | null): string => {
    if (!idOuNome) return usuarioLogado?.nome || 'Junio Fonteles';
    const str = String(idOuNome).trim();
    if (!str) return usuarioLogado?.nome || 'Junio Fonteles';

    // Se já for um nome textual
    if (isNaN(Number(str))) return str;

    // Busca na lista de membros do SharePoint
    const spService = SharePointService.getInstance();
    const membros = spService.getMembros();
    const membro = membros.find(m => String(m.id || m.ID) === str);
    if (membro && membro.nome) return membro.nome;

    if (str === '4') return 'Junio Fonteles';
    if (usuarioLogado && String(usuarioLogado.id || usuarioLogado.ID) === str) {
      return usuarioLogado.nome;
    }
    return usuarioLogado?.nome || 'Junio Fonteles';
  };

  // Desfazer Confirmação
  const handleDesfazerItem = async (id: string, celulaNome: string) => {
    if (onDesconfirmarLancamento) {
      onDesconfirmarLancamento(id);
    } else {
      try {
        await fetch('/api/sharepoint/validar-relatorio', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ id, recebido: false })
        });
      } catch (e) {
        console.warn('Erro ao desfazer confirmação no servidor:', e);
      }
    }

    exibirAlerta(`Confirmação da célula "${celulaNome}" desfeita.`);
  };

  // Alternar seleção de item
  const toggleSelecionado = (id: string) => {
    setSelecionados(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  // Selecionar todos os itens da página atual
  const toggleSelecionarTodos = () => {
    if (selecionados.size === relatoriosFiltrados.length && relatoriosFiltrados.length > 0) {
      setSelecionados(new Set());
    } else {
      setSelecionados(new Set(relatoriosFiltrados.map(r => r.id)));
    }
  };

  // Confirmar todos os selecionados
  const handleConfirmarSelecionados = async () => {
    const ids = Array.from(selecionados);
    const idTesoureiro = getIdTesoureiroLogado();
    const dataHojeBR = getHojeDataBR();

    for (const id of ids) {
      if (onConfirmarLancamento) {
        onConfirmarLancamento(id, idTesoureiro, dataHojeBR);
      } else {
        fetch('/api/sharepoint/validar-relatorio', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            id, 
            recebido: true,
            idTesoureiro,
            dataTesouraria: dataHojeBR
          })
        }).catch(() => {});
      }
    }

    const nomeTesoureiro = usuarioLogado?.nome || 'Junio Fonteles';
    exibirAlerta(`${ids.length} relatórios confirmados por ${nomeTesoureiro} em ${dataHojeBR}.`);
    setSelecionados(new Set());
  };

  const isTodosSelecionados = relatoriosFiltrados.length > 0 && selecionados.size === relatoriosFiltrados.length;

  return (
    <div id="validar-relatorios-container" className="p-3 sm:p-6 space-y-4 max-w-[1600px] mx-auto text-slate-100">
      
      {/* Toast Feedback */}
      {mensagemSucesso && (
        <div className="fixed top-5 right-5 z-50 bg-slate-900 border border-emerald-500 text-white px-4 py-3 rounded-lg shadow-2xl flex items-center gap-2.5 text-xs">
          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
          <span>{mensagemSucesso}</span>
        </div>
      )}

      {/* Barra Superior de Controles e Filtros */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#2d334d] pb-3">
        <div className="flex items-center flex-wrap gap-3">
          <h2 className="text-xl font-bold text-white tracking-tight flex items-center gap-2">
            <FileCheck className="w-5 h-5 text-indigo-400" />
            <span>Validar Relatórios</span>
            <span className="text-xs px-2.5 py-0.5 rounded-full bg-indigo-950/80 text-indigo-300 border border-indigo-500/30 font-semibold font-mono">
              Ano {anoSelecionado}
            </span>
          </h2>

          {/* Usuário Tesoureiro Ativo (Exibe somente o nome do Tesoureiro) */}
          <div className="flex items-center gap-2 bg-[#20263c] px-3 py-1 rounded-lg border border-[#313956] text-xs">
            <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
            <span className="text-slate-300">Tesoureiro:</span>
            <strong className="text-white">{usuarioLogado?.nome || 'Junio Fonteles'}</strong>
          </div>
        </div>

        <div className="flex items-center flex-wrap gap-2.5">
          {/* Alternador de Abas: P/ Validar vs Confirmados */}
          <div className="flex items-center bg-[#151724] p-1 rounded-lg border border-[#2c324b]">
            <button
              id="tab-btn-p-validar"
              onClick={() => {
                setTabAtiva('P_VALIDAR');
                setSelecionados(new Set());
              }}
              className={`px-3 py-1.5 rounded-md text-xs font-bold transition-all cursor-pointer flex items-center gap-2 ${
                tabAtiva === 'P_VALIDAR'
                  ? 'bg-amber-600 text-white shadow-md'
                  : 'text-slate-400 hover:text-white hover:bg-[#202538]'
              }`}
            >
              <span>P/ Validar</span>
              <span className="text-[10px] font-mono px-1.5 py-0.2 rounded-full bg-black/40 text-amber-200">
                {totalNaoConfirmadosGeral}
              </span>
            </button>

            <button
              id="tab-btn-confirmados"
              onClick={() => {
                setTabAtiva('CONFIRMADOS');
                setSelecionados(new Set());
              }}
              className={`px-3 py-1.5 rounded-md text-xs font-bold transition-all cursor-pointer flex items-center gap-2 ${
                tabAtiva === 'CONFIRMADOS'
                  ? 'bg-emerald-600 text-white shadow-md'
                  : 'text-slate-400 hover:text-white hover:bg-[#202538]'
              }`}
            >
              <span>Confirmados</span>
              <span className="text-[10px] font-mono px-1.5 py-0.2 rounded-full bg-black/40 text-emerald-200">
                {totalConfirmadosGeral}
              </span>
            </button>
          </div>

          {/* Campo de Busca Rápida */}
          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Buscar célula, líder..."
              value={buscaTexto}
              onChange={(e) => setBuscaTexto(e.target.value)}
              className="bg-[#151724] border border-[#2c324b] text-xs text-white pl-8 pr-3 py-1.5 rounded-lg focus:outline-none focus:border-indigo-400 w-36 sm:w-44 placeholder:text-slate-500"
            />
          </div>

          {/* Botão de Atualizar */}
          <button
            onClick={onRefresh}
            className="p-2 rounded-lg bg-[#22273a] hover:bg-[#2c324b] text-slate-300 hover:text-white border border-[#343b57] transition-colors cursor-pointer"
            title="Recarregar dados"
            aria-label="Recarregar"
          >
            <RotateCw className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Barra de Seleção de Setores (Conforme imagem do usuário) */}
      <div className="flex items-center flex-wrap gap-1 text-xs py-2 px-1 text-slate-300 font-medium">
        <button
          id="filtro-setor-todos"
          onClick={() => setSetorFiltro('todos')}
          className={`px-2.5 py-1 rounded transition-colors cursor-pointer ${
            setorFiltro === 'todos'
              ? 'bg-[#222842] text-white font-bold border border-indigo-500/40 shadow-sm'
              : 'text-slate-300 hover:text-white hover:bg-[#202538]'
          }`}
        >
          Todos ({totalNaoConfirmadosGeral})
        </button>
        <span className="text-slate-600 select-none">|</span>

        {listaSetores.map((setor, index) => {
          const isSelected = setorFiltro.toLowerCase() === setor.toLowerCase();
          const count = naoConfirmadosPorSetor[setor] ?? 0;

          return (
            <React.Fragment key={setor}>
              <button
                id={`filtro-setor-${setor}`}
                onClick={() => setSetorFiltro(isSelected ? 'todos' : setor)}
                className={`px-2.5 py-1 rounded transition-colors cursor-pointer ${
                  isSelected
                    ? 'bg-[#222842] text-white font-bold border border-indigo-500/40 shadow-sm'
                    : 'text-slate-300 hover:text-white hover:bg-[#202538]'
                }`}
              >
                {setor} ({count})
              </button>
              {index < listaSetores.length - 1 && (
                <span className="text-slate-600 select-none">|</span>
              )}
            </React.Fragment>
          );
        })}
      </div>

      {/* Tabela de Relatórios Formatada Exatamente Como a Imagem em Anexo */}
      <div className="rounded-xl overflow-hidden border border-[#2d334d] bg-[#161a29]">
        <div className="overflow-x-auto scrollbar-thin">
          <table className="w-full text-xs text-left min-w-[780px]">
            <thead className="bg-[#151724] text-slate-300 uppercase tracking-wider text-[11px] border-b border-[#2d334d]">
              <tr>
                <th className="px-5 py-3 font-bold">Célula</th>
                <th className="px-4 py-3 font-bold text-center">Data</th>
                <th className="px-4 py-3 font-bold text-right">PIX</th>
                <th className="px-4 py-3 font-bold text-right">Espécie</th>
                <th className="px-4 py-3 font-bold text-center">Total</th>
                <th className="px-3 py-3 font-bold text-center">-</th>
                <th className="px-3 py-3 font-bold text-center w-10">
                  <button 
                    onClick={toggleSelecionarTodos}
                    className="cursor-pointer text-slate-400 hover:text-white"
                    title={isTodosSelecionados ? "Desmarcar todos" : "Selecionar todos"}
                  >
                    {isTodosSelecionados ? (
                      <CheckSquare className="w-4 h-4 text-indigo-400 mx-auto" />
                    ) : (
                      <Square className="w-4 h-4 mx-auto" />
                    )}
                  </button>
                </th>
                <th className="px-4 py-3 font-bold text-center">Ações</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#23283e] p-2">
              {relatoriosFiltrados.length === 0 ? (
                <tr>
                  <td colSpan={8} className="text-center py-12 text-slate-400 text-sm">
                    {lancamentosDoAno.length === 0 
                      ? `Nenhum relatório encontrado para o ano ${anoSelecionado}.`
                      : `Nenhum relatório ${tabAtiva === 'P_VALIDAR' ? 'a validar' : 'confirmado'} no setor ${setorFiltro === 'todos' ? 'selecionado' : setorFiltro}.`}
                  </td>
                </tr>
              ) : (
                relatoriosFiltrados.map((item) => {
                  const pix = item.valorPix ?? (item.ValorOferta ?? 0);
                  const esp = item.valorEspecie ?? (item.OfertaEspecie ?? 0);
                  const total = item.valorTotal ?? (item.Total ?? (pix + esp));
                  const isConfirmado = item.TESOURARIA_RECEB === true;
                  const celulaNome = item.Célula || item.celulaNome || 'Célula';
                  const dataFormatada = getFormatDateBR(item);
                  const isItemSelecionado = selecionados.has(item.id);

                  return (
                    <tr 
                      key={item.id} 
                      className={`transition-colors border-b border-[#23283e] ${
                        isItemSelecionado 
                          ? 'bg-[#cbd5e1] text-slate-950 font-medium hover:bg-[#bcc8d7]' 
                          : 'bg-[#dfe4ec] text-slate-900 hover:bg-[#d2d8e3]'
                      }`}
                    >
                      {/* Célula */}
                      <td className="px-5 py-3.5 font-bold text-slate-900 text-sm">
                        <div>{celulaNome}</div>
                        {isConfirmado && (item.ID_TESOUREIRO || item.DATA_TESOURARIA) && (
                          <div className="text-[11px] font-normal text-slate-600 flex flex-wrap items-center gap-1.5 mt-1">
                            <span className="font-semibold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded text-[11px]">
                              Tesoureiro: {getNomeTesoureiroPorId(item.ID_TESOUREIRO)}
                            </span>
                            {item.DATA_TESOURARIA && (
                              <span className="text-[10px] text-slate-500 font-medium">
                                • Validado em: {item.DATA_TESOURARIA}
                              </span>
                            )}
                          </div>
                        )}
                      </td>

                      {/* Data */}
                      <td className="px-4 py-3.5 text-center font-medium text-slate-800 text-xs">
                        {dataFormatada}
                      </td>

                      {/* PIX */}
                      <td className="px-4 py-3.5 text-right font-bold text-slate-900 text-xs">
                        {formatBRL(pix)}
                      </td>

                      {/* Espécie */}
                      <td className="px-4 py-3.5 text-right font-bold text-slate-900 text-xs">
                        {formatBRL(esp)}
                      </td>

                      {/* Total */}
                      <td className="px-4 py-3.5 text-center">
                        <span className="inline-block font-black px-3 py-1 rounded bg-[#24293f] text-white text-xs shadow-sm">
                          {formatBRL(total)}
                        </span>
                      </td>

                      {/* Traço - */}
                      <td className="px-3 py-3.5 text-center text-slate-600 font-bold">
                        -
                      </td>

                      {/* Checkbox */}
                      <td className="px-3 py-3.5 text-center">
                        <input
                          type="checkbox"
                          checked={isItemSelecionado}
                          onChange={() => toggleSelecionado(item.id)}
                          className="w-4 h-4 rounded border-slate-400 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                        />
                      </td>

                      {/* Botões de Ações: Editar e Confirmar / Desfazer */}
                      <td className="px-4 py-3.5 text-center">
                        <div className="flex items-center justify-center gap-2">
                          {/* Botão Editar (Azul) */}
                          <button
                            onClick={() => handleAbrirEditar(item)}
                            className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                            title="Editar dados deste relatório"
                          >
                            <Edit className="w-3.5 h-3.5" />
                            <span>Editar</span>
                          </button>

                          {/* Botão Confirmar (Verde) */}
                          {!isConfirmado && (
                            <button
                              onClick={() => handleConfirmarItem(item.id, celulaNome)}
                              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                              title="Confirmar relatório"
                            >
                              <Check className="w-3.5 h-3.5" />
                              <span>Confirmar</span>
                            </button>
                          )}

                          {/* Botão Desfazer (Amarelo/Laranja) */}
                          {isConfirmado && (
                            <button
                              onClick={() => handleDesfazerItem(item.id, celulaNome)}
                              className="flex items-center gap-1.5 px-3 py-1.5 rounded bg-amber-600 hover:bg-amber-700 text-white text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                              title="Desfazer confirmação deste relatório"
                            >
                              <Undo2 className="w-3.5 h-3.5" />
                              <span>Desfazer</span>
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Barra Flutuante de Ação em Lote quando há itens selecionados */}
      {selecionados.size > 0 && tabAtiva === 'P_VALIDAR' && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-40 bg-[#161a29] border border-indigo-500/50 shadow-2xl rounded-xl px-5 py-3 flex items-center gap-4 animate-in slide-in-from-bottom-3">
          <span className="text-xs font-bold text-white">
            {selecionados.size} relatório(s) selecionado(s)
          </span>
          <button
            onClick={handleConfirmarSelecionados}
            className="flex items-center gap-1.5 px-3.5 py-1.5 rounded bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shadow transition-colors cursor-pointer"
          >
            <Check className="w-4 h-4" />
            <span>Confirmar Selecionados</span>
          </button>
          <button
            onClick={() => setSelecionados(new Set())}
            className="text-xs text-slate-400 hover:text-white underline cursor-pointer"
          >
            Desmarcar Todos
          </button>
        </div>
      )}

      {/* Modal de Edição de Relatório */}
      {modalEditarItem && (
        <div 
          id="modal-editar-relatorio"
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-xs p-4"
        >
          <div className="bg-[#1c2030] border border-[#2d334d] rounded-xl w-full max-w-md overflow-hidden shadow-2xl animate-in zoom-in-95 duration-200">
            {/* Cabeçalho do Modal */}
            <div className="bg-[#151724] px-5 py-4 border-b border-[#2d334d] flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <Edit className="w-4 h-4 text-blue-400" />
                  <span>Editar Relatório de Célula</span>
                </h3>
                <p className="text-[11px] text-slate-400">
                  ID #{modalEditarItem.id} • Setor: {modalEditarItem.Setor || modalEditarItem.setor || '-'}
                </p>
              </div>
              <button
                onClick={() => setModalEditarItem(null)}
                className="p-1 rounded text-slate-400 hover:text-white hover:bg-[#252a40] transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Formulário de Edição */}
            <div className="p-5 space-y-4 text-xs">
              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Nome da Célula
                </label>
                <input
                  type="text"
                  value={editCelula}
                  onChange={(e) => setEditCelula(e.target.value)}
                  className="w-full bg-[#151724] border border-[#2c324b] text-white px-3 py-2 rounded-lg focus:outline-none focus:border-blue-500"
                />
              </div>

              <div>
                <label className="block text-slate-300 font-semibold mb-1">
                  Data da Célula (DD/MM/AAAA)
                </label>
                <input
                  type="text"
                  value={editData}
                  onChange={(e) => setEditData(e.target.value)}
                  placeholder="Ex: 10/09/2026"
                  className="w-full bg-[#151724] border border-[#2c324b] text-white px-3 py-2 rounded-lg focus:outline-none focus:border-blue-500 font-mono"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Valor PIX (R$)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={editPix}
                    onChange={(e) => setEditPix(parseFloat(e.target.value) || 0)}
                    className="w-full bg-[#151724] border border-[#2c324b] text-white px-3 py-2 rounded-lg focus:outline-none focus:border-blue-500 font-mono font-bold"
                  />
                </div>

                <div>
                  <label className="block text-slate-300 font-semibold mb-1">
                    Valor Espécie (R$)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={editEspecie}
                    onChange={(e) => setEditEspecie(parseFloat(e.target.value) || 0)}
                    className="w-full bg-[#151724] border border-[#2c324b] text-white px-3 py-2 rounded-lg focus:outline-none focus:border-blue-500 font-mono font-bold"
                  />
                </div>
              </div>

              {/* Card de Total Calculado */}
              <div className="bg-[#151724] border border-[#2c324b] p-3 rounded-lg flex items-center justify-between">
                <span className="text-slate-400 font-medium">Total Calculado:</span>
                <span className="text-base font-black text-white font-mono">
                  {formatBRL(editPix + editEspecie)}
                </span>
              </div>
            </div>

            {/* Rodapé do Modal */}
            <div className="bg-[#151724] px-5 py-3 border-t border-[#2d334d] flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setModalEditarItem(null)}
                className="px-3.5 py-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-[#252a40] transition-colors cursor-pointer"
              >
                Cancelar
              </button>

              <button
                type="button"
                onClick={handleSalvarEdicao}
                disabled={salvandoEdicao}
                className="flex items-center gap-1.5 px-4 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-bold shadow-md transition-colors cursor-pointer disabled:opacity-50"
              >
                {salvandoEdicao ? (
                  <RotateCw className="w-4 h-4 animate-spin" />
                ) : (
                  <Save className="w-4 h-4" />
                )}
                <span>Salvar Alterações</span>
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};
