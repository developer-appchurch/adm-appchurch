import React, { useState, useMemo, useEffect, useCallback } from 'react';
import { Calendar, CheckCircle, CheckCircle2, Clock, Copy, X, AlertTriangle } from 'lucide-react';
import { LancamentoTesouraria } from '../../types';
import { SharePointService } from '../../services/sharepointService';

interface RelacaoEnvelopesViewProps {
  lancamentos: LancamentoTesouraria[];
  anoSelecionado: number;
  onSelectAno: (ano: number) => void;
  mesSelecionado?: string;
  onSelectMes?: (mes: string) => void;
  setorSelecionado?: string;
  onSelectSetor?: (setor: string) => void;
  onSetoresDisponiveisChange?: (setores: string[]) => void;
  onRefresh: () => void;
}

const MESES = [
  { valor: 'todos', label: 'Todos os Meses' },
  { valor: '1', label: 'Janeiro' },
  { valor: '2', label: 'Fevereiro' },
  { valor: '3', label: 'Março' },
  { valor: '4', label: 'Abril' },
  { valor: '5', label: 'Maio' },
  { valor: '6', label: 'Junho' },
  { valor: '7', label: 'Julho' },
  { valor: '8', label: 'Agosto' },
  { valor: '9', label: 'Setembro' },
  { valor: '10', label: 'Outubro' },
  { valor: '11', label: 'Novembro' },
  { valor: '12', label: 'Dezembro' }
];

export const RelacaoEnvelopesView: React.FC<RelacaoEnvelopesViewProps> = ({
  lancamentos,
  anoSelecionado,
  onSelectAno,
  mesSelecionado: mesSelecionadoProp,
  onSelectMes: onSelectMesProp,
  setorSelecionado: setorSelecionadoProp,
  onSelectSetor: onSelectSetorProp,
  onSetoresDisponiveisChange,
  onRefresh
}) => {
  const spService = useMemo(() => SharePointService.getInstance(), []);
  const [internalSetor, setInternalSetor] = useState<string>('Safira');
  const [internalMes, setInternalMes] = useState<string>('9'); // Padrão Setembro (mês atual) ou 'todos'

  const setorSelecionado = setorSelecionadoProp ?? internalSetor;
  const setSetorSelecionado = onSelectSetorProp ?? setInternalSetor;
  const mesSelecionado = mesSelecionadoProp ?? internalMes;
  const setMesSelecionado = onSelectMesProp ?? setInternalMes;

  const [celulasBD, setCelulasBD] = useState<any[]>(spService.getCelulas());
  const [membrosBD, setMembrosBD] = useState<any[]>(spService.getMembros());
  const [isCarregando, setIsCarregando] = useState<boolean>(false);
  const [modalDuplicadosAberto, setModalDuplicadosAberto] = useState<boolean>(false);

  // Carregar células e membros diretamente do SharePoint / Cache
  const carregarCelulasEMembros = useCallback(async () => {
    try {
      setIsCarregando(true);
      // Tenta sincronizar se necessário
      const [resCel, resMem] = await Promise.all([
        fetch('/api/sharepoint/celulas').then(r => r.ok ? r.json() : null).catch(() => null),
        fetch('/api/sharepoint/membros').then(r => r.ok ? r.json() : null).catch(() => null)
      ]);

      if (resCel && Array.isArray(resCel.celulas) && resCel.celulas.length > 0) {
        setCelulasBD(resCel.celulas);
        spService.salvarCelulas(resCel.celulas);
      } else {
        setCelulasBD(spService.getCelulas());
      }

      if (resMem && Array.isArray(resMem.membros) && resMem.membros.length > 0) {
        setMembrosBD(resMem.membros);
        spService.salvarMembros(resMem.membros);
      } else {
        setMembrosBD(spService.getMembros());
      }
    } finally {
      setIsCarregando(false);
    }
  }, [spService]);

  useEffect(() => {
    carregarCelulasEMembros();
  }, [carregarCelulasEMembros]);

  // Handler de atualização sincronizado
  const handleAtualizar = async () => {
    await carregarCelulasEMembros();
    onRefresh();
  };

  // Mapa rápido de ID_Lider -> Nome do Líder em BD_Membros
  const membrosMap = useMemo(() => {
    const map: Record<string, string> = {};
    membrosBD.forEach(m => {
      const id = String(m.id || m.ID || '').trim();
      const nome = String(m.nome || m.Title || '').trim();
      if (id && nome) {
        map[id] = nome;
      }
    });
    return map;
  }, [membrosBD]);

  // Setores disponíveis obtidos de BD_celulas e relatórios
  const setoresDisponiveis = useMemo(() => {
    const setoresDasCelulas = Array.from(
      new Set(
        celulasBD
          .map(c => String(c.Setor || c.setor || '').trim())
          .filter(Boolean)
      )
    ).sort();

    if (setoresDasCelulas.length > 0) {
      return setoresDasCelulas;
    }

    // Fallback: setores dos lançamentos ou padrão
    const setoresLancs = Array.from(
      new Set(lancamentos.map(l => String(l.Setor || l.setor || '').trim()).filter(Boolean))
    ).sort();

    return setoresLancs.length > 0 
      ? setoresLancs 
      : ['Safira', 'Fire', 'White', 'Azul', 'Amarelo', 'Black', 'Diamante', 'Legacy', 'Onix', 'Titanium'];
  }, [celulasBD, lancamentos]);

  // Notifica componentes pais sobre os setores disponíveis
  useEffect(() => {
    if (onSetoresDisponiveisChange && setoresDisponiveis.length > 0) {
      onSetoresDisponiveisChange(setoresDisponiveis);
    }
  }, [setoresDisponiveis, onSetoresDisponiveisChange]);

  // Ajusta o setor selecionado para um válido caso o atual não exista na lista
  useEffect(() => {
    if (setoresDisponiveis.length > 0 && !setoresDisponiveis.includes(setorSelecionado)) {
      setSetorSelecionado(setoresDisponiveis[0]);
    }
  }, [setoresDisponiveis, setorSelecionado, setSetorSelecionado]);

  // Células pertencentes ao setor selecionado extraídas da tabela/lista BD_celulas
  // Coluna de líder resolvida via ID_Lider / LiderCelula contra BD_Membros
  const celulasDoSetor = useMemo(() => {
    if (!setorSelecionado) return [];

    const filtradas = celulasBD.filter(c => {
      const s = String(c.Setor || c.setor || '').trim().toLowerCase();
      return s === setorSelecionado.trim().toLowerCase();
    });

    if (filtradas.length > 0) {
      return filtradas.map(c => {
        const nomeCelula = String(c.Celula || c.nome || c.Célula || 'Célula').trim();

        // Onde mostra o nome do líder, tem uma coluna ID_Lider (ou LiderCelula / Id_Lider):
        // busca o ID conferindo na tabela/lista BD_Membros e traz apenas o nome do líder
        const idLiderRaw = String(c.ID_Lider || c.Id_Lider || c.id_lider || c.LiderCelula || '').trim();
        let nomeLider = '';

        if (idLiderRaw) {
          // Trata caso venha com múltiplos IDs separados por ponto-e-vírgula ou vírgula
          const ids = idLiderRaw.split(/[;,]+/).map(s => s.trim()).filter(Boolean);
          const nomesEncontrados = ids
            .map(id => membrosMap[id] || membrosMap[String(Number(id))] || null)
            .filter(Boolean);

          if (nomesEncontrados.length > 0) {
            nomeLider = nomesEncontrados.join(' / ');
          } else if (isNaN(Number(ids[0]))) {
            // Se o campo já contiver o nome textual
            nomeLider = idLiderRaw;
          }
        }

        if (!nomeLider) {
          nomeLider = c.lider || c.LíderCelula || '-';
        }

        return {
          id: c.ID || c.Id || c.id || nomeCelula,
          nome: nomeCelula,
          lider: nomeLider,
          setor: c.Setor || c.setor || setorSelecionado
        };
      }).sort((a, b) => a.nome.localeCompare(b.nome));
    }

    // Fallback: se BD_celulas ainda não estiver carregado, busca dos relatórios daquele setor
    const lancsDoSetor = lancamentos.filter(
      l => (l.Setor || l.setor || '').trim().toLowerCase() === setorSelecionado.trim().toLowerCase()
    );
    const nomesCelulasUnicas = Array.from(
      new Set(lancsDoSetor.map(l => l.Célula || l.celulaNome).filter(Boolean))
    ).sort();

    return nomesCelulasUnicas.map(nome => {
      const itemExemplo = lancsDoSetor.find(l => (l.Célula || l.celulaNome) === nome);
      return {
        id: itemExemplo?.id || nome,
        nome,
        lider: itemExemplo?.LíderCelula || '-',
        setor: setorSelecionado
      };
    });
  }, [celulasBD, setorSelecionado, membrosMap, lancamentos]);

  // Segundo Bloco: Envelopes Validados e Envelopes Pendentes
  // Seguindo rigorosamente os filtros de Setor, Mês e Ano selecionados
  const resumoSKUs = useMemo(() => {
    const lancsDoSetor = lancamentos.filter(l => {
      const matchSetor = (l.Setor || l.setor || '').trim().toLowerCase() === setorSelecionado.trim().toLowerCase();
      const matchAno = !anoSelecionado || l.ano === anoSelecionado || (l.dataBR && l.dataBR.endsWith(`/${anoSelecionado}`));
      const matchMes = mesSelecionado === 'todos' || l.mes === Number(mesSelecionado);
      return matchSetor && matchAno && matchMes;
    });

    // Envelopes Validados = Qtd de Envelopes Validados daquele setor (TESOURARIA_RECEB === true)
    const validados = lancsDoSetor.filter(l => l.TESOURARIA_RECEB === true);

    // Envelopes Pendentes = Qtd de Relatórios lançados no aplicativo que a tesouraria não recebeu/validou (TESOURARIA_RECEB !== true)
    const pendentes = lancsDoSetor.filter(l => l.TESOURARIA_RECEB !== true);

    return {
      qtdValidados: validados.length,
      qtdPendentes: pendentes.length,
      totalLancados: lancsDoSetor.length
    };
  }, [lancamentos, setorSelecionado, anoSelecionado, mesSelecionado]);

  // Semanas dinâmicas da tabela:
  // Regra do Usuário:
  // Sempre mostra 5 semanas de cada mês.
  // A última coluna (5ª) no campo data vem selecionado o último sábado de cada mês.
  // Cada coluna é uma semana anterior à coluna da direita:
  // Coluna 5 = último sábado do mês
  // Coluna 4 = 1 semana antes
  // Coluna 3 = 2 semanas antes
  // Coluna 2 = 3 semanas antes
  // Coluna 1 = 4 semanas antes
  const semanas = useMemo(() => {
    // Ano base
    const ano = anoSelecionado || 2026;
    // Mês base (1 a 12). Se 'todos', adota o mês atual ou setembro (mês 9)
    const mesNum = (mesSelecionado && mesSelecionado !== 'todos') ? parseInt(mesSelecionado, 10) : 9;

    // Achar o último sábado do mês:
    // O último dia do mês é dado por new Date(ano, mesNum, 0)
    const ultimoDiaDoMes = new Date(ano, mesNum, 0); // ex: para mesNum=9 (Setembro), dia 30
    const diaSemanaUltimoDia = ultimoDiaDoMes.getDay(); // 0=Dom, 1=Seg, ..., 6=Sáb
    
    // Distância para voltar até o sábado:
    // Se for sábado (6) => volta 0 dias
    // Se for domingo (0) => volta 1 dia
    // Se for sexta (5) => volta 6 dias
    // fórmula: (diaSemanaUltimoDia - 6 + 7) % 7
    const diasParaSubtrair = (diaSemanaUltimoDia - 6 + 7) % 7;
    const dataUltimoSabado = new Date(ano, mesNum - 1, ultimoDiaDoMes.getDate() - diasParaSubtrair);

    // Função utilitária para obter número da semana no ano (ISO / padrão)
    const getWeekNumber = (d: Date): number => {
      const target = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      const dayNr = (target.getUTCDay() + 6) % 7;
      target.setUTCDate(target.getUTCDate() - dayNr + 3);
      const firstThursday = target.getTime();
      target.setUTCMonth(0, 1);
      if (target.getUTCDay() !== 4) {
        target.setUTCMonth(0, 1 + ((4 - target.getUTCDay()) + 7) % 7);
      }
      return 1 + Math.ceil((firstThursday - target.getTime()) / 604800000);
    };

    // Gera as 5 datas correspondentes aos sábados (da coluna 1 à coluna 5, da esquerda para a direita)
    // Coluna 1 = dataUltimoSabado - 28 dias (i = 4)
    // Coluna 2 = dataUltimoSabado - 21 dias (i = 3)
    // Coluna 3 = dataUltimoSabado - 14 dias (i = 2)
    // Coluna 4 = dataUltimoSabado - 7 dias  (i = 1)
    // Coluna 5 = dataUltimoSabado - 0 dias  (i = 0)
    const listaSemanas = [];
    for (let step = 4; step >= 0; step--) {
      const d = new Date(dataUltimoSabado);
      d.setDate(d.getDate() - step * 7);

      const diaStr = String(d.getDate()).padStart(2, '0');
      const mesStr = String(d.getMonth() + 1).padStart(2, '0');
      const anoStr = String(d.getFullYear());
      const dataFormatada = `${diaStr}/${mesStr}/${anoStr}`;
      const dataIso = `${anoStr}-${mesStr}-${diaStr}`;
      const numSemana = getWeekNumber(d);

      listaSemanas.push({
        num: numSemana,
        data: dataFormatada,
        dataIso,
        dataObj: d
      });
    }

    return listaSemanas;
  }, [anoSelecionado, mesSelecionado]);

  // Identificação e contagem de Relatórios Duplicados para as semanas exibidas naquele mês, ano e setor selecionado:
  // Exemplo da solicitação: Mês Agosto, Setor Safira -> Célula Efratá tem 2 relatórios duplicados (semana 32 e 34)
  const duplicadosInfo = useMemo(() => {
    const mesNumFiltro = (mesSelecionado && mesSelecionado !== 'todos') ? Number(mesSelecionado) : null;
    const anoNumFiltro = anoSelecionado ? Number(anoSelecionado) : 2026;

    // Números das semanas exibidas na tela do mês atual
    const numerosSemanasTela = new Set(semanas.map(s => s.num));

    // 1. Filtrar lançamentos que pertencem estritamente ao setor, ano e ao mês selecionado
    const lancsDoSetorEMes = lancamentos.filter(l => {
      // Setor
      const matchSetor = (l.Setor || l.setor || '').trim().toLowerCase() === setorSelecionado.trim().toLowerCase();
      if (!matchSetor) return false;

      // Ano
      const lancAno = l.ano || (l.dataBR && l.dataBR.includes('/') ? Number(l.dataBR.split('/')[2]) : null);
      if (lancAno && lancAno !== anoNumFiltro) return false;

      // Mês: verifica se bate com o mês selecionado
      if (mesNumFiltro !== null) {
        const lancMes = l.mes || (l.dataBR && l.dataBR.includes('/') ? Number(l.dataBR.split('/')[1]) : null);
        // Também permite se o número da semana do relatório coincide diretamente com as semanas do mês exibidas
        const semNumLanc = l.NumSemana ?? l.semanaNumero;
        const bateComSemanaDaTela = typeof semNumLanc === 'number' && numerosSemanasTela.has(semNumLanc);

        if (lancMes && lancMes !== mesNumFiltro && !bateComSemanaDaTela) {
          return false;
        }
      }

      return true;
    });

    // Para cada semana das 5 semanas da tela e para cada célula:
    interface GrupoDuplicado {
      chave: string;
      semanaNum: number;
      semanaData: string;
      celulaNome: string;
      relatorios: LancamentoTesouraria[];
      qtdDuplicados: number; // relatorios.length - 1
    }

    const grupos: GrupoDuplicado[] = [];
    let totalRelatoriosDuplicados = 0;

    // Agrupar lançamentos para cada par (Semana, Célula)
    semanas.forEach(semInfo => {
      const relsPorCelula: Record<string, LancamentoTesouraria[]> = {};

      lancsDoSetorEMes.forEach(l => {
        const celNome = (l.Célula || l.celulaNome || '').trim();
        if (!celNome) return;

        // Determina o número da semana do relatório
        let semNumLanc: number | null = null;
        if (typeof l.NumSemana === 'number') {
          semNumLanc = l.NumSemana;
        } else if (typeof l.NumSemana === 'string') {
          const m = l.NumSemana.match(/\d+/);
          if (m) semNumLanc = parseInt(m[0], 10);
        }
        if (semNumLanc === null && typeof l.semanaNumero === 'number') {
          semNumLanc = l.semanaNumero;
        }

        // Verifica correspondência com a semana atual da tela:
        // 1. Número exato da semana (NumSemana / semanaNumero)
        let pertenceASemana = false;
        if (semNumLanc !== null && semNumLanc === semInfo.num) {
          pertenceASemana = true;
        }

        // 2. Data exata coincidente
        if (!pertenceASemana) {
          if (l.dataBR === semInfo.data || l.data === semInfo.dataIso || l.DataCelula === semInfo.dataIso) {
            pertenceASemana = true;
          }
        }

        if (pertenceASemana) {
          const chaveNorm = celNome.toLowerCase();
          if (!relsPorCelula[chaveNorm]) {
            relsPorCelula[chaveNorm] = [];
          }
          relsPorCelula[chaveNorm].push(l);
        }
      });

      // Identifica células que possuem mais de 1 relatório lançado na mesma semana
      Object.entries(relsPorCelula).forEach(([, rels]) => {
        if (rels.length > 1) {
          const celNomePrincipal = rels[0].Célula || rels[0].celulaNome || 'Célula';
          grupos.push({
            chave: `${semInfo.num}-${celNomePrincipal}`,
            semanaNum: semInfo.num,
            semanaData: semInfo.data,
            celulaNome: celNomePrincipal,
            relatorios: rels,
            qtdDuplicados: rels.length - 1
          });
          // Soma a quantidade de relatórios excedentes duplicados
          totalRelatoriosDuplicados += (rels.length - 1);
        }
      });
    });

    return {
      totalDuplicados: totalRelatoriosDuplicados,
      grupos
    };
  }, [lancamentos, setorSelecionado, anoSelecionado, mesSelecionado, semanas]);

  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(val || 0);
  };

  // Helper para buscar lançamento da célula na semana e verificar TESOURARIA_RECEB
  // Verifica por número da semana, data exata ou proximidade do sábado (mesma semana)
  const getValores = (celulaNome: string, semInfo: { num: number; data: string; dataIso: string; dataObj: Date }) => {
    const lanc = lancamentos.find(l => {
      const matchNome = (l.Célula || l.celulaNome || '').trim().toLowerCase() === celulaNome.trim().toLowerCase();
      if (!matchNome) return false;

      // 1. Confere por data exata (DD/MM/YYYY ou YYYY-MM-DD)
      if (l.dataBR === semInfo.data || l.data === semInfo.dataIso || l.DataCelula === semInfo.dataIso) {
        return true;
      }

      // 2. Confere por NumSemana
      const sem = l.NumSemana ?? l.semanaNumero;
      const matchAno = !anoSelecionado || l.ano === anoSelecionado || (l.dataBR && l.dataBR.endsWith(`/${anoSelecionado}`));
      if (sem === semInfo.num && matchAno) {
        return true;
      }

      // 3. Confere por proximidade de data (mesma semana de sábado a domingo)
      if (l.data || l.dataBR) {
        let lancDate: Date | null = null;
        if (l.data && l.data.includes('-')) {
          const parts = l.data.split('-');
          lancDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        } else if (l.dataBR && l.dataBR.includes('/')) {
          const parts = l.dataBR.split('/');
          lancDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
        }

        if (lancDate && !isNaN(lancDate.getTime())) {
          const diffDays = Math.abs((lancDate.getTime() - semInfo.dataObj.getTime()) / (1000 * 60 * 60 * 24));
          if (diffDays <= 4) {
            return true;
          }
        }
      }

      return false;
    });

    if (!lanc) {
      return { temDado: false, pix: null, dinheiro: null, validadoTesouraria: false, total: 0 };
    }

    const pix = lanc.ValorOferta ?? lanc.valorPix ?? 0;
    const esp = lanc.OfertaEspecie ?? lanc.valorEspecie ?? 0;
    const total = lanc.Total ?? lanc.valorTotal ?? (pix + esp);

    return {
      temDado: true,
      pix,
      dinheiro: esp,
      total,
      validadoTesouraria: lanc.TESOURARIA_RECEB === true
    };
  };

  return (
    <div id="relacao-envelopes-container" className="p-3.5 sm:p-6 space-y-4 sm:space-y-5 max-w-[1600px] mx-auto text-slate-100">
      {/* 1. Primeiro Bloco: SKUs (Envelopes Validados, Envelopes Pendentes e Relatório Duplicado) */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5 sm:gap-4">
        {/* Card 1: Envelopes Validados */}
        <div 
          id="card-envelopes-validados"
          className="bg-[#202538] p-4 rounded-xl border border-emerald-500/30 flex items-center justify-between shadow-md"
        >
          <div>
            <span className="text-slate-400 text-xs font-medium block">Envelopes Validados</span>
            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-2xl sm:text-3xl font-black text-emerald-400 font-mono">
                {resumoSKUs.qtdValidados}
              </span>
              <span className="text-xs font-medium text-slate-400">
                de {resumoSKUs.totalLancados} relatórios do setor
              </span>
            </div>
            <p className="text-[10px] text-slate-500 mt-0.5">
              Recebidos e confirmados pela tesouraria
            </p>
          </div>
          <div className="p-3 rounded-xl bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
            <CheckCircle2 className="w-6 h-6" />
          </div>
        </div>

        {/* Card 2: Envelopes Pendentes */}
        <div 
          id="card-envelopes-pendentes"
          className="bg-[#202538] p-4 rounded-xl border border-amber-500/30 flex items-center justify-between shadow-md"
        >
          <div>
            <span className="text-slate-400 text-xs font-medium block">Envelopes Pendentes</span>
            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-2xl sm:text-3xl font-black text-amber-400 font-mono">
                {resumoSKUs.qtdPendentes}
              </span>
              <span className="text-xs font-medium text-slate-400">
                aguardando validação
              </span>
            </div>
            <p className="text-[10px] text-slate-500 mt-0.5">
              Lançados no app, aguardando recebimento na tesouraria
            </p>
          </div>
          <div className="p-3 rounded-xl bg-amber-500/10 text-amber-400 border border-amber-500/20">
            <Clock className="w-6 h-6" />
          </div>
        </div>

        {/* Card 3: Relatório Duplicado */}
        <div 
          id="card-relatorios-duplicados"
          onClick={() => {
            if (duplicadosInfo.totalDuplicados > 0) {
              setModalDuplicadosAberto(true);
            }
          }}
          className={`p-4 rounded-xl border flex items-center justify-between shadow-md transition-all ${
            duplicadosInfo.totalDuplicados > 0
              ? 'bg-[#202538] border-rose-500/50 hover:border-rose-500 hover:bg-[#252b41] cursor-pointer ring-1 ring-rose-500/20 active:scale-[0.99]'
              : 'bg-[#202538] border-slate-700/60 opacity-90'
          }`}
          title={
            duplicadosInfo.totalDuplicados > 0
              ? 'Clique para visualizar detalhes dos relatórios duplicados'
              : 'Nenhum relatório duplicado encontrado'
          }
        >
          <div>
            <div className="flex items-center gap-1.5">
              <span className="text-slate-400 text-xs font-medium block">Relatório Duplicado</span>
              {duplicadosInfo.totalDuplicados > 0 && (
                <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[9px] font-bold bg-rose-500/20 text-rose-300 border border-rose-500/30">
                  Atenção
                </span>
              )}
            </div>
            <div className="flex items-baseline gap-2 mt-1">
              <span className={`text-2xl sm:text-3xl font-black font-mono ${
                duplicadosInfo.totalDuplicados > 0 ? 'text-rose-400' : 'text-slate-300'
              }`}>
                {duplicadosInfo.totalDuplicados}
              </span>
              <span className="text-xs font-medium text-slate-400">
                {duplicadosInfo.totalDuplicados === 1 ? 'duplicado detectado' : 'duplicados detectados'}
              </span>
            </div>
            <p className="text-[10px] text-slate-500 mt-0.5">
              {duplicadosInfo.totalDuplicados > 0 
                ? 'Mesma célula com múltiplos relatórios na semana (clique p/ ver)'
                : 'Nenhum relatório duplicado nas semanas deste mês'}
            </p>
          </div>
          <div className={`p-3 rounded-xl border ${
            duplicadosInfo.totalDuplicados > 0
              ? 'bg-rose-500/10 text-rose-400 border-rose-500/30'
              : 'bg-slate-700/20 text-slate-400 border-slate-700/30'
          }`}>
            <Copy className="w-6 h-6" />
          </div>
        </div>
      </div>

      {/* 3. Terceiro Bloco: Tabela de Células de BD_celulas e Semana a Semana */}
      <div className="bg-[#e4e7ed] text-slate-900 rounded-xl shadow-md overflow-hidden border border-[#c5cbda]">
        <div className="overflow-x-auto scrollbar-thin">
          <table className="w-full text-xs border-collapse min-w-[760px]">
            <thead>
              <tr className="bg-[#1c2030] text-white">
                <th className="py-3 px-4 text-left font-bold text-sm w-64 border-r border-[#2f354e]">
                  Célula / Líder
                </th>
                <th className="py-3 px-2 text-center font-bold w-24 border-r border-[#2f354e]">
                  Oferta
                </th>
                {semanas.map((sem, sIdx) => (
                  <th key={`header-sem-${sem.data}-${sIdx}`} className="py-2.5 px-3 text-center border-r border-[#2f354e] last:border-r-0 min-w-[130px]">
                    <div className="text-[11px] font-semibold text-slate-200">Semana {sem.num}</div>
                    <div className="mt-1 flex items-center justify-center gap-1 bg-[#252b41] py-0.5 px-2 rounded text-[11px] font-mono text-slate-300 border border-[#3a4364]">
                      <span>{sem.data}</span>
                      <Calendar className="w-3 h-3 text-slate-400 inline shrink-0" />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>

            <tbody>
              {celulasDoSetor.length === 0 ? (
                <tr>
                  <td colSpan={2 + semanas.length} className="py-12 text-center text-slate-600 font-medium">
                    Nenhuma célula cadastrada encontrada na tabela <span className="font-semibold text-slate-800">BD_celulas</span> para o setor <strong className="text-slate-900">{setorSelecionado}</strong>.
                  </td>
                </tr>
              ) : (
                celulasDoSetor.map((celula) => {
                  return (
                    <React.Fragment key={celula.id}>
                      {/* Linha 1: PIX */}
                      <tr className="border-t-2 border-[#b0b8cc]">
                        <td rowSpan={2} className="py-2.5 px-3.5 font-bold text-xs text-slate-900 align-middle bg-[#f3f5f9] border-r border-[#b0b8cc]">
                          <p className="font-bold text-slate-900 text-[13px] leading-tight truncate max-w-[220px]">{celula.nome}</p>
                          <p className="text-[11px] text-slate-600 font-medium leading-tight mt-1 truncate max-w-[220px]">
                            Líder: <span className="text-slate-800 font-semibold">{celula.lider || '-'}</span>
                          </p>
                        </td>

                        <td className="py-1.5 px-2 text-center font-extrabold bg-[#1c2030] text-white text-[10px] tracking-wider border-b border-[#2e344e] border-r border-[#b0b8cc] h-8">
                          PIX
                        </td>

                        {semanas.map((sem, sIdx) => {
                          const vals = getValores(celula.nome, sem);
                          const hasRelatorio = vals.temDado;
                          const isValidado = vals.validadoTesouraria;
                          const valorPix = vals.pix ?? 0;

                          return (
                            <td 
                              key={`pix-${sem.data}-${sIdx}`} 
                              className={`py-2 px-2 text-center border-r border-[#b0b8cc] last:border-r-0 border-b border-[#ccd2e0] h-8 transition-colors ${
                                isValidado
                                  ? 'bg-[#eef1f6]'
                                  : 'bg-[#fecdd3]/60'
                              }`}
                            >
                              {hasRelatorio ? (
                                <div className="flex items-center justify-center gap-1.5">
                                  <span className={`font-bold text-[11.5px] ${isValidado ? 'text-slate-900 font-mono' : 'text-rose-950 font-mono'}`}>
                                    {formatBRL(valorPix)}
                                  </span>
                                  {isValidado && (
                                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600 inline shrink-0" />
                                  )}
                                </div>
                              ) : (
                                // Não lançado relatório naquela semana: fica apenas o fundo no tom de vermelho claro
                                <span className="text-rose-400/70 text-[11px] font-mono select-none">-</span>
                              )}
                            </td>
                          );
                        })}
                      </tr>

                      {/* Linha 2: ESPÉCIE */}
                      <tr className="bg-[#e4e7ed]">
                        <td className="py-1.5 px-2 text-center font-extrabold bg-[#1c2030] text-white text-[10px] tracking-wider border-r border-[#b0b8cc] h-8">
                          ESPÉCIE
                        </td>

                        {semanas.map((sem, sIdx) => {
                          const vals = getValores(celula.nome, sem);
                          const hasRelatorio = vals.temDado;
                          const isValidado = vals.validadoTesouraria;
                          const valorEspecie = vals.dinheiro ?? 0;

                          return (
                            <td 
                              key={`esp-${sem.data}-${sIdx}`} 
                              className={`py-2 px-2 text-center border-r border-[#b0b8cc] last:border-r-0 h-8 transition-colors ${
                                isValidado
                                  ? 'bg-[#e4e7ed]'
                                  : 'bg-[#fecdd3]/60'
                              }`}
                            >
                              {hasRelatorio ? (
                                <div className="flex items-center justify-center gap-1.5">
                                  <span className={`font-bold text-[11.5px] ${isValidado ? 'text-slate-900 font-mono' : 'text-rose-950 font-mono'}`}>
                                    {formatBRL(valorEspecie)}
                                  </span>
                                  {isValidado && (
                                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600 inline shrink-0" />
                                  )}
                                </div>
                              ) : (
                                // Não lançado relatório naquela semana: fica apenas o fundo no tom de vermelho claro
                                <span className="text-rose-400/70 text-[11px] font-mono select-none">-</span>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    </React.Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Modal de Relatórios Duplicados */}
      {modalDuplicadosAberto && (
        <div 
          id="modal-relatorios-duplicados"
          className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-fadeIn"
          onClick={() => setModalDuplicadosAberto(false)}
        >
          <div 
            className="bg-[#1c2030] text-slate-100 rounded-2xl max-w-3xl w-full max-h-[85vh] flex flex-col shadow-2xl border border-rose-500/40 overflow-hidden"
            onClick={e => e.stopPropagation()}
          >
            {/* Cabeçalho do Modal */}
            <div className="p-4 sm:p-5 border-b border-slate-700/80 flex items-center justify-between bg-[#202538]">
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-xl bg-rose-500/20 text-rose-400 border border-rose-500/30">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-bold text-white flex items-center gap-2">
                    Relatórios Duplicados no Setor {setorSelecionado}
                  </h3>
                  <p className="text-xs text-slate-400 mt-0.5">
                    {duplicadosInfo.totalDuplicados} {duplicadosInfo.totalDuplicados === 1 ? 'relatório excedente' : 'relatórios excedentes'} em {duplicadosInfo.grupos.length} {duplicadosInfo.grupos.length === 1 ? 'célula/semana' : 'células/semanas'}
                  </p>
                </div>
              </div>
              <button
                id="btn-fechar-modal-duplicados"
                onClick={() => setModalDuplicadosAberto(false)}
                className="p-2 text-slate-400 hover:text-white hover:bg-slate-700/50 rounded-lg transition-colors"
                title="Fechar"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Conteúdo com os Grupos de Duplicados */}
            <div className="p-4 sm:p-6 overflow-y-auto space-y-5 scrollbar-thin">
              {duplicadosInfo.grupos.length === 0 ? (
                <div className="text-center py-10 text-slate-400 text-sm">
                  Nenhum relatório duplicado encontrado para este período e setor.
                </div>
              ) : (
                duplicadosInfo.grupos.map((grupo, gIdx) => (
                  <div 
                    key={`grupo-${grupo.chave}-${gIdx}`} 
                    className="bg-[#242a3e] rounded-xl border border-rose-500/30 p-4 shadow-sm"
                  >
                    {/* Topo do Grupo: Célula e Semana */}
                    <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-700/70">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-sm text-white">{grupo.celulaNome}</span>
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-500/20 text-rose-300 border border-rose-500/30">
                            {grupo.relatorios.length} lançamentos ({grupo.qtdDuplicados} duplicado{grupo.qtdDuplicados > 1 ? 's' : ''})
                          </span>
                        </div>
                        <p className="text-xs text-slate-400 mt-0.5">
                          Setor: <strong className="text-slate-200">{setorSelecionado}</strong>
                        </p>
                      </div>

                      <div className="flex items-center gap-2 bg-[#1c2030] px-3 py-1.5 rounded-lg border border-slate-700">
                        <Calendar className="w-3.5 h-3.5 text-slate-400" />
                        <span className="text-xs font-semibold text-slate-200">Semana {grupo.semanaNum}</span>
                        <span className="text-[11px] text-slate-400 font-mono">({grupo.semanaData})</span>
                      </div>
                    </div>

                    {/* Lista dos relatórios duplicados lançados para essa célula naquela semana */}
                    <div className="mt-3 space-y-2">
                      <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">
                        Relatórios encontrados nesta semana:
                      </p>
                      <div className="grid grid-cols-1 gap-2">
                        {grupo.relatorios.map((rel, rIdx) => {
                          const pix = rel.ValorOferta ?? rel.valorPix ?? 0;
                          const esp = rel.OfertaEspecie ?? rel.valorEspecie ?? 0;
                          const total = rel.Total ?? rel.valorTotal ?? (pix + esp);
                          const lider = rel.LíderCelula || '-';
                          const idRel = rel.ID || rel.id || `#${rIdx + 1}`;
                          const isVal = rel.TESOURARIA_RECEB === true;
                          const dataRel = rel.dataBR || rel.DataCelula || rel.data || grupo.semanaData;

                          return (
                            <div 
                              key={`rel-${idRel}-${rIdx}`}
                              className="bg-[#1c2030] p-3 rounded-lg border border-slate-700/60 flex flex-wrap items-center justify-between gap-3 text-xs"
                            >
                              <div className="flex items-center gap-2.5">
                                <div className="w-6 h-6 rounded-full bg-slate-700/70 text-slate-300 flex items-center justify-center font-bold text-[10px]">
                                  {rIdx + 1}º
                                </div>
                                <div>
                                  <div className="flex items-center gap-2">
                                    <span className="font-semibold text-slate-200">ID: {String(idRel)}</span>
                                    <span className="text-[11px] text-slate-400 font-mono">Data: {dataRel}</span>
                                  </div>
                                  <p className="text-[11px] text-slate-400 mt-0.5">
                                    Líder informado: <span className="text-slate-300">{lider}</span>
                                  </p>
                                </div>
                              </div>

                              <div className="flex items-center gap-4">
                                <div className="text-right">
                                  <div className="font-mono font-bold text-slate-200">
                                    Total: {formatBRL(total)}
                                  </div>
                                  <div className="text-[10px] text-slate-400 font-mono">
                                    PIX: {formatBRL(pix)} • Esp: {formatBRL(esp)}
                                  </div>
                                </div>

                                <div className="pl-2 border-l border-slate-700">
                                  {isVal ? (
                                    <span className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-400 bg-emerald-500/10 px-2 py-1 rounded border border-emerald-500/20">
                                      <CheckCircle className="w-3 h-3" /> Validado
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 text-[10px] font-bold text-amber-400 bg-amber-500/10 px-2 py-1 rounded border border-amber-500/20">
                                      <Clock className="w-3 h-3" /> Pendente
                                    </span>
                                  )}
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* Rodapé do Modal */}
            <div className="p-4 border-t border-slate-700/80 bg-[#202538] flex justify-end">
              <button
                id="btn-fechar-rodape-modal-duplicados"
                onClick={() => setModalDuplicadosAberto(false)}
                className="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white text-xs font-semibold rounded-lg transition-colors"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
