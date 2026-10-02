import React, { useState, useMemo, useEffect } from 'react';
import { RotateCw, ShieldCheck, Lock, Menu } from 'lucide-react';
import { LancamentoTesouraria, ViewMode } from '../../types';
import { SharePointService } from '../../services/sharepointService';

interface DashboardViewProps {
  lancamentos: LancamentoTesouraria[];
  anoSelecionado: number | string;
  onSelectAno: (ano: number | string) => void;
  onRefresh: () => void;
  usuarioConectado?: { nome: string; email: string } | null;
  onSelectView?: (view: ViewMode) => void;
  onToggleMobileMenu?: () => void;
}

const NOMES_MESES = [
  'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
  'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
];

const NOMES_MESES_ABREV = [
  'JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN',
  'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'
];

interface CelulaSharePointItem {
  Id?: number;
  ID?: number;
  Title?: string;
  Celula?: string;
  Setor?: string;
  Status?: string;
  Created?: string;
  Criado?: string;
  [key: string]: any;
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  lancamentos,
  anoSelecionado,
  onSelectAno,
  onRefresh,
  usuarioConectado,
  onSelectView,
  onToggleMobileMenu
}) => {
  const spService = useMemo(() => SharePointService.getInstance(), []);

  const [mesSelecionado, setMesSelecionado] = useState(() => {
    const hoje = new Date();
    const mesAtualIndex = hoje.getMonth();
    return NOMES_MESES[mesAtualIndex] || 'SETEMBRO';
  });
  // Inicializa já com células do cache local para evitar flash de tela e duplo carregamento
  const [bdCelulas, setBdCelulas] = useState<CelulaSharePointItem[]>(() => {
    const cached = spService.getCelulas();
    return Array.isArray(cached) ? cached : [];
  });
  const scrollContainerRef1 = React.useRef<HTMLDivElement>(null);
  const scrollContainerRef2 = React.useRef<HTMLDivElement>(null);

  const isTodosAnos = String(anoSelecionado).toLowerCase() === 'todos' || String(anoSelecionado).toLowerCase() === 'todos os anos' || Number(anoSelecionado) === 0;
  const isTodosMeses = String(mesSelecionado).toUpperCase() === 'TODOS' || String(mesSelecionado).toUpperCase() === 'TODOS OS MESES';
  const isModoAnosNoGrafico = isTodosAnos && !isTodosMeses;

  // Anos disponíveis extraídos dos lançamentos no banco de dados (do mais recente ao mais antigo)
  const anosDisponiveis = useMemo(() => {
    const anosSet = new Set<number>();
    const anoAtualReal = new Date().getFullYear();
    anosSet.add(anoAtualReal);

    lancamentos.forEach(l => {
      if (l.ano && typeof l.ano === 'number' && l.ano > 2000) {
        anosSet.add(l.ano);
      } else if (l.dataBR && l.dataBR.includes('/')) {
        const parts = l.dataBR.split('/');
        const a = parseInt(parts[2], 10);
        if (a && a > 2000) anosSet.add(a);
      } else if (l.data && l.data.includes('-')) {
        const a = parseInt(l.data.split('-')[0], 10);
        if (a && a > 2000) anosSet.add(a);
      }
    });

    return Array.from(anosSet).sort((a, b) => b - a);
  }, [lancamentos]);

  // Meses disponíveis para o ano selecionado: se for o ano atual, exibe somente até o mês atual + opção TODOS OS MESES
  const mesesDisponiveis = useMemo(() => {
    const hoje = new Date();
    const anoAtualReal = hoje.getFullYear();
    const mesAtualIndex = hoje.getMonth(); // 0 = Jan, 9 = Out (Outubro)

    let baseMeses = NOMES_MESES;
    if (!isTodosAnos && Number(anoSelecionado) === anoAtualReal) {
      baseMeses = NOMES_MESES.slice(0, mesAtualIndex + 1);
    }
    return ['TODOS OS MESES', ...baseMeses];
  }, [anoSelecionado, isTodosAnos]);

  // Se o mês selecionado não estiver na lista de meses permitidos para o ano, ajusta para o mais recente disponível
  useEffect(() => {
    if (mesesDisponiveis.length > 0 && !mesesDisponiveis.includes(mesSelecionado)) {
      setMesSelecionado(mesesDisponiveis[mesesDisponiveis.length - 1]);
    }
  }, [mesesDisponiveis, mesSelecionado]);

  // Carrega células da lista BD_celulas do SharePoint se ainda não estiverem carregadas
  useEffect(() => {
    if (bdCelulas.length > 0) return;
    let isMounted = true;
    fetch('/api/sharepoint/celulas')
      .then(res => res.json())
      .then(data => {
        if (isMounted && data?.celulas && Array.isArray(data.celulas)) {
          setBdCelulas(data.celulas);
        }
      })
      .catch(err => {
        console.warn('Erro ao carregar BD_celulas no Dashboard:', err);
      });

    return () => {
      isMounted = false;
    };
  }, [bdCelulas.length]);

  // Auto-scroll inicial para deixar o mês atual/selecionado à mostra nas colunas dos gráficos
  useEffect(() => {
    if (isModoAnosNoGrafico) return; // Não faz scroll quando exibe anos lado a lado

    const scrollToActiveMonth = () => {
      let mesAbrevAlvo = 'SET';
      if (!isTodosMeses) {
        const idxMes = NOMES_MESES.indexOf(mesSelecionado.toUpperCase());
        if (idxMes !== -1) mesAbrevAlvo = NOMES_MESES_ABREV[idxMes];
      } else {
        const hoje = new Date();
        mesAbrevAlvo = NOMES_MESES_ABREV[hoje.getMonth()] || 'SET';
      }

      const containers = [scrollContainerRef1.current, scrollContainerRef2.current];
      containers.forEach(container => {
        if (!container) return;
        const target = container.querySelector(`[data-mes="${mesAbrevAlvo}"]`) as HTMLElement;
        if (target) {
          const scrollPos = target.offsetLeft - (container.clientWidth / 2) + (target.clientWidth / 2);
          container.scrollTo({ left: Math.max(0, scrollPos), behavior: 'smooth' });
        } else {
          container.scrollTo({ left: container.scrollWidth, behavior: 'smooth' });
        }
      });
    };

    const timer = setTimeout(scrollToActiveMonth, 120);
    return () => clearTimeout(timer);
  }, [mesSelecionado, anoSelecionado, isTodosMeses, isModoAnosNoGrafico]);

  const numMesSelecionado = useMemo(() => {
    if (isTodosMeses) return 0;
    const idx = NOMES_MESES.indexOf(mesSelecionado.toUpperCase());
    return idx !== -1 ? idx + 1 : 0;
  }, [mesSelecionado, isTodosMeses]);

  // REGRA PRINCIPAL: Total a ser mostrado sendo SOMENTE os validados pela coluna TESOURARIA_RECEB
  const lancamentosValidados = useMemo(() => {
    return lancamentos.filter(l => l.TESOURARIA_RECEB === true);
  }, [lancamentos]);

  // Lançamentos validados do ano e mês selecionado (suporta TODOS OS ANOS e TODOS OS MESES)
  const lancamentosMesAtual = useMemo(() => {
    return lancamentosValidados.filter(l => {
      const itemAno = l.ano !== undefined ? Number(l.ano) : (l.data ? new Date(l.data).getFullYear() : 0);
      const itemMes = l.mes !== undefined ? Number(l.mes) : (l.data ? new Date(l.data).getMonth() + 1 : 0);
      
      const matchAno = isTodosAnos || itemAno === Number(anoSelecionado);
      const matchMes = isTodosMeses || itemMes === Number(numMesSelecionado);
      return matchAno && matchMes;
    });
  }, [lancamentosValidados, isTodosAnos, anoSelecionado, isTodosMeses, numMesSelecionado]);

  // Todos os relatórios lançados (registrados) do ano e mês selecionado
  const lancamentosTodosMesAtual = useMemo(() => {
    return lancamentos.filter(l => {
      const itemAno = l.ano !== undefined ? Number(l.ano) : (l.data ? new Date(l.data).getFullYear() : 0);
      const itemMes = l.mes !== undefined ? Number(l.mes) : (l.data ? new Date(l.data).getMonth() + 1 : 0);
      
      const matchAno = isTodosAnos || itemAno === Number(anoSelecionado);
      const matchMes = isTodosMeses || itemMes === Number(numMesSelecionado);
      return matchAno && matchMes;
    });
  }, [lancamentos, isTodosAnos, anoSelecionado, isTodosMeses, numMesSelecionado]);

  // Totais do Mês Selecionado (somente TESOURARIA_RECEB === true)
  const totalMesPix = useMemo(() => {
    return lancamentosMesAtual.reduce((acc, curr) => acc + (curr.ValorOferta ?? curr.valorPix ?? 0), 0);
  }, [lancamentosMesAtual]);

  const totalMesEspecie = useMemo(() => {
    return lancamentosMesAtual.reduce((acc, curr) => acc + (curr.OfertaEspecie ?? curr.valorEspecie ?? 0), 0);
  }, [lancamentosMesAtual]);

  const totalMesGeral = totalMesPix + totalMesEspecie;

  // Células Ativas: Quantidade de células da BD_celulas que foram criadas até o mês e ano selecionado
  const celulasAtivas = useMemo(() => {
    const anoAlvo = isTodosAnos ? new Date().getFullYear() : Number(anoSelecionado);
    const mesAlvo = isTodosMeses ? 12 : numMesSelecionado;
    const dataLimite = new Date(anoAlvo, mesAlvo, 0, 23, 59, 59, 999);

    if (bdCelulas.length > 0) {
      const filtradas = bdCelulas.filter(c => {
        const dtStr = c.Criado || c.Created;
        if (!dtStr) return true; // Se não tiver data de criação registrada, considera ativa
        const dtCriado = new Date(dtStr);
        if (isNaN(dtCriado.getTime())) return true;
        return dtCriado <= dataLimite;
      });

      return filtradas.length;
    }

    // Fallback caso a lista da API ainda não tenha retornado: busca células dos relatórios
    const setCelulas = new Set<string>();
    lancamentos.forEach(l => {
      const nome = l.Célula || l.celulaNome;
      if (nome && (isTodosAnos || !l.ano || l.ano <= anoAlvo)) {
        setCelulas.add(nome);
      }
    });
    return setCelulas.size > 0 ? setCelulas.size : 67;
  }, [bdCelulas, isTodosAnos, anoSelecionado, isTodosMeses, numMesSelecionado, lancamentos]);

  // Relatórios Previstos = Quantidade de Células Ativas * 4 (se todos os meses, multiplica pelos meses)
  const fatorMeses = useMemo(() => {
    if (isTodosMeses) {
      return (isTodosAnos || Number(anoSelecionado) === new Date().getFullYear() ? new Date().getMonth() + 1 : 12);
    }
    return 1;
  }, [isTodosMeses, isTodosAnos, anoSelecionado]);

  const relatoriosPrevistos = useMemo(() => {
    if (isTodosAnos && isTodosMeses) {
      const qtdAnos = Math.max(1, anosDisponiveis.length);
      return celulasAtivas * 4 * fatorMeses * qtdAnos;
    }
    if (isTodosAnos && !isTodosMeses) {
      const qtdAnos = Math.max(1, anosDisponiveis.length);
      return celulasAtivas * 4 * qtdAnos;
    }
    return celulasAtivas * 4 * fatorMeses;
  }, [isTodosAnos, isTodosMeses, celulasAtivas, fatorMeses, anosDisponiveis]);

  // Dados para o gráfico:
  // 1) Se for TODOS OS ANOS + MÊS ESPECÍFICO => exibe os ANOS no eixo X com o total daquele mês em cada ano
  // 2) Se for ANO ESPECÍFICO ou TODOS OS ANOS + TODOS OS MESES => exibe os 12 MESES no eixo X
  const mesesGrafico = useMemo(() => {
    // CASO A: TODOS OS ANOS + MÊS ESPECÍFICO
    if (isModoAnosNoGrafico) {
      const anosOrdenados = [...anosDisponiveis].sort((a, b) => a - b);

      return anosOrdenados.map(anoItem => {
        const doAnoEMes = lancamentosValidados.filter(l => {
          const itemAno = l.ano !== undefined ? Number(l.ano) : (l.data ? new Date(l.data).getFullYear() : 0);
          const itemMes = l.mes !== undefined ? Number(l.mes) : (l.data ? new Date(l.data).getMonth() + 1 : 0);
          return itemAno === anoItem && itemMes === Number(numMesSelecionado);
        });

        const esp = doAnoEMes.reduce((acc, curr) => acc + Number(curr.OfertaEspecie ?? curr.valorEspecie ?? 0), 0);
        const pix = doAnoEMes.reduce((acc, curr) => {
          const v = curr.ValorOferta ?? curr.valorPix ?? (curr.Bairro && !isNaN(Number(curr.Bairro)) ? Number(curr.Bairro) : 0);
          return acc + Number(v);
        }, 0);

        let ativas = 0;
        if (bdCelulas.length > 0) {
          const dataLimite = new Date(anoItem, numMesSelecionado, 0, 23, 59, 59, 999);
          ativas = bdCelulas.filter(c => {
            const dtStr = c.Criado || c.Created;
            if (!dtStr) return true;
            const dtCriado = new Date(dtStr);
            if (isNaN(dtCriado.getTime())) return true;
            return dtCriado <= dataLimite;
          }).length;
        } else {
          ativas = celulasAtivas;
        }

        const previstos = ativas * 4;
        const confirmadosTesouraria = doAnoEMes.length;

        let perc = 0;
        if (previstos > 0 && confirmadosTesouraria > 0) {
          const proporcao = confirmadosTesouraria / previstos;
          perc = proporcao >= 1 ? 100 : Math.round(proporcao * 100);
        }

        return {
          nome: String(anoItem), // Exibe o ANO embaixo da coluna
          subRotulo: mesSelecionado,
          esp,
          pix,
          total: esp + pix,
          confirmados: confirmadosTesouraria,
          previstos,
          perc
        };
      });
    }

    // CASO B: ANO ESPECÍFICO OU TODOS OS ANOS + TODOS OS MESES
    return NOMES_MESES_ABREV.map((nomeAbrev, idx) => {
      const mesNum = idx + 1;
      const doMes = lancamentosValidados.filter(l => {
        const itemAno = l.ano !== undefined ? Number(l.ano) : (l.data ? new Date(l.data).getFullYear() : 0);
        const itemMes = l.mes !== undefined ? Number(l.mes) : (l.data ? new Date(l.data).getMonth() + 1 : 0);
        const matchAno = isTodosAnos || itemAno === Number(anoSelecionado);
        return matchAno && itemMes === mesNum;
      });

      const esp = doMes.reduce((acc, curr) => acc + Number(curr.OfertaEspecie ?? curr.valorEspecie ?? 0), 0);
      const pix = doMes.reduce((acc, curr) => {
        const v = curr.ValorOferta ?? curr.valorPix ?? (curr.Bairro && !isNaN(Number(curr.Bairro)) ? Number(curr.Bairro) : 0);
        return acc + Number(v);
      }, 0);
      
      let ativasMes = 0;
      const anoBase = isTodosAnos ? new Date().getFullYear() : Number(anoSelecionado);
      if (bdCelulas.length > 0) {
        const dataLimiteMes = new Date(anoBase, mesNum, 0, 23, 59, 59, 999);
        ativasMes = bdCelulas.filter(c => {
          const dtStr = c.Criado || c.Created;
          if (!dtStr) return true;
          const dtCriado = new Date(dtStr);
          if (isNaN(dtCriado.getTime())) return true;
          return dtCriado <= dataLimiteMes;
        }).length;
      } else {
        ativasMes = celulasAtivas;
      }

      const multAnos = isTodosAnos ? Math.max(1, anosDisponiveis.length) : 1;
      const previstos = ativasMes * 4 * multAnos;
      const confirmadosTesouraria = doMes.length;

      let perc = 0;
      if (previstos > 0 && confirmadosTesouraria > 0) {
        const proporcao = confirmadosTesouraria / previstos;
        perc = proporcao >= 1 ? 100 : Math.round(proporcao * 100);
      }

      return {
        nome: nomeAbrev, // Exibe o MÊS embaixo da coluna
        subRotulo: isTodosAnos ? 'Todos os Anos' : String(anoSelecionado),
        esp,
        pix,
        total: esp + pix,
        confirmados: confirmadosTesouraria,
        previstos,
        perc
      };
    });
  }, [isModoAnosNoGrafico, anosDisponiveis, lancamentosValidados, numMesSelecionado, bdCelulas, celulasAtivas, isTodosAnos, anoSelecionado, mesSelecionado]);

  const maxOferta = useMemo(() => {
    const maxVal = Math.max(...mesesGrafico.map(m => Math.max(m.esp, m.pix)), 0);
    return maxVal > 0 ? maxVal * 1.3 : 1000;
  }, [mesesGrafico]);

  // Formatar valores para exibição acima das colunas com 2 casas decimais (Ex.: 8.000,00)
  const formatarValorColuna = (val: number, temMovimento: boolean = false): string => {
    if (val === undefined || val === null) return '';
    if (val <= 0) {
      return temMovimento ? '0,00' : '';
    }
    return val.toLocaleString('pt-BR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  };

  // Helper para as cores do percentual conforme regra do usuário:
  // Acima de 85% - Verde
  // Abaixo de 85% e Acima de 65% - Amarelo
  // Abaixo 65% - Vermelho
  const getCorPercentual = (perc: number) => {
    if (perc > 85) {
      return {
        text: 'text-emerald-400',
        bar: 'bg-emerald-500',
        badge: 'bg-emerald-500/15 text-emerald-300 border-emerald-500/30'
      };
    }
    if (perc >= 65) {
      return {
        text: 'text-amber-400',
        bar: 'bg-amber-500',
        badge: 'bg-amber-500/15 text-amber-300 border-amber-500/30'
      };
    }
    return {
      text: 'text-rose-400',
      bar: 'bg-rose-500',
      badge: 'bg-rose-500/15 text-rose-300 border-rose-500/30'
    };
  };

  // Ranking por setores baseado em:
  // (Qtd Célula Ativas do Setor * 4 / Qtd Relatórios Entregues a Tesouraria do Setor)*100
  // Sendo o teto de 100% quando todos os relatórios daquele mês forem entregues.
  // Ajusta-se dinamicamente conforme o mês e ano selecionados na tela.
  const rankingSetores = useMemo(() => {
    // 1. Células ativas do setor até o mês e ano selecionados
    const anoAlvo = isTodosAnos ? new Date().getFullYear() : Number(anoSelecionado);
    const mesAlvo = isTodosMeses ? 12 : numMesSelecionado;
    const dataLimiteMes = new Date(anoAlvo, mesAlvo, 0, 23, 59, 59, 999);
    const celulasAtivasPorSetor = new Map<string, number>();

    if (bdCelulas.length > 0) {
      bdCelulas.forEach(c => {
        const dtStr = c.Criado || c.Created;
        let ativa = true;
        if (dtStr) {
          const dtCriado = new Date(dtStr);
          if (!isNaN(dtCriado.getTime()) && dtCriado > dataLimiteMes) {
            ativa = false;
          }
        }
        if (c.Status && c.Status === 'Inativa') ativa = false;

        if (ativa) {
          const setor = (c.Setor || c.setor || 'Sem Setor').trim();
          celulasAtivasPorSetor.set(setor, (celulasAtivasPorSetor.get(setor) || 0) + 1);
        }
      });
    } else {
      // Fallback: agrupa células a partir dos lançamentos
      const celulasSet = new Map<string, Set<string>>();
      lancamentos.forEach(l => {
        const nome = l.Célula || l.celulaNome;
        const setor = (l.Setor || l.setor || 'Sem Setor').trim();
        if (nome && (isTodosAnos || !l.ano || l.ano <= anoAlvo)) {
          if (!celulasSet.has(setor)) celulasSet.set(setor, new Set());
          celulasSet.get(setor)!.add(nome);
        }
      });
      celulasSet.forEach((set, setor) => {
        celulasAtivasPorSetor.set(setor, set.size);
      });
    }

    // 2. Relatórios entregues à tesouraria do setor no mês e ano selecionado (TESOURARIA_RECEB === true)
    const entreguesPorSetor = new Map<string, number>();
    lancamentos.forEach(l => {
      if (l.TESOURARIA_RECEB !== true) return;

      const itemAno = l.ano !== undefined ? Number(l.ano) : (l.data ? new Date(l.data).getFullYear() : 0);
      const itemMes = l.mes !== undefined ? Number(l.mes) : (l.data ? new Date(l.data).getMonth() + 1 : 0);

      const matchAno = isTodosAnos || itemAno === Number(anoSelecionado);
      const matchMes = isTodosMeses || itemMes === Number(numMesSelecionado);

      if (matchAno && matchMes) {
        const setor = (l.Setor || l.setor || 'Sem Setor').trim();
        entreguesPorSetor.set(setor, (entreguesPorSetor.get(setor) || 0) + 1);
      }
    });

    // 3. Consolidar todos os setores
    const todosSetores = new Set<string>([
      ...Array.from(celulasAtivasPorSetor.keys()),
      ...Array.from(entreguesPorSetor.keys())
    ]);

    if (todosSetores.size === 0) {
      ['Safira', 'Fire', 'White', 'Black', 'Azul', 'Amarelo', 'Legacy', 'Onix', 'Diamante', 'Titanium'].forEach(s => todosSetores.add(s));
    }

    const lista = Array.from(todosSetores)
      .filter(nome => Boolean(nome) && (nome !== 'Sem Setor' || (celulasAtivasPorSetor.get(nome) || 0) > 0 || (entreguesPorSetor.get(nome) || 0) > 0))
      .map(nome => {
        const ativas = celulasAtivasPorSetor.get(nome) || 0;
        const previstos = ativas * 4 * fatorMeses;
        const entregues = entreguesPorSetor.get(nome) || 0;

        let perc = 0;
        if (previstos > 0 && entregues > 0) {
          const proporcao = entregues / previstos;
          perc = proporcao >= 1 ? 100 : Math.round(proporcao * 100);
        } else if (previstos === 0 && entregues > 0) {
          perc = 100;
        }

        return {
          nome,
          ativas,
          previstos,
          entregues,
          perc
        };
      });

    lista.sort((a, b) => b.perc - a.perc || b.entregues - a.entregues);
    return lista;
  }, [lancamentos, bdCelulas, isTodosAnos, anoSelecionado, isTodosMeses, numMesSelecionado, fatorMeses]);

  // Bloco Relatório Lançados x Recebidos (com base no filtro de mês e ano selecionado)
  // Relatórios Previstos = Nº de Células Ativas * 4
  const relatoriosPrevistosMes = relatoriosPrevistos;
  const totalLancadosMes = lancamentosTodosMesAtual.length;
  const totalValidadosMes = lancamentosMesAtual.length;

  // Cálculo Relatórios Previstos: (Nº Relatórios Lançados nesse mês e ano selecionado / Nº Relatórios Previstos [Células Ativas * 4]) * 100, topo 100%
  const percPrevistosMes = relatoriosPrevistosMes > 0
    ? Math.min(100, Math.round((totalLancadosMes / relatoriosPrevistosMes) * 100))
    : 0;

  // Cálculo Relatórios Validados: (Nº de Relatórios Validados do mês e ano selecionado / Nº Relatórios Previstos para o mês e ano selecionado) * 100, topo 100%
  const percValidadosMes = relatoriosPrevistosMes > 0 
    ? Math.min(100, Math.round((totalValidadosMes / relatoriosPrevistosMes) * 100)) 
    : 0;

  // Cores dinâmicas para as barras horizontais seguindo a mesma regra do Ranking por Setor:
  // Acima de 85%: Verde | Entre 65% e 85%: Amarelo | Abaixo de 65%: Vermelho
  const corPrevistosMes = getCorPercentual(percPrevistosMes);
  const corValidadosMes = getCorPercentual(percValidadosMes);

  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(val || 0);
  };

  return (
    <div id="dashboard-view-container" className="p-3.5 sm:p-6 space-y-4 sm:space-y-6 max-w-[1600px] mx-auto text-slate-100">
      {/* Top Header */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#2d334d] pb-4">
        <div className="flex items-center gap-2.5">
          {onToggleMobileMenu && (
            <button
              onClick={onToggleMobileMenu}
              className="md:hidden p-2 -ml-1 rounded-lg text-slate-300 hover:text-white hover:bg-[#282e48] transition-colors cursor-pointer"
              aria-label="Abrir Menu"
              title="Menu"
            >
              <Menu className="w-5 h-5" />
            </button>
          )}

          <div>
            <p className="text-[11px] sm:text-xs font-medium text-slate-400">
              Olá {usuarioConectado?.nome || 'Admin'}
            </p>
            <h2 className="text-lg sm:text-2xl font-black text-white tracking-tight">
              DashBoard | Ofertas Célula
            </h2>
          </div>
        </div>

        <div className="flex items-center flex-wrap gap-2 sm:gap-3 ml-auto">
          {/* SharePoint Login status badge */}
          {usuarioConectado ? (
            <div 
              className="hidden sm:flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[#22273c] border border-[#343b57] text-xs text-slate-300"
              title="Conta Microsoft / SharePoint conectada"
            >
              <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0" />
              <div className="text-left">
                <span className="block text-[11px] font-bold text-white leading-tight">SharePoint Conectado</span>
                <span className="block text-[10px] text-slate-400 leading-tight truncate max-w-[120px]">{usuarioConectado.nome}</span>
              </div>
            </div>
          ) : (
            <button
              id="dashboard-btn-login-sp"
              onClick={() => onSelectView?.('login')}
              className="flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-sm transition-colors cursor-pointer"
            >
              <Lock className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Login SharePoint</span>
              <span className="sm:hidden">Login</span>
            </button>
          )}

          <div className="flex items-center gap-1.5 bg-[#22273c] px-2.5 py-1.5 rounded-lg border border-[#343b57]">
            <span className="text-[11px] sm:text-xs text-slate-400">Ano:</span>
            <select
              value={isTodosAnos ? 'todos' : anoSelecionado}
              onChange={(e) => onSelectAno(e.target.value === 'todos' ? 'todos' : Number(e.target.value))}
              className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
            >
              <option value="todos" className="bg-[#1c2030] text-white">
                TODOS OS ANOS
              </option>
              {anosDisponiveis.map(ano => (
                <option key={ano} value={ano} className="bg-[#1c2030] text-white">
                  {ano}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-1.5 bg-[#22273c] px-2.5 py-1.5 rounded-lg border border-[#343b57]">
            <select
              value={mesSelecionado}
              onChange={(e) => setMesSelecionado(e.target.value)}
              className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
            >
              {mesesDisponiveis.map(m => (
                <option key={m} value={m} className="bg-[#1c2030] text-white">
                  {m}
                </option>
              ))}
            </select>
          </div>

          <button
            onClick={onRefresh}
            className="p-2 rounded-full hover:bg-[#282d46] text-slate-300 hover:text-white transition-colors cursor-pointer"
            title="Atualizar Dashboard"
          >
            <RotateCw className="w-4 sm:w-5 h-4 sm:h-5" />
          </button>
        </div>
      </div>

      {/* 5 KPI Cards com dados do SharePoint do Mês e Ano Selecionados */}
      <div className="grid grid-cols-2 md:grid-cols-2 lg:grid-cols-5 gap-2.5 sm:gap-3.5">
        {/* Total Mês PIX */}
        <div className="bg-[#24293f] p-3 sm:p-4 rounded-xl border border-[#323955]">
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium mb-1 truncate">
            {isModoAnosNoGrafico ? `Total ${mesSelecionado} (PIX)` : isTodosMeses ? 'Total Período (PIX)' : 'Total Mês PIX'}
          </p>
          <p className="text-base sm:text-xl lg:text-2xl font-extrabold text-white tracking-tight truncate">
            {formatBRL(totalMesPix)}
          </p>
        </div>

        {/* Total Mês Espécie */}
        <div className="bg-[#24293f] p-3 sm:p-4 rounded-xl border border-[#323955]">
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium mb-1 truncate">
            {isModoAnosNoGrafico ? `Total ${mesSelecionado} (Espécie)` : isTodosMeses ? 'Total Período (Espécie)' : 'Total Oferta Espécie'}
          </p>
          <p className="text-base sm:text-xl lg:text-2xl font-extrabold text-white tracking-tight truncate">
            {formatBRL(totalMesEspecie)}
          </p>
        </div>

        {/* Total Validado */}
        <div className="bg-[#24293f] p-3 sm:p-4 rounded-xl border border-[#323955] col-span-2 lg:col-span-1">
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium mb-1 truncate">
            {isModoAnosNoGrafico ? `Total Validado (${mesSelecionado})` : isTodosMeses ? 'Total Validado Geral' : 'Total Validado'}
          </p>
          <p className="text-base sm:text-xl lg:text-2xl font-extrabold text-emerald-400 tracking-tight truncate">
            {formatBRL(totalMesGeral)}
          </p>
        </div>

        {/* Células Ativas */}
        <div className="bg-[#24293f] p-3 sm:p-4 rounded-xl border border-[#323955] text-center">
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium mb-1 truncate">Células Ativas</p>
          <p className="text-2xl sm:text-3xl font-black text-white tracking-tight">
            {celulasAtivas}
          </p>
        </div>

        {/* Relatórios Validados / Previstos */}
        <div className="bg-[#24293f] p-3 sm:p-4 rounded-xl border border-[#323955] text-center flex flex-col justify-center">
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium mb-1 truncate" title="Relatórios Validados / Previstos">
            Relatórios Validados / Previstos
          </p>
          <p className="text-xl sm:text-2xl lg:text-3xl font-black text-white tracking-tight flex items-center justify-center gap-1.5">
            <span className="text-emerald-400">{totalValidadosMes}</span>
            <span className="text-slate-500 font-normal">/</span>
            <span>{relatoriosPrevistos}</span>
          </p>
        </div>
      </div>

      {/* Grade Principal: Coluna Esquerda (Ofertas + Relatórios Recebidos) com ~67% e Coluna Direita (Lançados x Recebidos + Ranking) reduzida em ~40% com ~33% */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-5">
        {/* Coluna Esquerda: Ofertas por Mês R$ e abaixo Relatórios Recebidos Mês a Mês (%) */}
        <div className="lg:col-span-8 flex flex-col gap-4 sm:gap-5">
          {/* Ofertas por Mês R$ */}
          <div className="bg-[#24293f] p-3.5 sm:p-4 rounded-xl border border-[#323955] flex flex-col justify-between">
            <div className="flex items-center justify-between mb-2 sm:mb-3">
              <div>
                <h3 className="text-sm font-bold text-white">
                  {isModoAnosNoGrafico
                    ? `Ofertas de ${mesSelecionado} por Ano R$`
                    : isTodosAnos && isTodosMeses
                    ? 'Ofertas por Mês R$ (Consolidado Todos os Anos)'
                    : 'Ofertas por Mês R$'}
                </h3>
              </div>
              <div className="flex items-center gap-3 text-xs">
                <span className="flex items-center gap-1.5 text-slate-300">
                  <span className="w-3 h-3 rounded-xs bg-[#22c55e]"></span>
                  PIX
                </span>
                <span className="flex items-center gap-1.5 text-slate-300">
                  <span className="w-3 h-3 rounded-xs bg-[#cbd5e1]"></span>
                  Espécie
                </span>
              </div>
            </div>

            {/* Bar Chart Container */}
            <div ref={scrollContainerRef1} className={`overflow-x-auto pb-1 scrollbar-thin ${isModoAnosNoGrafico ? 'w-full' : ''}`}>
              <div className={`h-52 sm:h-56 items-end pt-6 pb-2 px-3 sm:px-4 border-b border-[#303752] ${
                isModoAnosNoGrafico
                  ? 'w-full flex justify-around gap-2 sm:gap-6 min-w-[320px]'
                  : 'min-w-[760px] xl:min-w-0 grid grid-cols-12 gap-1.5 sm:gap-2'
              }`}>
                {mesesGrafico.map((m) => {
                  const altPix = maxOferta > 0 ? (m.pix / maxOferta) * 100 : 0;
                  const altEsp = maxOferta > 0 ? (m.esp / maxOferta) * 100 : 0;
                  const temMovimento = m.total > 0 || m.confirmados > 0;
                  const valPixStr = formatarValorColuna(m.pix, temMovimento);
                  const valEspStr = formatarValorColuna(m.esp, temMovimento);

                  return (
                    <div
                      key={m.nome}
                      data-mes={m.nome}
                      className="flex flex-col items-center h-full justify-end group relative"
                      title={`${m.nome} - Total: ${formatBRL(m.total)} (PIX: ${formatBRL(m.pix)} | Espécie: ${formatBRL(m.esp)})`}
                    >
                      {/* Barras Lado a Lado com valores individuais no formato com 2 decimais (Ex.: 8.000,00) */}
                      <div className="w-full h-36 sm:h-40 flex items-end justify-center gap-1 px-0.5">
                        {/* Coluna Verde - PIX */}
                        <div className="flex flex-col items-center justify-end h-full min-w-0">
                          {valPixStr && (
                            <span
                              className="text-[7.5px] sm:text-[9px] font-bold text-[#22c55e] mb-1 leading-none text-center whitespace-nowrap select-none tracking-tight"
                              title={`PIX: R$ ${valPixStr}`}
                            >
                              {valPixStr}
                            </span>
                          )}
                          <div
                            style={{ height: `${m.pix > 0 ? Math.max(6, Math.round(altPix)) : (temMovimento ? 2 : 0)}%` }}
                            className={`w-3.5 sm:w-4.5 transition-all duration-300 ${
                              m.pix > 0 ? 'bg-[#22c55e] hover:bg-[#16a34a]' : temMovimento ? 'bg-[#22c55e]/30' : 'bg-transparent'
                            } rounded-t-xs`}
                            title={`PIX: ${formatBRL(m.pix)}`}
                          />
                        </div>

                        {/* Coluna Cinza - Espécie */}
                        <div className="flex flex-col items-center justify-end h-full min-w-0">
                          {valEspStr && (
                            <span
                              className="text-[7.5px] sm:text-[9px] font-bold text-slate-300 mb-1 leading-none text-center whitespace-nowrap select-none tracking-tight"
                              title={`Espécie: R$ ${valEspStr}`}
                            >
                              {valEspStr}
                            </span>
                          )}
                          <div
                            style={{ height: `${m.esp > 0 ? Math.max(6, Math.round(altEsp)) : (temMovimento ? 2 : 0)}%` }}
                            className={`w-3.5 sm:w-4.5 transition-all duration-300 ${
                              m.esp > 0 ? 'bg-[#cbd5e1] hover:bg-white' : temMovimento ? 'bg-[#cbd5e1]/30' : 'bg-transparent'
                            } rounded-t-xs`}
                            title={`Espécie: ${formatBRL(m.esp)}`}
                          />
                        </div>
                      </div>

                      <span className="text-[10px] font-bold text-slate-400 mt-1.5">
                        {m.nome}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Relatórios Recebidos Mês a Mês (%) */}
          <div className="bg-[#24293f] p-4 sm:p-5 rounded-xl border border-[#323955] flex flex-col justify-between">
            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="text-sm font-bold text-white">
                  {isModoAnosNoGrafico
                    ? `Relatórios Recebidos de ${mesSelecionado} por Ano (%)`
                    : 'Relatórios Recebidos Mês a Mês (%)'}
                </h3>
                <p className="text-[11px] text-slate-400 mt-0.5">
                  {isModoAnosNoGrafico
                    ? `% de relatórios validados pela Tesouraria no mês de ${mesSelecionado} em relação à meta prevista`
                    : '% de relatórios validados pela Tesouraria em relação à meta prevista'}
                </p>
              </div>
            </div>

            <div ref={scrollContainerRef2} className={`overflow-x-auto pb-1 scrollbar-thin ${isModoAnosNoGrafico ? 'w-full' : ''}`}>
              <div className={`h-56 sm:h-60 items-end pt-4 pb-2 px-3 sm:px-4 border-b border-[#303752] ${
                isModoAnosNoGrafico
                  ? 'w-full flex justify-around gap-2 sm:gap-6 min-w-[320px]'
                  : 'min-w-[760px] xl:min-w-0 grid grid-cols-12 gap-1.5 sm:gap-2'
              }`}>
                {mesesGrafico.map((m) => {
                  const isRed = m.perc < 80 && m.perc > 0;
                  const barColor = isRed ? 'bg-[#c85a5a]' : 'bg-[#22c55e]';

                  return (
                    <div 
                      key={m.nome} 
                      data-mes={m.nome}
                      className="flex flex-col items-center h-full justify-end" 
                      title={`${m.nome}: ${m.confirmados} validados / ${m.previstos} previstos (${m.perc}%)`}
                    >
                      <div className="h-5 flex items-center justify-center mb-1">
                        {m.perc > 0 && (
                          <span className="text-[9px] text-slate-300 font-bold">
                            {m.perc}%
                          </span>
                        )}
                      </div>
                      <div className="w-full h-36 sm:h-40 flex items-end justify-center px-0.5">
                        <div
                          style={{ height: `${m.perc > 0 ? Math.max(5, m.perc) : 0}%` }}
                          className={`w-3.5 sm:w-4.5 transition-all duration-300 ${m.perc > 0 ? barColor : 'bg-transparent'} rounded-t-xs`}
                        />
                      </div>
                      <span className="text-[10px] font-bold text-slate-400 mt-2">
                        {m.nome}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* Coluna Direita (reduzida em ~40%): Relatórios Lançados x Recebidos e abaixo Ranking por Setores em coluna única */}
        <div className="lg:col-span-4 flex flex-col gap-4 sm:gap-5">
          {/* Relatório Lançados x Recebidos */}
          <div className="bg-[#24293f] p-4 sm:p-5 rounded-xl border border-[#323955] flex flex-col justify-center space-y-5">
            <div>
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold text-white">Relatório Lançados x Recebidos</h3>
                <span className="text-[10px] text-slate-400 font-medium">
                  ({isTodosMeses ? 'Todos os Meses' : mesSelecionado}/{isTodosAnos ? 'Todos os Anos' : anoSelecionado})
                </span>
              </div>
              <p className="text-[11px] text-slate-400 mt-1 leading-snug">
                Percentual de envelopes em relação à meta prevista por célula
              </p>
            </div>

            {/* Relat. Lançados */}
            <div>
              <div className="flex justify-between items-baseline text-xs font-semibold text-slate-300 mb-1">
                <span>Relat. Lançados</span>
                <span className={`font-bold ${corPrevistosMes.text}`}>{percPrevistosMes}%</span>
              </div>
              <p className="text-[10px] text-slate-400 mb-1.5">
                Lançados pelas células ({totalLancadosMes} de {relatoriosPrevistosMes} previstos)
              </p>
              <div className="w-full bg-[#181b2a] rounded-sm h-5 overflow-hidden p-0.5 border border-[#303752]">
                <div 
                  className={`${corPrevistosMes.bar} h-full rounded-xs transition-all duration-500`} 
                  style={{ width: `${Math.min(100, Math.max(0, percPrevistosMes))}%` }} 
                />
              </div>
            </div>

            {/* Relatórios Validados */}
            <div>
              <div className="flex justify-between items-baseline text-xs font-semibold text-slate-300 mb-1">
                <span>Relatórios Validados</span>
                <span className={`font-bold ${corValidadosMes.text}`}>{percValidadosMes}%</span>
              </div>
              <p className="text-[10px] text-slate-400 mb-1.5">
                Confirmados pela Tesouraria ({totalValidadosMes} de {relatoriosPrevistosMes} previstos)
              </p>
              <div className="w-full bg-[#181b2a] rounded-sm h-5 overflow-hidden p-0.5 border border-[#303752]">
                <div 
                  className={`${corValidadosMes.bar} h-full rounded-xs transition-all duration-500`} 
                  style={{ width: `${Math.min(100, Math.max(0, percValidadosMes))}%` }} 
                />
              </div>
            </div>
          </div>

          {/* Ranking por Setores */}
          <div id="card-ranking-setores" className="bg-[#24293f] p-4 sm:p-5 rounded-xl border border-[#323955] flex-1 flex flex-col">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
              <div>
                <h3 className="text-sm font-bold text-white">Ranking por Setores</h3>
                <p className="text-[10px] text-slate-400 mt-0.5">
                  ({mesSelecionado}/{anoSelecionado})
                </p>
              </div>
              {/* Legenda de Cores do Percentual */}
              <div className="flex items-center gap-2 text-[10px] font-semibold bg-[#1a1e30] px-2 py-1 rounded-lg border border-[#303752] self-start sm:self-auto shrink-0">
                <span className="flex items-center gap-1 text-emerald-400" title="Acima de 85% - Verde">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0" />
                  &gt;85%
                </span>
                <span className="flex items-center gap-1 text-amber-400" title="Abaixo de 85% e Acima de 65% - Amarelo">
                  <span className="w-1.5 h-1.5 rounded-full bg-amber-500 shrink-0" />
                  65-85%
                </span>
                <span className="flex items-center gap-1 text-rose-400" title="Abaixo de 65% - Vermelho">
                  <span className="w-1.5 h-1.5 rounded-full bg-rose-500 shrink-0" />
                  &lt;65%
                </span>
              </div>
            </div>

            {/* Listagem vertical: cada setor abaixo do outro em coluna única */}
            <div className="flex flex-col divide-y divide-[#2d334c]/50">
              {rankingSetores.map((s, idx) => {
                const cor = getCorPercentual(s.perc);
                return (
                  <div 
                    key={s.nome} 
                    className="flex items-center justify-between text-xs py-2 px-1.5 hover:bg-white/[0.02] transition-colors"
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <span className="w-4 text-xs font-bold text-slate-400 shrink-0">
                        {idx + 1}º
                      </span>
                      <span className="text-sm font-bold text-white tracking-wide truncate">{s.nome}</span>
                      <span className="text-[11px] text-slate-400 shrink-0 font-medium">
                        ({s.ativas} {s.ativas === 1 ? 'célula ativa' : 'células ativas'})
                      </span>
                    </div>
                    <span className={`font-bold text-sm shrink-0 pl-2 ${cor.text}`}>
                      {s.perc}%
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
