import React, { useState, useMemo } from 'react';
import { Users, MapPin, CheckCircle, Clock, ShieldCheck } from 'lucide-react';
import { SETORES_DISPONIVEIS, CELULAS_INICIAIS } from '../../data/mockSharePointData';
import { LancamentoTesouraria, SetorTipo } from '../../types';

interface SupervisaoViewProps {
  lancamentos: LancamentoTesouraria[];
}

export const SupervisaoView: React.FC<SupervisaoViewProps> = ({ lancamentos }) => {
  const [setorSelecionado, setSetorSelecionado] = useState<string>('todos');

  const formatBRL = (val: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(val || 0);
  };

  // Extração dinâmica de supervisores e dados por setor a partir do BD_Relatorio
  const estatisticasSetores = useMemo(() => {
    return SETORES_DISPONIVEIS.map(setor => {
      const lancsSetor = lancamentos.filter(l => (l.Setor || l.setor) === setor);
      const celulasUnicas = Array.from(new Set(lancsSetor.map(l => l.Célula || l.celulaNome).filter(Boolean)));
      const lideresSetor = Array.from(new Set(lancsSetor.map(l => l.LíderSetor).filter(Boolean)));
      const liderSetorNome = lideresSetor.length > 0 ? lideresSetor.join(' / ') : `Líder Setor ${setor}`;

      // Apenas validados por TESOURARIA_RECEB
      const validados = lancsSetor.filter(l => l.TESOURARIA_RECEB === true);
      const totalValidado = validados.reduce((acc, curr) => acc + (curr.Total ?? curr.valorTotal ?? 0), 0);
      const pendentes = lancsSetor.filter(l => l.TESOURARIA_RECEB !== true);
      const comSupervisao = lancsSetor.filter(l => l.Supervisao === true);

      const areaSetor = lancsSetor[0]?.Area || lancsSetor[0]?.area || 'Setorial';

      return {
        setor,
        supervisor: liderSetorNome,
        area: areaSetor,
        totalRelatorios: lancsSetor.length,
        qtdCelulas: celulasUnicas.length > 0 ? celulasUnicas.length : CELULAS_INICIAIS.filter(c => c.setor === setor).length,
        qtdValidados: validados.length,
        qtdPendentes: pendentes.length,
        qtdSupervisionados: comSupervisao.length,
        totalValidado
      };
    });
  }, [lancamentos]);

  const setoresFiltrados = useMemo(() => {
    if (setorSelecionado === 'todos') return estatisticasSetores;
    return estatisticasSetores.filter(s => s.setor === setorSelecionado);
  }, [estatisticasSetores, setorSelecionado]);

  // Células dinâmicas com seus líderes de célula
  const celulasTabela = useMemo(() => {
    const filtrados = setorSelecionado === 'todos'
      ? lancamentos
      : lancamentos.filter(l => (l.Setor || l.setor) === setorSelecionado);

    const celulasMap = new Map<string, {
      nome: string;
      lider: string;
      setor: string;
      area: string;
      supervisao: boolean;
      qtdRelatoriosValidados: number;
    }>();

    filtrados.forEach(l => {
      const nome = l.Célula || l.celulaNome;
      if (!nome) return;
      const prev = celulasMap.get(nome) || {
        nome,
        lider: l.LíderCelula || '-',
        setor: l.Setor || l.setor,
        area: l.Area || l.area,
        supervisao: Boolean(l.Supervisao),
        qtdRelatoriosValidados: 0
      };

      if (l.Supervisao) prev.supervisao = true;
      if (l.TESOURARIA_RECEB === true) prev.qtdRelatoriosValidados += 1;

      celulasMap.set(nome, prev);
    });

    if (celulasMap.size > 0) {
      return Array.from(celulasMap.values());
    }

    // Fallback padrão se lista do SharePoint estiver vazia
    return CELULAS_INICIAIS
      .filter(c => setorSelecionado === 'todos' || c.setor === setorSelecionado)
      .map(c => ({
        nome: c.nome,
        lider: c.lider,
        setor: c.setor,
        area: c.area,
        supervisao: false,
        qtdRelatoriosValidados: 0
      }));
  }, [lancamentos, setorSelecionado]);

  return (
    <div id="supervisao-view-container" className="p-3.5 sm:p-6 space-y-4 sm:space-y-6 max-w-[1600px] mx-auto text-slate-100">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#2d334d] pb-4">
        <div>
          <h2 className="text-xl sm:text-2xl font-black text-white tracking-tight flex items-center gap-2">
            <Users className="w-5 h-5 sm:w-6 sm:h-6 text-indigo-400 shrink-0" />
            Supervisão de Setores & Áreas
          </h2>
          <p className="text-[11px] sm:text-xs text-slate-400">
            Acompanhamento de Líderes de Setor, Líderes de Célula e validação <strong className="text-emerald-400">TESOURARIA_RECEB</strong>
          </p>
        </div>

        <div className="flex items-center gap-2 bg-[#22273c] px-2.5 py-1.5 rounded-lg border border-[#343b57]">
          <span className="text-[11px] sm:text-xs text-slate-400">Filtrar Setor:</span>
          <select
            value={setorSelecionado}
            onChange={(e) => setSetorSelecionado(e.target.value)}
            className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
          >
            <option value="todos" className="bg-[#1c2030]">Todos os Setores</option>
            {SETORES_DISPONIVEIS.map(s => (
              <option key={s} value={s} className="bg-[#1c2030]">{s}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Grid de Supervisores de Setor */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-4">
        {setoresFiltrados.map((sup) => (
          <div key={sup.setor} className="bg-[#24293f] p-3.5 sm:p-4 rounded-xl border border-[#323955] space-y-3">
            <div className="flex items-center justify-between border-b border-[#2d334c] pb-2">
              <span className="text-xs font-extrabold text-indigo-300 bg-indigo-950/80 px-2.5 py-1 rounded border border-indigo-500/30">
                Setor {sup.setor}
              </span>
              <span className="text-[11px] font-medium text-slate-400 flex items-center gap-1">
                <MapPin className="w-3.5 h-3.5 text-slate-400" />
                {sup.area}
              </span>
            </div>

            <div>
              <p className="text-xs text-slate-400">Líder do Setor:</p>
              <p className="text-sm font-bold text-white mt-0.5">{sup.supervisor}</p>
            </div>

            <div className="grid grid-cols-2 gap-2 text-xs pt-1 border-t border-[#2d334c]">
              <div className="bg-[#1a1d2d] p-2 rounded">
                <span className="text-slate-400 text-[10px] block">Validados Tesouraria</span>
                <strong className="text-emerald-400 font-bold font-mono">
                  {sup.qtdValidados} <span className="text-slate-400 text-[10px]">/ {sup.totalRelatorios}</span>
                </strong>
              </div>
              <div className="bg-[#1a1d2d] p-2 rounded">
                <span className="text-slate-400 text-[10px] block">Total Validado</span>
                <strong className="text-white font-bold font-mono text-[11px]">{formatBRL(sup.totalValidado)}</strong>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Tabela de Células e Líderes */}
      <div className="bg-[#24293f] rounded-xl border border-[#323955] overflow-hidden">
        <div className="p-3.5 sm:p-4 border-b border-[#323955] flex flex-wrap justify-between items-center gap-2">
          <h3 className="text-sm font-bold text-white">
            Células & Lideranças Registradas ({celulasTabela.length})
          </h3>
          <span className="text-[11px] sm:text-xs text-indigo-300 font-mono">
            Origem: Base Oficial
          </span>
        </div>

        <div className="overflow-x-auto scrollbar-thin">
          <table className="w-full text-xs text-left min-w-[580px]">
            <thead className="bg-[#1a1d2d] text-slate-300 uppercase text-[11px] border-b border-[#2d334c]">
              <tr>
                <th className="px-4 py-3">Célula</th>
                <th className="px-4 py-3">Líder da Célula</th>
                <th className="px-4 py-3">Setor</th>
                <th className="px-4 py-3">Área</th>
                <th className="px-4 py-3 text-center">Supervisão</th>
                <th className="px-4 py-3 text-center">Relatórios Validados</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#2a3048]">
              {celulasTabela.map((c, idx) => (
                <tr key={`${c.nome}-${idx}`} className="hover:bg-[#282d46] transition-colors">
                  <td className="px-4 py-3 font-bold text-white">{c.nome}</td>
                  <td className="px-4 py-3 text-slate-300">{c.lider}</td>
                  <td className="px-4 py-3 font-semibold text-indigo-300">{c.setor}</td>
                  <td className="px-4 py-3 text-slate-300">{c.area}</td>
                  <td className="px-4 py-3 text-center">
                    {c.supervisao ? (
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-indigo-950 text-indigo-300 border border-indigo-500/30">
                        <ShieldCheck className="w-3 h-3" />
                        Sim
                      </span>
                    ) : (
                      <span className="text-slate-500 text-[10px]">Não</span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-center">
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold font-mono bg-emerald-950 text-emerald-300 border border-emerald-500/30">
                      {c.qtdRelatoriosValidados} validados
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
