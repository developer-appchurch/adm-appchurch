import React from 'react';
import { RotateCw, Database, FileSpreadsheet, FileText, UserCheck, Menu } from 'lucide-react';
import { ViewMode } from '../types';

interface HeaderProps {
  currentView: ViewMode;
  anoSelecionado: number;
  onSelectAno: (ano: number) => void;
  onRefresh: () => void;
  isRefreshing: boolean;
  onExportExcel?: () => void;
  onExportPdf?: () => void;
  usuarioConectado?: { nome: string; email: string } | null;
  onSelectView?: (view: ViewMode) => void;
  onToggleMobileMenu?: () => void;
  // Filtros adicionais para telas específicas como relacao-envelopes
  mesSelecionado?: string;
  onSelectMes?: (mes: string) => void;
  setorSelecionado?: string;
  onSelectSetor?: (setor: string) => void;
  setoresDisponiveis?: string[];
}

const MESES_HEADER = [
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

export const Header: React.FC<HeaderProps> = ({
  currentView,
  anoSelecionado,
  onSelectAno,
  onRefresh,
  isRefreshing,
  onExportExcel,
  onExportPdf,
  usuarioConectado,
  onSelectView,
  onToggleMobileMenu,
  mesSelecionado,
  onSelectMes,
  setorSelecionado,
  onSelectSetor,
  setoresDisponiveis
}) => {
  const getTitle = () => {
    switch (currentView) {
      case 'dashboard':
        return 'DashBoard | Ofertas Célula';
      case 'relacao-envelopes':
        return 'Relação de Envelopes';
      case 'validar-relatorios':
        return 'Validar Entrega de Envelope';
      case 'fluxo-caixa':
        return 'Relatórios de Fluxo de Caixa';
      case 'indicador-trilho':
        return 'Indicador do Trilho de Capacitação';
      default:
        return 'Tesouraria AppChurch';
    }
  };

  return (
    <header 
      id="app-header" 
      className="bg-[#1c2030] text-slate-100 px-3.5 sm:px-6 py-3 sm:py-4 border-b border-[#2a2f48] flex flex-wrap items-center justify-between gap-3"
    >
      <div className="flex items-center gap-2.5">
        {/* Mobile Hamburger Button */}
        {onToggleMobileMenu && (
          <button
            onClick={onToggleMobileMenu}
            className="md:hidden p-2 -ml-1 rounded-lg text-slate-300 hover:text-white hover:bg-[#282e48] transition-colors cursor-pointer"
            aria-label="Abrir Menu Lateral"
            title="Menu"
          >
            <Menu className="w-5 h-5" />
          </button>
        )}

        <div>
          <p className="text-[11px] sm:text-xs text-slate-400 font-medium tracking-wide">
            Olá {usuarioConectado?.nome || 'Admin'}
          </p>
          <div className="flex items-center gap-2 sm:gap-3">
            <h2 className="text-base sm:text-xl font-bold tracking-tight text-white flex items-center gap-2">
              {getTitle()}
            </h2>
          </div>
        </div>
      </div>

      {/* Right Controls */}
      <div className="flex items-center flex-wrap gap-2 sm:gap-2.5 ml-auto">
        {/* Export Quick Buttons when on Cash Flow */}
        {currentView === 'fluxo-caixa' && onExportExcel && onExportPdf && (
          <div className="flex items-center gap-1 sm:gap-1.5">
            <button
              id="header-btn-excel"
              onClick={onExportExcel}
              className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold rounded bg-emerald-700/80 hover:bg-emerald-600 text-white shadow-sm border border-emerald-500/50 transition-colors cursor-pointer"
              title="Exportar para Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Excel</span>
            </button>
            <button
              id="header-btn-pdf"
              onClick={onExportPdf}
              className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold rounded bg-rose-700/80 hover:bg-rose-600 text-white shadow-sm border border-rose-500/50 transition-colors cursor-pointer"
              title="Exportar para PDF"
            >
              <FileText className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">PDF</span>
            </button>
          </div>
        )}

        {/* Filtros Ano, Mês e Setor para a tela Relação de Envelopes (ou Ano padrão para as demais) */}
        {currentView === 'relacao-envelopes' ? (
          <>
            {/* 1. Ano */}
            <div className="flex items-center gap-1.5 bg-[#252a40] px-2.5 py-1.5 rounded border border-[#394164]">
              <label htmlFor="select-ano-header" className="text-[11px] sm:text-xs text-slate-300 font-medium">
                Ano:
              </label>
              <select
                id="select-ano-header"
                value={anoSelecionado}
                onChange={(e) => onSelectAno(Number(e.target.value))}
                className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
              >
                <option value={2026} className="bg-[#1c2030] text-white">2026</option>
                <option value={2025} className="bg-[#1c2030] text-white">2025</option>
                <option value={2024} className="bg-[#1c2030] text-white">2024</option>
                <option value={2023} className="bg-[#1c2030] text-white">2023</option>
              </select>
            </div>

            {/* 2. Mês */}
            {mesSelecionado && onSelectMes && (
              <div className="flex items-center gap-1.5 bg-[#252a40] px-2.5 py-1.5 rounded border border-[#394164]">
                <label htmlFor="select-mes-header" className="text-[11px] sm:text-xs text-slate-300 font-medium">
                  Mês:
                </label>
                <select
                  id="select-mes-header"
                  value={mesSelecionado}
                  onChange={(e) => onSelectMes(e.target.value)}
                  className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
                >
                  {MESES_HEADER.map(m => (
                    <option key={m.valor} value={m.valor} className="bg-[#1c2030] text-white">
                      {m.label}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* 3. Setor */}
            {setorSelecionado && onSelectSetor && (
              <div className="flex items-center gap-1.5 bg-[#252a40] px-2.5 py-1.5 rounded border border-[#394164]">
                <label htmlFor="select-setor-header" className="text-[11px] sm:text-xs text-slate-300 font-medium">
                  Setor:
                </label>
                <select
                  id="select-setor-header"
                  value={setorSelecionado}
                  onChange={(e) => onSelectSetor(e.target.value)}
                  className="bg-white text-xs font-bold text-slate-900 px-2 py-0.5 rounded focus:outline-none cursor-pointer shadow-xs"
                >
                  {(setoresDisponiveis || ['Safira', 'Fire', 'White', 'Azul', 'Amarelo', 'Black', 'Diamante', 'Legacy', 'Onix', 'Titanium']).map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>
            )}
          </>
        ) : currentView === 'dashboard' ? (
          /* Ano selector apenas para a tela Dashboard */
          <div className="flex items-center gap-1.5 bg-[#252a40] px-2.5 py-1.5 rounded border border-[#394164]">
            <label htmlFor="select-ano-header" className="text-[11px] sm:text-xs text-slate-300 font-medium">
              Ano:
            </label>
            <select
              id="select-ano-header"
              value={anoSelecionado}
              onChange={(e) => onSelectAno(Number(e.target.value))}
              className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
            >
              <option value={2026} className="bg-[#1c2030] text-white">2026</option>
              <option value={2025} className="bg-[#1c2030] text-white">2025</option>
              <option value={2024} className="bg-[#1c2030] text-white">2024</option>
              <option value={2023} className="bg-[#1c2030] text-white">2023</option>
            </select>
          </div>
        ) : null}

        {/* Refresh button (botão de sincronizar os dados) */}
        <button
          id="header-refresh-btn"
          onClick={onRefresh}
          className="p-2 rounded-full hover:bg-[#282e46] text-slate-300 hover:text-white transition-colors cursor-pointer"
          title="Recarregar e sincronizar dados do banco"
          aria-label="Atualizar dados"
        >
          <RotateCw className={`w-4 h-4 ${isRefreshing ? 'animate-spin text-indigo-400' : ''}`} />
        </button>
      </div>
    </header>
  );
};
