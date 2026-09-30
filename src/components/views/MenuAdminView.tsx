import React from 'react';
import { ArrowLeft, User } from 'lucide-react';
import { ViewMode, MembroItem } from '../../types';

interface MenuAdminViewProps {
  onSelectView: (view: ViewMode) => void;
  usuarioLogado?: MembroItem | null;
}

export const MenuAdminView: React.FC<MenuAdminViewProps> = ({ onSelectView, usuarioLogado }) => {
  const menuItems: { id: ViewMode; label: string; highlight?: boolean }[] = [
    { id: 'validar-relatorios', label: 'Validar Relatórios' },
    { id: 'relacao-envelopes', label: 'Relação Envelopes' },
    { id: 'dashboard', label: 'DashBoard' },
    { id: 'indicador-trilho', label: 'Indicador Trilho de Capacitação', highlight: true },
    { id: 'fluxo-caixa', label: 'Fluxo de Caixa (Relatórios & SharePoint)' },
  ];

  return (
    <div 
      id="menu-admin-container" 
      className="min-h-screen w-full bg-[#1c2030] text-white p-4 sm:p-12 relative flex flex-col justify-between"
    >
      {/* Top Header with Back button and user display */}
      <div className="flex items-center justify-between mb-4 sm:mb-6">
        <button
          id="btn-back-menu-admin"
          onClick={() => onSelectView('dashboard')}
          className="w-10 h-10 rounded-full border-2 border-white flex items-center justify-center hover:bg-white hover:text-slate-900 transition-colors cursor-pointer shrink-0"
          title="Ir para o Dashboard"
        >
          <ArrowLeft className="w-5 h-5" />
        </button>

        <div className="text-center flex-1 px-4">
          <h2 className="text-xl sm:text-2xl font-black tracking-tight text-white">
            Menu Administrativo
          </h2>
          {usuarioLogado && (
            <div className="flex items-center justify-center gap-2 mt-1">
              <span className="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 text-xs font-semibold">
                <User className="w-3.5 h-3.5 text-indigo-400" />
                Usuário Conectado: <strong className="text-white">{usuarioLogado.nome}</strong> ({usuarioLogado.cargo || 'Membro'})
              </span>
            </div>
          )}
        </div>

        <div className="w-10 h-10 shrink-0" />
      </div>

      {/* Grid of Large White Buttons exactly matching the mobile design in menu.jpeg */}
      <div className="max-w-4xl mx-auto w-full grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 sm:gap-6 my-auto py-4">
        {menuItems.map((item) => (
          <button
            key={item.id}
            id={`menu-btn-${item.id}`}
            onClick={() => onSelectView(item.id)}
            className={`h-20 sm:h-28 rounded-xl font-bold text-sm sm:text-lg flex items-center justify-center p-3 sm:p-4 text-center shadow-lg transition-transform hover:scale-[1.02] sm:hover:scale-[1.03] active:scale-[0.98] cursor-pointer ${
              item.highlight
                ? 'bg-gradient-to-br from-indigo-50 to-white text-indigo-950 border-2 border-indigo-400 hover:bg-indigo-100'
                : 'bg-white text-slate-800 hover:bg-slate-50'
            }`}
          >
            {item.label}
          </button>
        ))}
      </div>

      {/* Footer Info */}
      <div className="text-center text-xs text-slate-500 pt-4">
        Paz Church Sobral • Sistema Integrado
      </div>
    </div>
  );
};
