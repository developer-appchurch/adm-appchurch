import React from 'react';
import { 
  ClipboardCheck, 
  FileSpreadsheet, 
  LayoutDashboard, 
  TrendingUp, 
  User,
  LogOut,
  X,
  GraduationCap
} from 'lucide-react';
import { ViewMode } from '../types';

interface SidebarProps {
  currentView: ViewMode;
  onSelectView: (view: ViewMode) => void;
  sharepointStatus: string;
  totalSincronizado: number;
  usuarioConectado?: { nome: string; email: string; cargo?: string } | null;
  isMobileOpen?: boolean;
  onCloseMobile?: () => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  currentView,
  onSelectView,
  sharepointStatus,
  totalSincronizado,
  usuarioConectado,
  isMobileOpen = false,
  onCloseMobile
}) => {
  const menuItems: {
    id: ViewMode;
    label: string;
    icon: React.ComponentType<{ className?: string }>;
    badge?: string;
  }[] = [
    {
      id: 'validar-relatorios' as ViewMode,
      label: 'Validar Relatórios',
      icon: ClipboardCheck
    },
    {
      id: 'relacao-envelopes' as ViewMode,
      label: 'Relação Envelopes',
      icon: FileSpreadsheet
    },
    {
      id: 'fluxo-caixa' as ViewMode,
      label: 'Fluxo de Caixa',
      icon: TrendingUp
    },
    {
      id: 'dashboard' as ViewMode,
      label: 'DashBoard',
      icon: LayoutDashboard
    },
    {
      id: 'indicador-trilho' as ViewMode,
      label: 'Indicador Trilho',
      icon: GraduationCap
    }
  ];

  const renderContent = (isMobile: boolean) => (
    <>
      <div>
        {/* Logo Section */}
        <div id="adm-brand-header" className="px-5 pt-4 pb-3.5 border-b border-[#25293d] flex items-center justify-between">
          <div className="flex flex-col items-start select-none max-w-full">
            <div className="h-[62px] w-[210px] max-w-full overflow-hidden flex items-center">
              <img
                src="/logo-appchurch-branca.png"
                alt="AppChurch"
                className="w-full h-full object-contain object-left select-none"
                draggable={false}
                onError={(e) => {
                  (e.target as HTMLImageElement).src = '/logo-appchurch.png';
                }}
              />
            </div>
            <span className="text-xs font-semibold text-slate-300 mt-1 tracking-wide">
              Tesouraria
            </span>
          </div>
          {isMobile && (
            <button
              onClick={onCloseMobile}
              className="p-2 -mr-1 rounded-lg text-slate-400 hover:text-white hover:bg-[#25293d] cursor-pointer"
              aria-label="Fechar menu de navegação"
            >
              <X className="w-5 h-5" />
            </button>
          )}
        </div>

        {/* Navigation Menu */}
        <nav id="sidebar-navigation" className="px-3 py-4 space-y-1.5" aria-label="Navegação Principal">
          {menuItems.map((item) => {
            const Icon = item.icon;
            const isActive = currentView === item.id;

            return (
              <button
                key={item.id}
                id={`nav-btn-${item.id}${isMobile ? '-mob' : ''}`}
                onClick={() => {
                  if (currentView !== item.id) {
                    onSelectView(item.id);
                  }
                  if (isMobile && onCloseMobile) {
                    onCloseMobile();
                  }
                }}
                className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-lg text-sm font-medium transition-all duration-150 text-left cursor-pointer ${
                  isActive
                    ? 'bg-[#282d46] text-white shadow-md border border-[#3b4366]'
                    : 'text-slate-300 hover:bg-[#202438] hover:text-white'
                }`}
              >
                <div className="flex items-center gap-3">
                  <Icon className={`w-5 h-5 shrink-0 ${isActive ? 'text-indigo-300' : 'text-slate-400'}`} />
                  <span className="truncate">{item.label}</span>
                </div>
                {item.badge && (
                  <span className="text-[10px] uppercase font-bold tracking-wider px-1.5 py-0.5 rounded bg-emerald-600/30 text-emerald-300 border border-emerald-500/40">
                    {item.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>
      </div>

      {/* Footer / Ações de Sessão */}
      <div id="sidebar-footer" className="p-3 border-t border-[#25293d] space-y-2 mt-auto">
        <div className="flex items-center gap-1.5">
          <button
            id={`btn-sidebar-login-sp${isMobile ? '-mob' : ''}`}
            onClick={() => {
              onSelectView('login');
              if (isMobile && onCloseMobile) {
                onCloseMobile();
              }
            }}
            className="flex-1 flex items-center justify-center gap-1.5 py-2 px-2.5 rounded-md text-xs font-medium bg-[#1b1e2f] text-slate-300 hover:text-white hover:bg-[#25293d] border border-[#2b3048] transition-colors cursor-pointer"
            title="Ir para tela de Login / Trocar Usuário"
          >
            <User className="w-3.5 h-3.5 text-indigo-400 shrink-0" />
            <span>Trocar Usuário</span>
          </button>

          <button
            id={`btn-logout${isMobile ? '-mob' : ''}`}
            onClick={() => {
              onSelectView('login');
              if (isMobile && onCloseMobile) {
                onCloseMobile();
              }
            }}
            className="flex items-center justify-center py-2 px-3 rounded-md text-xs font-medium bg-[#1b1e2f] text-slate-400 hover:text-red-300 hover:bg-red-950/20 border border-[#2b3048] cursor-pointer gap-1.5"
            title="Sair do Sistema"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Sair</span>
          </button>
        </div>
      </div>
    </>
  );

  return (
    <>
      {/* Desktop Persistent Sidebar */}
      <aside 
        id="sidebar-container" 
        className="hidden md:flex w-64 bg-[#181a28] text-slate-200 border-r border-[#2a2f48] flex-col justify-between select-none shrink-0 h-full overflow-y-auto"
      >
        {renderContent(false)}
      </aside>

      {/* Mobile Drawer Overlay & Sidebar */}
      {isMobileOpen && (
        <div className="fixed inset-0 z-50 flex md:hidden" id="mobile-sidebar-drawer">
          <div 
            className="fixed inset-0 bg-black/75 backdrop-blur-xs transition-opacity"
            onClick={onCloseMobile}
            aria-hidden="true"
          />
          <aside 
            className="relative w-72 max-w-[85vw] h-full bg-[#181a28] text-slate-200 border-r border-[#2a2f48] flex flex-col justify-between select-none z-10 shadow-2xl overflow-y-auto animate-in slide-in-from-left duration-200"
          >
            {renderContent(true)}
          </aside>
        </div>
      )}
    </>
  );
};
