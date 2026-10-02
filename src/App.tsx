import React, { useState, useEffect, useCallback } from 'react';
import { RotateCw } from 'lucide-react';
import { ViewMode, LancamentoTesouraria, SharePointConfig, MembroItem } from './types';
import { SharePointService } from './services/sharepointService';
import { Sidebar } from './components/Sidebar';
import { Header } from './components/Header';
import { FluxoCaixaView } from './components/views/FluxoCaixaView';
import { DashboardView } from './components/views/DashboardView';
import { RelacaoEnvelopesView } from './components/views/RelacaoEnvelopesView';
import { ValidarRelatoriosView } from './components/views/ValidarRelatoriosView';
import { IndicadorTrilhoView } from './components/views/IndicadorTrilhoView';
import { LoginView } from './components/views/LoginView';

export default function App() {
  const spService = SharePointService.getInstance();

  // A primeira tela que deve aparecer é o login
  const [currentView, setCurrentView] = useState<ViewMode>('login');
  const [refreshKey, setRefreshKey] = useState<number>(0);
  // Dados do usuário logado salvos em variável de estado (com padrão Junio Fonteles - ID: 4)
  const [usuarioLogado, setUsuarioLogado] = useState<MembroItem | null>(() => {
    try {
      const saved = localStorage.getItem('tesouraria_usuario_logado');
      if (saved) return JSON.parse(saved);
    } catch {}
    return {
      id: 4,
      ID: 4,
      nome: 'Junio Fonteles',
      login: 'Jfonteles',
      email: 'juniosina@hotmail.com',
      cargo: 'Líder de Setor',
      celula: 'Adonai',
      setor: 'Fire'
    };
  });
  // Estados de Filtro Independentes por Tela (Dashboard, Validar Relatórios, Envelopes, Fluxo de Caixa)
  const [anoDashboard, setAnoDashboard] = useState<number | string>(2026);
  const [anoValidar, setAnoValidar] = useState<number | string>(2026);
  const [anoEnvelopes, setAnoEnvelopes] = useState<number | string>(2026);
  const [mesEnvelopes, setMesEnvelopes] = useState<string>('todos');
  const [setorEnvelopes, setSetorEnvelopes] = useState<string>('Safira');
  const [anoFluxo, setAnoFluxo] = useState<number | string>(2026);
  const [setoresDisponiveisEnvelopes, setSetoresDisponiveisEnvelopes] = useState<string[]>([]);
  const [lancamentos, setLancamentos] = useState<LancamentoTesouraria[]>([]);
  const [sharePointConfig, setSharePointConfig] = useState<SharePointConfig>(spService.getConfig());
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [notificacao, setNotificacao] = useState<string | null>(null);
  const [isMobileNavOpen, setIsMobileNavOpen] = useState(false);

  // Navegação segura evitando recarregamento duplo ao clicar na mesma tela
  const handleSelectView = useCallback((novaView: ViewMode) => {
    setCurrentView(prev => (prev === novaView ? prev : novaView));
  }, []);

  const showNotification = (msg: string) => {
    setNotificacao(msg);
    setTimeout(() => {
      setNotificacao(null);
    }, 4000);
  };

  // Carrega lançamentos iniciais e conecta automaticamente ao SharePoint
  const carregarDados = useCallback(() => {
    const dados = spService.getLancamentos();
    setLancamentos(dados);
    setSharePointConfig(spService.getConfig());
  }, [spService]);

  // Sincroniza ativamente com o backend e atualiza o estado
  const sincronizarDadosCompletos = useCallback(async () => {
    setIsRefreshing(true);
    try {
      const res = await spService.conectarEAtualizarAutomatico();
      carregarDados();
      return res;
    } finally {
      setIsRefreshing(false);
    }
  }, [spService, carregarDados]);

  useEffect(() => {
    carregarDados();
    // Conecta automaticamente a todas as listas do SharePoint ao ser executado
    sincronizarDadosCompletos();
  }, [carregarDados, sincronizarDadosCompletos]);

  // Sincroniza apenas caso a base esteja vazia ao navegar para outra view, sem re-render duplo
  useEffect(() => {
    if (currentView !== 'login' && lancamentos.length === 0 && spService.getLancamentos().length === 0) {
      sincronizarDadosCompletos();
    }
  }, [currentView, lancamentos.length, sincronizarDadosCompletos, spService]);

  // Ação de Atualizar / Refresh
  const handleRefresh = async () => {
    setIsRefreshing(true);
    try {
      const res = await sincronizarDadosCompletos();
      setRefreshKey(prev => prev + 1);
      carregarDados();
      showNotification(`Dados sincronizados com o SharePoint com sucesso!`);
      return res;
    } finally {
      setIsRefreshing(false);
    }
  };

  // Confirmar validação de envelope com dados do usuário logado
  const handleConfirmarLancamento = (id: string, idTesoureiro?: string | number, dataTesouraria?: string) => {
    const idFinal = idTesoureiro ?? (usuarioLogado?.id || usuarioLogado?.ID || 4);
    spService.confirmarLancamento(id, idFinal, dataTesouraria);
    carregarDados();
    showNotification('Lançamento validado e confirmado com sucesso.');
  };

  // Desconfirmar validação de envelope
  const handleDesconfirmarLancamento = (id: string) => {
    spService.desconfirmarLancamento(id);
    carregarDados();
    showNotification('Confirmação do lançamento desfeita com sucesso.');
  };

  // Atualizar dados de um lançamento
  const handleAtualizarLancamento = (id: string, dados: Partial<LancamentoTesouraria>) => {
    spService.atualizarLancamento(id, dados);
    carregarDados();
    showNotification('Lançamento atualizado e sincronizado com a base.');
  };

  // Salvar nova configuração do SharePoint
  const handleSharePointConfigSaved = (newCfg: SharePointConfig) => {
    setSharePointConfig(newCfg);
    carregarDados();
    showNotification('Conexão e lista do SharePoint atualizadas.');
  };

  // Callback de sucesso ao autenticar usuário presente na tabela BD_membros
  const handleLoginSuccess = async (membro: MembroItem) => {
    console.log('[App] handleLoginSuccess recebido para:', membro.nome);
    
    // Salva o nome e os dados do usuário em variável de estado
    setUsuarioLogado(membro);
    try {
      localStorage.setItem('tesouraria_usuario_logado', JSON.stringify(membro));
      console.log('[App] tesouraria_usuario_logado salvo no localStorage.');
    } catch (e) {
      console.warn('[App] Aviso ao salvar usuário logado no localStorage:', e);
    }

    // Atualiza a sessão e configuração geral
    try {
      const updatedCfg = spService.updateConfig({
        usuarioConectado: {
          nome: membro.nome || membro.Title || 'Admin',
          email: membro.email || `${membro.login}@pazchurch.com`,
          cargo: membro.cargo || 'Membro / Liderança',
          conectadoEm: new Date().toISOString().replace('T', ' ').slice(0, 19)
        }
      });
      setSharePointConfig(updatedCfg);
    } catch (eCfg) {
      console.warn('[App] Aviso ao atualizar config com usuário:', eCfg);
    }

    // Navega diretamente para a tela de Validar Relatórios
    console.log('[App] Mudando currentView de "login" para "validar-relatorios"...');
    setCurrentView('validar-relatorios');
    carregarDados();
    showNotification(`Bem-vindo(a), ${membro.nome}!`);

    // Sincroniza em background para garantir que todos os relatórios estejam carregados
    sincronizarDadosCompletos().then(() => {
      carregarDados();
    });
  };

  // Se a view for Login
  if (currentView === 'login') {
    return (
      <LoginView 
        onLoginSuccess={handleLoginSuccess}
        configSharePoint={sharePointConfig}
        onConfigChanged={handleSharePointConfigSaved}
        onSelectView={setCurrentView}
      />
    );
  }

  const anoAtivoHeader = currentView === 'relacao-envelopes' ? anoEnvelopes : currentView === 'validar-relatorios' ? anoValidar : currentView === 'fluxo-caixa' ? anoFluxo : anoDashboard;
  const setAnoAtivoHeader = currentView === 'relacao-envelopes' ? setAnoEnvelopes : currentView === 'validar-relatorios' ? setAnoValidar : currentView === 'fluxo-caixa' ? setAnoFluxo : setAnoDashboard;

  return (
    <div id="app-root-container" className="flex h-screen bg-[#1c2030] text-slate-100 overflow-hidden font-sans">
      {/* Sidebar de Navegação */}
      <Sidebar
        currentView={currentView}
        onSelectView={(v) => {
          handleSelectView(v);
          setIsMobileNavOpen(false);
        }}
        sharepointStatus={sharePointConfig.status}
        totalSincronizado={sharePointConfig.totalItensSincronizados}
        usuarioConectado={sharePointConfig.usuarioConectado}
        isMobileOpen={isMobileNavOpen}
        onCloseMobile={() => setIsMobileNavOpen(false)}
      />

      {/* Conteúdo Principal */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Header Superior (exibido para visualizações que não possuem barra própria) */}
        {currentView !== 'dashboard' && (
          <Header
            currentView={currentView}
            onSelectView={handleSelectView}
            anoSelecionado={anoAtivoHeader}
            onSelectAno={setAnoAtivoHeader}
            lancamentos={lancamentos}
            usuarioConectado={sharePointConfig.usuarioConectado}
            onRefresh={handleRefresh}
            isRefreshing={isRefreshing}
            onToggleMobileMenu={() => setIsMobileNavOpen(prev => !prev)}
            mesSelecionado={mesEnvelopes}
            onSelectMes={setMesEnvelopes}
            setorSelecionado={setorEnvelopes}
            onSelectSetor={setSetorEnvelopes}
            setoresDisponiveis={setoresDisponiveisEnvelopes}
          />
        )}

        {/* Notificação Toast */}
        {notificacao && (
          <div className="fixed bottom-5 right-5 z-50 bg-[#161a29] border border-emerald-500/40 text-white px-4 py-3 rounded-lg shadow-xl flex items-center gap-3 text-xs animate-in slide-in-from-bottom-2 duration-300">
            <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
            <span>{notificacao}</span>
          </div>
        )}

        {/* Overlay de Carregamento e Sincronização em Tempo Real */}
        {isRefreshing && (
          <div className="fixed inset-0 z-[120] bg-black/60 backdrop-blur-xs flex flex-col items-center justify-center p-4 animate-in fade-in duration-150 select-none">
            <div className="bg-[#161a29] border border-[#2e3752] p-6 rounded-2xl shadow-2xl flex flex-col items-center text-center max-w-sm w-full space-y-3.5">
              <div className="w-12 h-12 rounded-2xl bg-indigo-600/25 border border-indigo-500/40 flex items-center justify-center text-indigo-400 shadow-md">
                <RotateCw className="w-6 h-6 animate-spin text-indigo-400" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-white">
                  Sincronizando com o SharePoint
                </h3>
                <p className="text-xs text-slate-400 mt-1">
                  Atualizando tabelas e recalculando indicadores da tela...
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Área de Visualização com Scroll */}
        <main 
          key={`view-${currentView}-${refreshKey}`}
          className="flex-1 overflow-y-auto bg-[#1c2030] scrollbar-thin scrollbar-thumb-[#313752] scrollbar-track-[#1c2030]"
        >
          {currentView === 'fluxo-caixa' && (
            <FluxoCaixaView
              anoSelecionado={anoFluxo}
              onSelectAno={setAnoFluxo}
              onRefresh={handleRefresh}
              onAtualizarDados={handleRefresh}
              usuarioLogado={usuarioLogado}
            />
          )}

          {currentView === 'dashboard' && (
            <DashboardView
              lancamentos={lancamentos}
              anoSelecionado={anoDashboard}
              onSelectAno={setAnoDashboard}
              onRefresh={handleRefresh}
              usuarioConectado={sharePointConfig.usuarioConectado}
              onSelectView={handleSelectView}
              onToggleMobileMenu={() => setIsMobileNavOpen(prev => !prev)}
            />
          )}

          {currentView === 'relacao-envelopes' && (
            <RelacaoEnvelopesView
              lancamentos={lancamentos}
              anoSelecionado={anoEnvelopes}
              onSelectAno={setAnoEnvelopes}
              mesSelecionado={mesEnvelopes}
              onSelectMes={setMesEnvelopes}
              setorSelecionado={setorEnvelopes}
              onSelectSetor={setSetorEnvelopes}
              onSetoresDisponiveisChange={setSetoresDisponiveisEnvelopes}
              onRefresh={handleRefresh}
            />
          )}

          {currentView === 'validar-relatorios' && (
            <ValidarRelatoriosView
              lancamentos={lancamentos}
              anoSelecionado={anoValidar}
              onSelectAno={setAnoValidar}
              usuarioLogado={usuarioLogado}
              onAtualizarLancamento={handleAtualizarLancamento}
              onConfirmarLancamento={handleConfirmarLancamento}
              onDesconfirmarLancamento={handleDesconfirmarLancamento}
              onRefresh={handleRefresh}
            />
          )}

          {currentView === 'indicador-trilho' && (
            <IndicadorTrilhoView />
          )}
        </main>
      </div>
    </div>
  );
}
