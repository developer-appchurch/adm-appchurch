import React, { useState, useEffect, useCallback } from 'react';
import { ViewMode, LancamentoTesouraria, SharePointConfig, MembroItem } from './types';
import { SharePointService } from './services/sharepointService';
import { Sidebar } from './components/Sidebar';
import { Header } from './components/Header';
import { FluxoCaixaView } from './components/views/FluxoCaixaView';
import { DashboardView } from './components/views/DashboardView';
import { RelacaoEnvelopesView } from './components/views/RelacaoEnvelopesView';
import { ValidarRelatoriosView } from './components/views/ValidarRelatoriosView';
import { IndicadorTrilhoView } from './components/views/IndicadorTrilhoView';
import { MenuAdminView } from './components/views/MenuAdminView';
import { LoginView } from './components/views/LoginView';

export default function App() {
  const spService = SharePointService.getInstance();

  // A primeira tela que deve aparecer é o login
  const [currentView, setCurrentView] = useState<ViewMode>('login');
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
  const [anoSelecionado, setAnoSelecionado] = useState<number>(2026);
  const [mesEnvelopes, setMesEnvelopes] = useState<string>('9');
  const [setorEnvelopes, setSetorEnvelopes] = useState<string>('Safira');
  const [setoresDisponiveisEnvelopes, setSetoresDisponiveisEnvelopes] = useState<string[]>([]);
  const [lancamentos, setLancamentos] = useState<LancamentoTesouraria[]>([]);
  const [sharePointConfig, setSharePointConfig] = useState<SharePointConfig>(spService.getConfig());
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [notificacao, setNotificacao] = useState<string | null>(null);
  const [isMobileNavOpen, setIsMobileNavOpen] = useState(false);

  const showNotification = (msg: string) => {
    setNotificacao(msg);
    setTimeout(() => {
      setNotificacao(null);
    }, 4000);
  };

  // Carrega lançamentos iniciais e conecta automaticamente ao SharePoint
  const carregarDados = useCallback(() => {
    const dados = spService.getLancamentos();
    console.log(`[App] carregarDados executado. Lançamentos carregados: ${dados.length}`);
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

  // Atualiza dados sempre que a visualização mudar
  useEffect(() => {
    carregarDados();
    if (currentView !== 'login' && spService.getLancamentos().length === 0) {
      sincronizarDadosCompletos();
    }
  }, [currentView, carregarDados, sincronizarDadosCompletos, spService]);

  // Ação de Atualizar / Refresh
  const handleRefresh = async () => {
    const res = await sincronizarDadosCompletos();
    showNotification(`Sincronizado com SharePoint: ${res.relatoriosCount} relatórios e ${res.membrosCount} membros.`);
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

    // Navega diretamente para a tela de Menu do aplicativo
    console.log('[App] Mudando currentView de "login" para "menu-admin"...');
    setCurrentView('menu-admin');
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

  // Se a view for Menu Administrativo (navega para cá após login bem-sucedido)
  if (currentView === 'menu-admin') {
    return (
      <MenuAdminView 
        onSelectView={setCurrentView} 
        usuarioLogado={usuarioLogado}
      />
    );
  }

  return (
    <div id="app-root-container" className="flex h-screen bg-[#1c2030] text-slate-100 overflow-hidden font-sans">
      {/* Sidebar de Navegação */}
      <Sidebar
        currentView={currentView}
        onSelectView={(v) => {
          setCurrentView(v);
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
            onSelectView={setCurrentView}
            anoSelecionado={anoSelecionado}
            onSelectAno={setAnoSelecionado}
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

        {/* Área de Visualização com Scroll */}
        <main className="flex-1 overflow-y-auto bg-[#1c2030] scrollbar-thin scrollbar-thumb-[#313752] scrollbar-track-[#1c2030]">
          {currentView === 'fluxo-caixa' && (
            <FluxoCaixaView
              todosLancamentos={lancamentos}
              lancamentos={lancamentos}
              anoSelecionado={anoSelecionado}
              onSelectAno={setAnoSelecionado}
              onRefresh={handleRefresh}
              onAtualizarDados={handleRefresh}
            />
          )}

          {currentView === 'dashboard' && (
            <DashboardView
              lancamentos={lancamentos}
              anoSelecionado={anoSelecionado}
              onSelectAno={setAnoSelecionado}
              onRefresh={handleRefresh}
              usuarioConectado={sharePointConfig.usuarioConectado}
              onSelectView={setCurrentView}
              onToggleMobileMenu={() => setIsMobileNavOpen(prev => !prev)}
            />
          )}

          {currentView === 'relacao-envelopes' && (
            <RelacaoEnvelopesView
              lancamentos={lancamentos}
              anoSelecionado={anoSelecionado}
              onSelectAno={setAnoSelecionado}
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
              anoSelecionado={anoSelecionado}
              onSelectAno={setAnoSelecionado}
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
