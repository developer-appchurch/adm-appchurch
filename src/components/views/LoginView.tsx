import React, { useState, useEffect } from 'react';
import { User, RefreshCw, CheckCircle2, AlertCircle } from 'lucide-react';
import { SharePointService } from '../../services/sharepointService';
import { SharePointConfig, MembroItem } from '../../types';

interface LoginViewProps {
  onLoginSuccess: (membro: MembroItem) => void;
  configSharePoint?: SharePointConfig;
  onConfigChanged?: (newCfg: SharePointConfig) => void;
}

export const LoginView: React.FC<LoginViewProps> = ({ 
  onLoginSuccess,
  onConfigChanged 
}) => {
  const [login, setLogin] = useState('');
  const [senha, setSenha] = useState('');
  const [isAuthenticating, setIsAuthenticating] = useState(false);
  const [erroLogin, setErroLogin] = useState<string | null>(null);
  const [sucessoLogin, setSucessoLogin] = useState<string | null>(null);
  const [conexaoStatus, setConexaoStatus] = useState<'conectado' | 'erro' | 'verificando'>('verificando');
  const [statusDetalhe, setStatusDetalhe] = useState<string>('Verificando conexão com o SharePoint...');

  const spService = SharePointService.getInstance();

  // Verifica o status de conexão da API com o SharePoint ao carregar a tela
  useEffect(() => {
    let montado = true;

    const verificarConexao = async () => {
      try {
        const resp = await fetch('/api/sharepoint/status');
        if (!resp.ok) {
          if (montado) {
            setConexaoStatus('erro');
            setStatusDetalhe('Falha ao comunicar com o servidor');
          }
          return;
        }
        const data = await resp.json();
        if (montado) {
          if (data.status === 'CONECTADO') {
            setConexaoStatus('conectado');
            setStatusDetalhe(
              `Conectado ao SharePoint (${data.totalRelatorios ?? 0} relatórios • ${data.totalMembros ?? 0} membros)`
            );
          } else {
            setConexaoStatus('erro');
            setStatusDetalhe(data.erro || 'Não foi possível sincronizar com o SharePoint');
          }
        }
      } catch (err) {
        if (montado) {
          setConexaoStatus('erro');
          setStatusDetalhe('Erro de rede ao conectar com o SharePoint');
        }
      }
    };

    verificarConexao();

    // Revalida a cada 30 segundos
    const intervalo = setInterval(verificarConexao, 30000);
    return () => {
      montado = false;
      clearInterval(intervalo);
    };
  }, []);

  const handleAppSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErroLogin(null);
    setSucessoLogin(null);

    const loginLimpo = login.trim();
    const senhaLimpa = senha.trim();

    console.log(`[LoginView] Submetendo formulário de login para usuário: "${loginLimpo}"`);

    // Valida se login e senha foram preenchidos
    if (!loginLimpo || !senhaLimpa) {
      console.warn('[LoginView] Validação impedida: login ou senha em branco.');
      setErroLogin('Por favor, preencha o login e a senha.');
      return;
    }

    setIsAuthenticating(true);

    try {
      console.log('[LoginView] Solicitando validação ao SharePointService...');
      // Consulta oficial contra a lista BD_membros do SharePoint via backend com validação de login e senha
      const resultado = await spService.consultarMembroSharePoint(loginLimpo, senhaLimpa);
      console.log('[LoginView] Resultado recebido do SharePointService:', resultado);

      if (!resultado.sucesso || !resultado.membro) {
        console.warn(`[LoginView] Falha no login: ${resultado.erro}`);
        setErroLogin(resultado.erro || 'Login ou Senha incorretos');
        setIsAuthenticating(false);
        return;
      }

      console.log(`[LoginView] Login validado com sucesso para: ${resultado.membro.nome}! Persistindo sessão...`);
      setSucessoLogin(`Bem-vindo(a), ${resultado.membro.nome}!`);
      
      if (onConfigChanged) {
        onConfigChanged(spService.getConfig());
      }

      setTimeout(() => {
        console.log('[LoginView] Disparando callback onLoginSuccess com os dados do usuário.');
        onLoginSuccess(resultado.membro!);
      }, 300);
    } catch (err: any) {
      console.error('[LoginView] Erro inesperado durante o fluxo de autenticação:', err);
      setErroLogin(err?.message || 'Erro ao processar autenticação. Tente novamente.');
    } finally {
      setIsAuthenticating(false);
    }
  };

  return (
    <div 
      id="login-page-container" 
      className="min-h-screen w-full bg-[#1c2030] flex items-center justify-center p-4 relative"
    >
      <div 
        id="login-card" 
        className="bg-white rounded-2xl p-7 sm:p-9 w-full max-w-md shadow-2xl animate-in zoom-in-95 duration-200 text-center"
      >
        {/* Brand Header */}
        <div className="flex flex-col items-center justify-center mb-6">
          <div className="w-14 h-14 rounded-2xl bg-[#242a42] flex items-center justify-center shadow-md mb-3 text-indigo-400">
            <User className="w-7 h-7" />
          </div>
          <h1 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight">
            AppChurch
          </h1>
          <p className="text-xs font-semibold text-slate-500 mt-1">
            ADM Tesouraria • Paz Church Sobral
          </p>
        </div>

        {/* Mensagem de Erro destacada conforme regra de negócio */}
        {erroLogin && (
          <div 
            id="msg-erro-login"
            className="mb-5 p-3.5 rounded-xl bg-red-50 border-2 border-red-500 text-red-700 flex items-center gap-2.5 text-left animate-in fade-in slide-in-from-top-2 duration-150 shadow-sm"
          >
            <AlertCircle className="w-5 h-5 shrink-0 text-red-600" />
            <div className="flex-1">
              <p className="font-extrabold text-sm tracking-wide text-red-700">
                {erroLogin}
              </p>
              <p className="text-[11px] text-red-600/90 font-medium">
                Verifique se o login e a senha digitados estão corretos e tente novamente.
              </p>
            </div>
          </div>
        )}

        {/* Mensagem de Sucesso */}
        {sucessoLogin && (
          <div 
            id="msg-sucesso-login"
            className="mb-5 p-3.5 rounded-xl bg-emerald-50 border border-emerald-400 text-emerald-800 flex items-center gap-2.5 text-left animate-in fade-in duration-150"
          >
            <CheckCircle2 className="w-5 h-5 shrink-0 text-emerald-600" />
            <div>
              <p className="font-bold text-xs">{sucessoLogin}</p>
              <p className="text-[11px] text-emerald-700">Carregando o Menu Administrativo...</p>
            </div>
          </div>
        )}

        {/* Formulário Único de Login Direto App */}
        <form onSubmit={handleAppSubmit} className="space-y-4 text-left">
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 ml-1">
              Login do Usuário
            </label>
            <div className="relative">
              <input
                type="text"
                id="input-login"
                value={login}
                onChange={(e) => {
                  setLogin(e.target.value);
                  if (erroLogin) setErroLogin(null);
                }}
                placeholder="Digite seu login ou nome"
                className="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-2.5 text-sm text-slate-800 focus:outline-none focus:border-indigo-600 focus:bg-white transition-colors"
                required
                autoFocus
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1.5 ml-1">
              Senha
            </label>
            <div className="relative">
              <input
                type="password"
                id="input-senha"
                value={senha}
                onChange={(e) => {
                  setSenha(e.target.value);
                  if (erroLogin) setErroLogin(null);
                }}
                placeholder="Digite sua senha"
                className="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-2.5 text-sm text-slate-800 focus:outline-none focus:border-indigo-600 focus:bg-white transition-colors"
                required
              />
            </div>
          </div>

          <button
            type="submit"
            id="btn-login-submit"
            disabled={isAuthenticating}
            className="w-full mt-2 bg-[#242a42] hover:bg-[#2e3655] text-white font-bold py-3 px-4 rounded-xl text-sm shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-60"
          >
            {isAuthenticating ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin text-indigo-400" />
                <span>Verificando credenciais...</span>
              </>
            ) : (
              <span>Entrar</span>
            )}
          </button>
        </form>
      </div>

      {/* Sinalizador Circular no canto inferior direito */}
      <div 
        id="sharepoint-status-indicator"
        className="fixed bottom-4 right-4 z-50 flex items-center gap-2 select-none"
        title={statusDetalhe}
      >
        <span className="relative flex h-3.5 w-3.5">
          {conexaoStatus === 'conectado' && (
            <>
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-3.5 w-3.5 bg-emerald-500 shadow-md shadow-emerald-500/50 border border-white/20"></span>
            </>
          )}
          {conexaoStatus === 'erro' && (
            <>
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-60"></span>
              <span className="relative inline-flex rounded-full h-3.5 w-3.5 bg-rose-500 shadow-md shadow-rose-500/50 border border-white/20"></span>
            </>
          )}
          {conexaoStatus === 'verificando' && (
            <span className="relative inline-flex rounded-full h-3.5 w-3.5 bg-amber-400 animate-pulse border border-white/20"></span>
          )}
        </span>
      </div>
    </div>
  );
};
