import React, { useState } from 'react';
import { Database, Lock, Globe, CheckCircle2, AlertCircle, RefreshCw, X, Eye, EyeOff, ShieldCheck, Server } from 'lucide-react';
import { SharePointService } from '../../services/sharepointService';

interface SharePointLoginModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConectadoComSucesso?: () => void;
}

export const SharePointLoginModal: React.FC<SharePointLoginModalProps> = ({
  isOpen,
  onClose,
  onConectadoComSucesso
}) => {
  const spService = SharePointService.getInstance();
  const currentCfg = spService.getConfig();

  const [username, setUsername] = useState('midia.sobral@paz.church');
  const [password, setPassword] = useState('Pazsobral23');
  const [siteUrl, setSiteUrl] = useState(currentCfg.siteUrl || 'https://pazchurch.sharepoint.com/sites/PazSobral');
  const [clientId, setClientId] = useState('d3590ed6-52b3-4102-aeff-aad2292ab01c');
  const [showPassword, setShowPassword] = useState(false);
  
  const [isConnecting, setIsConnecting] = useState(false);
  const [etapaStatus, setEtapaStatus] = useState<string | null>(null);
  const [erro, setErro] = useState<string | null>(null);
  const [sucesso, setSucesso] = useState<{
    membrosCount: number;
    relatoriosCount: number;
    celulasCount: number;
  } | null>(null);

  if (!isOpen) return null;

  const handleConectar = async (e: React.FormEvent) => {
    e.preventDefault();
    setErro(null);
    setSucesso(null);
    setIsConnecting(true);
    setEtapaStatus('1/3: Solicitando Token OAuth2 na Microsoft...');

    try {
      setEtapaStatus('2/3: Validando credenciais e acessando listas no SharePoint...');
      const resultado = await spService.conectarComCredenciaisSharePoint({
        username: username.trim(),
        password: password.trim(),
        siteUrl: siteUrl.trim(),
        clientId: clientId.trim()
      });

      if (!resultado.sucesso) {
        setErro(resultado.erro || 'Falha na autenticação com o SharePoint. Verifique as credenciais.');
        setIsConnecting(false);
        setEtapaStatus(null);
        return;
      }

      setEtapaStatus('3/3: Sincronizando BD_membros e relatórios...');
      setSucesso({
        membrosCount: resultado.membrosCount ?? 0,
        relatoriosCount: resultado.relatoriosCount ?? 0,
        celulasCount: resultado.celulasCount ?? 0
      });

      if (onConectadoComSucesso) {
        onConectadoComSucesso();
      }

      setTimeout(() => {
        setIsConnecting(false);
        setEtapaStatus(null);
      }, 500);
    } catch (err: any) {
      console.error('[SharePointLoginModal] Erro ao conectar:', err);
      setErro(err?.message || 'Erro inesperado ao conectar com o SharePoint.');
      setIsConnecting(false);
      setEtapaStatus(null);
    }
  };

  const preencherPadrao = () => {
    setUsername('midia.sobral@paz.church');
    setPassword('Pazsobral23');
    setSiteUrl('https://pazchurch.sharepoint.com/sites/PazSobral');
    setClientId('d3590ed6-52b3-4102-aeff-aad2292ab01c');
    setErro(null);
  };

  return (
    <div 
      id="sharepoint-login-modal-overlay" 
      className="fixed inset-0 z-[100] bg-black/75 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 animate-in fade-in duration-200"
    >
      <div 
        id="sharepoint-login-modal-card" 
        className="bg-[#181c2b] border border-[#2e3652] w-full max-w-lg rounded-2xl shadow-2xl overflow-hidden text-slate-100 animate-in zoom-in-95 duration-200"
      >
        {/* Modal Header */}
        <div className="bg-[#1f2438] px-5 py-4 border-b border-[#2e3652] flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-indigo-600/30 border border-indigo-500/40 flex items-center justify-center text-indigo-400">
              <Database className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold text-white flex items-center gap-1.5">
                Autenticação Microsoft SharePoint
              </h2>
              <p className="text-[11px] text-slate-400">
                Paz Church Sobral • Conexão do Banco de Dados
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-[#2a314b] transition-colors cursor-pointer"
            title="Fechar"
            aria-label="Fechar"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 sm:p-6 space-y-4 max-h-[80vh] overflow-y-auto">
          {/* Info Banner */}
          <div className="bg-[#20273f] border border-indigo-500/30 rounded-xl p-3.5 flex items-start gap-2.5 text-xs text-slate-300">
            <Server className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <p className="font-semibold text-white">Conexão do Sistema com as Listas do SharePoint</p>
              <p className="text-[11px] text-slate-400 leading-relaxed">
                Digite as credenciais da conta Microsoft que tem acesso ao site do SharePoint (listas <strong className="text-slate-200">BD_membros</strong>, <strong className="text-slate-200">BD_Relatorio</strong> e <strong className="text-slate-200">BD_celulas</strong>). Após conectar, você poderá realizar seu login pessoal de membro.
              </p>
            </div>
          </div>

          {/* Feedback de Erro */}
          {erro && (
            <div className="p-3 rounded-xl bg-red-950/60 border border-red-500/50 text-red-200 text-xs flex items-start gap-2.5 animate-in slide-in-from-top-1">
              <AlertCircle className="w-4 h-4 text-red-400 shrink-0 mt-0.5" />
              <div className="flex-1">
                <p className="font-bold text-red-300">Erro de Conexão com SharePoint:</p>
                <p className="text-[11px] text-red-200 mt-0.5">{erro}</p>
              </div>
            </div>
          )}

          {/* Feedback de Sucesso */}
          {sucesso && (
            <div className="p-3.5 rounded-xl bg-emerald-950/60 border border-emerald-500/50 text-emerald-100 text-xs flex items-start gap-2.5 animate-in slide-in-from-top-1">
              <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" />
              <div className="flex-1 space-y-1">
                <p className="font-bold text-emerald-300 text-sm">SharePoint Conectado com Sucesso!</p>
                <p className="text-[11px] text-emerald-200">
                  Base de dados pronta: <strong>{sucesso.membrosCount} membros</strong> sincronizados e relatórios carregados.
                </p>
                <p className="text-[10px] text-emerald-300/80 font-medium pt-1">
                  Agora você pode fechar esta tela e efetuar seu login pessoal.
                </p>
              </div>
            </div>
          )}

          {/* Status do processo de conexão */}
          {isConnecting && etapaStatus && (
            <div className="p-3 rounded-xl bg-[#20273f] border border-amber-500/40 text-amber-200 text-xs flex items-center gap-2.5 animate-pulse">
              <RefreshCw className="w-4 h-4 text-amber-400 animate-spin shrink-0" />
              <span>{etapaStatus}</span>
            </div>
          )}

          {/* Formulário de Credenciais do SharePoint */}
          <form onSubmit={handleConectar} className="space-y-3.5 text-xs text-left">
            <div>
              <label className="block text-[11px] font-bold text-slate-300 mb-1">
                E-mail / Usuário Microsoft 365
              </label>
              <div className="relative">
                <input
                  type="text"
                  id="input-sp-username"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="ex: midia.sobral@paz.church"
                  className="w-full bg-[#131622] border border-[#2f3754] rounded-xl px-3.5 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-400 transition-colors font-mono"
                  required
                />
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-300 mb-1">
                Senha da Conta Microsoft
              </label>
              <div className="relative">
                <input
                  type={showPassword ? 'text' : 'password'}
                  id="input-sp-password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Digite a senha da conta Microsoft"
                  className="w-full bg-[#131622] border border-[#2f3754] rounded-xl pl-3.5 pr-10 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-400 transition-colors font-mono"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white p-1"
                  tabIndex={-1}
                >
                  {showPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                </button>
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-300 mb-1">
                URL do Site SharePoint
              </label>
              <div className="relative">
                <input
                  type="text"
                  id="input-sp-siteurl"
                  value={siteUrl}
                  onChange={(e) => setSiteUrl(e.target.value)}
                  placeholder="https://pazchurch.sharepoint.com/sites/PazSobral"
                  className="w-full bg-[#131622] border border-[#2f3754] rounded-xl px-3.5 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-400 transition-colors font-mono text-[11px]"
                  required
                />
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-bold text-slate-400 mb-1">
                Microsoft Client ID (Padrão Office Public Client)
              </label>
              <input
                type="text"
                id="input-sp-clientid"
                value={clientId}
                onChange={(e) => setClientId(e.target.value)}
                className="w-full bg-[#131622]/60 border border-[#262c44] rounded-xl px-3.5 py-2 text-[10px] text-slate-400 focus:outline-none focus:border-indigo-400 transition-colors font-mono"
              />
            </div>

            <div className="pt-2 flex flex-col gap-2">
              <button
                type="submit"
                id="btn-conectar-sharepoint"
                disabled={isConnecting}
                className="w-full bg-gradient-to-r from-indigo-600 to-blue-600 hover:from-indigo-500 hover:to-blue-500 text-white font-bold py-3 px-4 rounded-xl text-xs shadow-lg shadow-indigo-600/30 transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-60"
              >
                {isConnecting ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    <span>Conectando e Validando Listas...</span>
                  </>
                ) : (
                  <>
                    <ShieldCheck className="w-4 h-4" />
                    <span>Testar e Conectar ao SharePoint</span>
                  </>
                )}
              </button>

              <div className="flex items-center justify-between gap-2 pt-1">
                <button
                  type="button"
                  onClick={preencherPadrao}
                  className="text-[11px] text-indigo-400 hover:text-indigo-300 underline cursor-pointer"
                >
                  Restaurar credenciais padrão da igreja
                </button>

                {sucesso && (
                  <button
                    type="button"
                    onClick={onClose}
                    className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg text-xs transition-colors cursor-pointer"
                  >
                    Fazer Login Pessoal →
                  </button>
                )}
              </div>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
};
