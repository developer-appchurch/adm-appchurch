import React, { useState } from 'react';
import { Database, Plus, Trash2, CheckCircle2, RefreshCw, ShieldCheck, UserCheck, Lock } from 'lucide-react';
import { SharePointConfig } from '../../types';
import { CATEGORIAS_ENTRADA, CATEGORIAS_SAIDA } from '../../data/mockSharePointData';
import { SharePointService } from '../../services/sharepointService';

interface CadastrosViewProps {
  configSharepoint: SharePointConfig;
  onConfigSharepointChanged: (c: SharePointConfig) => void;
  onRefresh: () => void;
}

export const CadastrosView: React.FC<CadastrosViewProps> = ({
  configSharepoint,
  onConfigSharepointChanged,
  onRefresh
}) => {
  const [categoriasEntrada, setCategoriasEntrada] = useState<string[]>([...CATEGORIAS_ENTRADA]);
  const [categoriasSaida, setCategoriasSaida] = useState<string[]>([...CATEGORIAS_SAIDA]);
  const [novaEntrada, setNovaEntrada] = useState('');
  const [novaSaida, setNovaSaida] = useState('');

  const [siteUrl, setSiteUrl] = useState(configSharepoint.siteUrl);
  const [listName, setListName] = useState(configSharepoint.listName);
  const [savedSuccess, setSavedSuccess] = useState(false);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const handleSalvarConfig = () => {
    const spService = SharePointService.getInstance();
    const updated = spService.updateConfig({
      siteUrl: siteUrl.trim() || 'https://appchurch.sharepoint.com/sites/Tesouraria',
      listName: listName.trim() || 'BD_Relatorio'
    });
    onConfigSharepointChanged(updated);
    setSavedSuccess(true);
    setTimeout(() => setSavedSuccess(false), 3000);
  };

  const handleLoginSharePoint = async () => {
    setIsLoggingIn(true);
    const spService = SharePointService.getInstance();
    await spService.loginSharePoint({
      email: 'admin@teste.com',
      nome: 'Admin',
      siteUrl: siteUrl.trim() || 'https://appchurch.sharepoint.com/sites/Tesouraria',
      listName: listName.trim() || 'BD_Relatorio'
    });
    onConfigSharepointChanged(spService.getConfig());
    onRefresh();
    setIsLoggingIn(false);
  };

  const handleRestaurarPadraoSharePoint = () => {
    const spService = SharePointService.getInstance();
    const padrao = spService.restaurarPadroesSharePoint();
    setSiteUrl(padrao.siteUrl);
    setListName(padrao.listName);
    onConfigSharepointChanged(padrao);
    onRefresh();
    setSavedSuccess(true);
    setTimeout(() => setSavedSuccess(false), 3000);
  };

  const handleAddEntrada = (e: React.FormEvent) => {
    e.preventDefault();
    if (!novaEntrada.trim()) return;
    setCategoriasEntrada(prev => [...prev, novaEntrada.trim()]);
    setNovaEntrada('');
  };

  const handleAddSaida = (e: React.FormEvent) => {
    e.preventDefault();
    if (!novaSaida.trim()) return;
    setCategoriasSaida(prev => [...prev, novaSaida.trim()]);
    setNovaSaida('');
  };

  const handleRemoveEntrada = (cat: string) => {
    setCategoriasEntrada(prev => prev.filter(c => c !== cat));
  };

  const handleRemoveSaida = (cat: string) => {
    setCategoriasSaida(prev => prev.filter(c => c !== cat));
  };

  return (
    <div id="cadastros-view-container" className="p-3.5 sm:p-6 space-y-4 sm:space-y-6 max-w-[1600px] mx-auto text-slate-100">
      <div className="border-b border-[#2d334d] pb-4">
        <h2 className="text-xl sm:text-2xl font-black text-white tracking-tight flex items-center gap-2">
          Cadastros & Integração SharePoint
        </h2>
        <p className="text-[11px] sm:text-xs text-slate-400">
          Gerenciamento de categorias financeiras de fluxo de caixa e parâmetros de conexão com o banco
        </p>
      </div>

      {/* SharePoint Integration Configuration Box */}
      <div className="bg-[#24293f] p-3.5 sm:p-5 rounded-xl border border-[#323955] space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#2e344e] pb-3">
          <div className="flex items-center gap-2">
            <Database className="w-5 h-5 text-indigo-400 shrink-0" />
            <h3 className="text-xs sm:text-sm font-bold text-white">
              Parâmetros do Banco de Dados SharePoint (Microsoft Graph)
            </h3>
          </div>
          <span className="text-xs font-semibold px-2.5 py-1 rounded bg-emerald-950/80 text-emerald-300 border border-emerald-500/30 flex items-center gap-1.5 shrink-0">
            <CheckCircle2 className="w-3.5 h-3.5" />
            {configSharepoint.status}
          </span>
        </div>

        {/* SharePoint Account Card */}
        <div className="bg-[#181b2a] p-3 sm:p-3.5 rounded-lg border border-[#2f354f] flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-full bg-indigo-950/80 border border-indigo-500/40 flex items-center justify-center shrink-0">
              <ShieldCheck className="w-5 h-5 text-indigo-400" />
            </div>
            <div>
              <p className="text-xs font-bold text-white flex flex-wrap items-center gap-1.5">
                <span>Conta Microsoft SharePoint:</span> 
                <span className="font-normal text-slate-300 break-all">
                  {configSharepoint.usuarioConectado?.email || 'Nenhuma conta conectada'}
                </span>
              </p>
              <p className="text-[11px] text-slate-400">
                {configSharepoint.usuarioConectado
                  ? `Logado como ${configSharepoint.usuarioConectado.nome} • Acesso concedido à lista ${configSharepoint.listName}`
                  : 'Faça login para permitir escrita e sincronização bidirecional'}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
            {configSharepoint.usuarioConectado ? (
              <button
                type="button"
                id="btn-cadastros-logout-sp"
                onClick={handleRestaurarPadraoSharePoint}
                className="px-3 py-1.5 rounded-lg bg-[#1a1d2e] hover:bg-[#282d46] text-indigo-300 border border-indigo-500/30 text-xs font-medium transition-colors cursor-pointer"
                title="Restaura a conexão padrão com Admin (teste) e BD_Relatorio"
              >
                Restaurar Padrões SharePoint
              </button>
            ) : (
              <button
                type="button"
                id="btn-cadastros-login-sp"
                onClick={handleLoginSharePoint}
                disabled={isLoggingIn}
                className="px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow transition-colors flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
              >
                {isLoggingIn ? (
                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Lock className="w-3.5 h-3.5" />
                )}
                Fazer Login do SharePoint
              </button>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4 text-xs">
          <div>
            <label className="block text-slate-300 font-medium mb-1">
              URL do Site SharePoint
            </label>
            <input
              type="text"
              value={siteUrl}
              onChange={(e) => setSiteUrl(e.target.value)}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-3 py-2 text-white focus:outline-none focus:border-indigo-400"
            />
          </div>

          <div>
            <label className="block text-slate-300 font-medium mb-1">
              Nome da Lista de Tesouraria
            </label>
            <input
              type="text"
              value={listName}
              onChange={(e) => setListName(e.target.value)}
              className="w-full bg-[#181b2a] border border-[#343b56] rounded-lg px-3 py-2 text-white focus:outline-none focus:border-indigo-400"
            />
          </div>
        </div>

        <div className="flex flex-wrap items-center justify-between gap-2 pt-2">
          <p className="text-[11px] sm:text-xs text-slate-400">
            Última sincronização: <strong>{configSharepoint.ultimaSincronizacao}</strong> ({configSharepoint.totalItensSincronizados} registros)
          </p>
          <div className="flex items-center gap-2">
            {savedSuccess && (
              <span className="text-xs text-emerald-400 font-semibold animate-in fade-in">
                Salvo com sucesso!
              </span>
            )}
            <button
              onClick={handleSalvarConfig}
              className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow transition-colors cursor-pointer"
            >
              Salvar Parâmetros
            </button>
          </div>
        </div>
      </div>

      {/* Categorias de Fluxo de Caixa (Modo Edição e Gestão) */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 sm:gap-6">
        {/* Categorias de Entrada */}
        <div className="bg-[#24293f] p-5 rounded-xl border border-[#323955] space-y-4">
          <h3 className="text-sm font-bold text-white flex items-center justify-between border-b border-[#2e344e] pb-2.5">
            <span className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-400" />
              Categorias de Entrada (+) Receitas
            </span>
            <span className="text-xs font-normal text-slate-400">{categoriasEntrada.length} cadastradas</span>
          </h3>

          {/* Form Adicionar Categoria Entrada */}
          <form onSubmit={handleAddEntrada} className="flex gap-2">
            <input
              type="text"
              value={novaEntrada}
              onChange={(e) => setNovaEntrada(e.target.value)}
              placeholder="Nova categoria de entrada..."
              className="flex-1 bg-[#181b2a] border border-[#2f354f] rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none focus:border-indigo-500"
            />
            <button
              type="submit"
              className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold shadow transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Adicionar</span>
            </button>
          </form>

          <ul className="space-y-1.5 text-xs max-h-80 overflow-y-auto pr-1">
            {categoriasEntrada.map((cat) => (
              <li key={cat} className="flex items-center justify-between bg-[#191d2e] px-3 py-2 rounded-lg border border-[#2b3149] hover:border-[#384060]">
                <span className="text-slate-200 font-medium">{cat}</span>
                <button
                  type="button"
                  onClick={() => handleRemoveEntrada(cat)}
                  className="text-slate-400 hover:text-red-400 p-1 transition-colors cursor-pointer"
                  title={`Remover categoria ${cat}`}
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </li>
            ))}
          </ul>
        </div>

        {/* Categorias de Saída */}
        <div className="bg-[#24293f] p-5 rounded-xl border border-[#323955] space-y-4">
          <h3 className="text-sm font-bold text-white flex items-center justify-between border-b border-[#2e344e] pb-2.5">
            <span className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-rose-400" />
              Categorias de Saída (-) Despesas
            </span>
            <span className="text-xs font-normal text-slate-400">{categoriasSaida.length} cadastradas</span>
          </h3>

          {/* Form Adicionar Categoria Saída */}
          <form onSubmit={handleAddSaida} className="flex gap-2">
            <input
              type="text"
              value={novaSaida}
              onChange={(e) => setNovaSaida(e.target.value)}
              placeholder="Nova categoria de despesa..."
              className="flex-1 bg-[#181b2a] border border-[#2f354f] rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none focus:border-indigo-500"
            />
            <button
              type="submit"
              className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold shadow transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Adicionar</span>
            </button>
          </form>

          <ul className="space-y-1.5 text-xs max-h-80 overflow-y-auto pr-1">
            {categoriasSaida.map((cat) => (
              <li key={cat} className="flex items-center justify-between bg-[#191d2e] px-3 py-2 rounded-lg border border-[#2b3149] hover:border-[#384060]">
                <span className="text-slate-200 font-medium">{cat}</span>
                <button
                  type="button"
                  onClick={() => handleRemoveSaida(cat)}
                  className="text-slate-400 hover:text-red-400 p-1 transition-colors cursor-pointer"
                  title={`Remover categoria ${cat}`}
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </li>
            ))}
          </ul>
        </div>
      </div>
    </div>
  );
};
