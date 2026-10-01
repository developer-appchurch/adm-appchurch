import React, { useState } from 'react';
import { 
  X, 
  PlusCircle, 
  TrendingUp, 
  TrendingDown, 
  DollarSign, 
  Calendar, 
  UserCheck, 
  CheckCircle2, 
  Clock, 
  FileText,
  AlertCircle,
  RefreshCw
} from 'lucide-react';
import { MembroItem, MovimentacaoFluxoCaixa } from '../../types';

interface NovoFluxoModalProps {
  isOpen: boolean;
  onClose: () => void;
  usuarioLogado?: MembroItem | null;
  onSalvoComSucesso: (novoFluxo: MovimentacaoFluxoCaixa) => void;
}

export const NovoFluxoModal: React.FC<NovoFluxoModalProps> = ({
  isOpen,
  onClose,
  usuarioLogado,
  onSalvoComSucesso
}) => {
  const [categoriaFluxo, setCategoriaFluxo] = useState<'Entrada' | 'Saída'>('Entrada');
  const [tipoFluxo, setTipoFluxo] = useState<'Pix' | 'Espécie'>('Pix');
  const [valorInput, setValorInput] = useState<string>('');
  const [dataFluxo, setDataFluxo] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [statusFluxo, setStatusFluxo] = useState<'OK' | 'Pendente'>('OK');
  const [descricaoFluxo, setDescricaoFluxo] = useState<string>('');
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [erroForm, setErroForm] = useState<string | null>(null);

  if (!isOpen) return null;

  // Formatador e manipulador de máscara monetária BRL
  const handleValorChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const rawDigits = e.target.value.replace(/\D/g, '');
    if (!rawDigits) {
      setValorInput('');
      return;
    }
    const num = Number(rawDigits) / 100;
    const formatted = new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(num);
    setValorInput(formatted);
  };

  const getValorNumerico = (): number => {
    const rawDigits = valorInput.replace(/\D/g, '');
    return rawDigits ? Number(rawDigits) / 100 : 0;
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErroForm(null);

    const valorFinal = getValorNumerico();
    if (valorFinal <= 0) {
      setErroForm('Por favor, digite um valor maior que zero em R$.');
      return;
    }

    if (!dataFluxo) {
      setErroForm('Por favor, selecione a data do fluxo de caixa.');
      return;
    }

    if (!descricaoFluxo.trim()) {
      setErroForm('Por favor, informe a descrição (destino / motivo) do fluxo de caixa.');
      return;
    }

    const idTesoureiroFinal = usuarioLogado?.id || usuarioLogado?.ID || 4;

    setIsSubmitting(true);
    try {
      const payload = {
        CategoriaFluxo: categoriaFluxo,
        TipoFluxo: tipoFluxo,
        ValorFluxo: valorFinal,
        DataFluxo: dataFluxo,
        Id_Tesoureiro: idTesoureiroFinal,
        StatusFluxo: statusFluxo,
        DescricaoFluxo: descricaoFluxo.trim()
      };

      const res = await fetch('/api/sharepoint/fluxo-caixa', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const data = await res.json().catch(() => null);

      if (!res.ok || !data?.sucesso) {
        throw new Error(data?.erro || 'Erro ao gravar na tabela BD_FluxoCaixa.');
      }

      // Sucesso
      onSalvoComSucesso(data.item);
      onClose();
      
      // Reseta campos
      setValorInput('');
      setDescricaoFluxo('');
      setCategoriaFluxo('Entrada');
      setTipoFluxo('Pix');
      setStatusFluxo('OK');
    } catch (err: any) {
      console.error('[NovoFluxoModal] Erro ao salvar:', err);
      setErroForm(err?.message || 'Falha ao salvar no banco de dados.');
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-black/75 backdrop-blur-sm animate-in fade-in duration-200">
      <div 
        id="modal-novo-fluxo"
        className="bg-[#181c2b] border border-[#2e3752] rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden flex flex-col max-h-[92vh] animate-in zoom-in-95 duration-150 text-slate-100"
      >
        {/* Modal Header */}
        <div className="px-5 py-4 border-b border-[#262d44] flex items-center justify-between bg-[#151825]">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-indigo-600/20 border border-indigo-500/30 flex items-center justify-center text-indigo-400">
              <PlusCircle className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold text-white leading-tight">
                Novo Fluxo de Caixa
              </h2>
              <p className="text-[11px] text-slate-400">
                Cadastro de entrada e saída financeira
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-[#23293e] transition-colors cursor-pointer"
            aria-label="Fechar"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Form */}
        <form onSubmit={handleSubmit} className="p-5 space-y-4 overflow-y-auto flex-1 text-xs">
          
          {erroForm && (
            <div className="p-3 rounded-xl bg-rose-950/50 border border-rose-500/60 text-rose-300 flex items-center gap-2 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0 text-rose-400" />
              <span>{erroForm}</span>
            </div>
          )}

          {/* 1. Escolha da Categoria (Entrada / Saída) */}
          <div>
            <label className="block text-slate-300 font-bold mb-1.5">
              Categoria do Fluxo
            </label>
            <div className="grid grid-cols-2 gap-2.5">
              <button
                type="button"
                onClick={() => setCategoriaFluxo('Entrada')}
                className={`py-2.5 px-3 rounded-xl font-bold flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  categoriaFluxo === 'Entrada'
                    ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300 shadow-md shadow-emerald-950/50'
                    : 'bg-[#1e2336] border-[#313955] text-slate-400 hover:text-slate-200'
                }`}
              >
                <TrendingUp className="w-4 h-4 text-emerald-400" />
                <span>Entrada (+)</span>
              </button>

              <button
                type="button"
                onClick={() => setCategoriaFluxo('Saída')}
                className={`py-2.5 px-3 rounded-xl font-bold flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  categoriaFluxo === 'Saída'
                    ? 'bg-rose-950/80 border-rose-500 text-rose-300 shadow-md shadow-rose-950/50'
                    : 'bg-[#1e2336] border-[#313955] text-slate-400 hover:text-slate-200'
                }`}
              >
                <TrendingDown className="w-4 h-4 text-rose-400" />
                <span>Saída (-)</span>
              </button>
            </div>
          </div>

          {/* 2. Escolha do Tipo do Fluxo (Pix / Espécie) */}
          <div>
            <label className="block text-slate-300 font-bold mb-1.5">
              Forma de Pagamento
            </label>
            <div className="grid grid-cols-2 gap-2.5">
              <button
                type="button"
                onClick={() => setTipoFluxo('Pix')}
                className={`py-2.5 px-3 rounded-xl font-bold flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  tipoFluxo === 'Pix'
                    ? 'bg-indigo-950/80 border-indigo-500 text-indigo-300 shadow-md shadow-indigo-950/50'
                    : 'bg-[#1e2336] border-[#313955] text-slate-400 hover:text-slate-200'
                }`}
              >
                <span>PIX / Transferência</span>
              </button>

              <button
                type="button"
                onClick={() => setTipoFluxo('Espécie')}
                className={`py-2.5 px-3 rounded-xl font-bold flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  tipoFluxo === 'Espécie'
                    ? 'bg-amber-950/80 border-amber-500 text-amber-300 shadow-md shadow-amber-950/50'
                    : 'bg-[#1e2336] border-[#313955] text-slate-400 hover:text-slate-200'
                }`}
              >
                <span>Espécie (Dinheiro)</span>
              </button>
            </div>
          </div>

          {/* 3. Valor da Movimentação & Data */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="block text-slate-300 font-bold mb-1">
                Valor da Movimentação *
              </label>
              <div className="relative">
                <DollarSign className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  required
                  value={valorInput}
                  onChange={handleValorChange}
                  placeholder="R$ 0,00"
                  className="w-full bg-[#121522] border border-[#313955] focus:border-indigo-500 rounded-xl pl-9 pr-3 py-2.5 text-sm font-mono font-bold text-white placeholder-slate-500 focus:outline-none"
                />
              </div>
            </div>

            <div>
              <label className="block text-slate-300 font-bold mb-1">
                Data da Movimentação *
              </label>
              <div className="relative">
                <Calendar className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="date"
                  required
                  value={dataFluxo}
                  onChange={(e) => setDataFluxo(e.target.value)}
                  className="w-full bg-[#121522] border border-[#313955] focus:border-indigo-500 rounded-xl pl-9 pr-3 py-2.5 text-xs text-white focus:outline-none cursor-pointer"
                />
              </div>
            </div>
          </div>

          {/* 4. Status do Fluxo & Tesoureiro Responsável */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="block text-slate-300 font-bold mb-1">
                Status
              </label>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => setStatusFluxo('OK')}
                  className={`py-2 px-2.5 rounded-xl font-bold flex items-center justify-center gap-1.5 border transition-all cursor-pointer ${
                    statusFluxo === 'OK'
                      ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300'
                      : 'bg-[#1e2336] border-[#313955] text-slate-400'
                  }`}
                >
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                  <span>OK</span>
                </button>

                <button
                  type="button"
                  onClick={() => setStatusFluxo('Pendente')}
                  className={`py-2 px-2.5 rounded-xl font-bold flex items-center justify-center gap-1.5 border transition-all cursor-pointer ${
                    statusFluxo === 'Pendente'
                      ? 'bg-amber-950/80 border-amber-500 text-amber-300'
                      : 'bg-[#1e2336] border-[#313955] text-slate-400'
                  }`}
                >
                  <Clock className="w-3.5 h-3.5 text-amber-400" />
                  <span>Pendente</span>
                </button>
              </div>
            </div>

            <div>
              <label className="block text-slate-300 font-bold mb-1">
                Tesoureiro Responsável
              </label>
              <div className="flex items-center gap-2 bg-[#121522] border border-[#313955] px-3 py-2 rounded-xl text-slate-300">
                <UserCheck className="w-4 h-4 text-indigo-400 shrink-0" />
                <div className="truncate">
                  <span className="font-bold text-white text-[11px]">
                    ID: {usuarioLogado?.id || usuarioLogado?.ID || 4}
                  </span>
                  <span className="text-[10px] text-slate-400 block truncate">
                    {usuarioLogado?.nome || 'Admin Tesouraria'}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* 5. Descrição do Fluxo (Destino / Motivo) */}
          <div>
            <label className="block text-slate-300 font-bold mb-1">
              Descrição (Destino / Motivo) *
            </label>
            <textarea
              required
              rows={3}
              value={descricaoFluxo}
              onChange={(e) => setDescricaoFluxo(e.target.value)}
              placeholder="Ex: Oferta de Domingo - Culto da Noite / Pagamento de conta de energia da nave principal"
              className="w-full bg-[#121522] border border-[#313955] focus:border-indigo-500 rounded-xl px-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-none resize-none"
            />
          </div>

          {/* Footer Actions */}
          <div className="pt-3 border-t border-[#262d44] flex items-center justify-end gap-2.5">
            <button
              type="button"
              onClick={onClose}
              disabled={isSubmitting}
              className="px-4 py-2 rounded-xl bg-[#202538] hover:bg-[#2c334d] text-slate-300 font-bold text-xs transition-colors cursor-pointer border border-[#303954]"
            >
              Cancelar
            </button>

            <button
              type="submit"
              disabled={isSubmitting}
              className="px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs shadow-lg shadow-indigo-600/30 transition-all flex items-center gap-2 cursor-pointer disabled:opacity-60"
            >
              {isSubmitting ? (
                <>
                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                  <span>Salvando Movimentação...</span>
                </>
              ) : (
                <>
                  <PlusCircle className="w-4 h-4" />
                  <span>Cadastrar Fluxo</span>
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
