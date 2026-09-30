export type ViewMode = 
  | 'fluxo-caixa'
  | 'dashboard'
  | 'relacao-envelopes'
  | 'validar-relatorios'
  | 'indicador-trilho'
  | 'menu-admin'
  | 'login';

export type SetorTipo = 
  | 'Laranja'
  | 'Amarelo 1'
  | 'Amarelo 2'
  | 'Amarelo 3'
  | 'Amarelo 4'
  | 'Roxo 1'
  | 'Roxo 2'
  | 'Roxo 3'
  | 'Verde'
  | 'Verde 2';

export type AreaTipo = 
  | 'Área Central'
  | 'Área Norte'
  | 'Área Sul'
  | 'Área Leste'
  | 'Área Oeste';

export type TipoTransacao = 'ENTRADA' | 'SAIDA';

export type MetodoPagamento = 'PIX' | 'ESPECIE' | 'TRANSFERENCIA' | 'BOLETO';

export interface BDRelatorioItem {
  ID: number | string;
  DataCelula: string;
  dataBR?: string;
  ValorOferta: number;
  OfertaEspecie: number;
  Célula: string;
  Membros?: number;
  Convidados?: number;
  Criancas?: number;
  Total?: number;
  Setor: string;
  Area: string;
  LíderCelula?: string;
  LíderSetor?: string;
  NumSemana?: number;
  MembrosPresentes?: number;
  Supervisao?: boolean;
  ID_CELULA?: string | number;
  SETOR_RECEB?: boolean | string;
  TESOURARIA_RECEB: boolean;
  DATA_TESOURARIA?: string;
  ID_TESOUREIRO?: string | number;
  MES?: number | string;
  Criado?: string;
}

export interface CelulaItem {
  id: string;
  nome: string;
  lider: string;
  setor: SetorTipo;
  area: AreaTipo;
  diaSemana: string;
  horario: string;
  ativo: boolean;
}

export interface LancamentoTesouraria {
  id: string;
  sharepointId?: number;
  data: string; // YYYY-MM-DD
  dataBR?: string; // DD/MM/YYYY
  semanaNumero: number;
  ano: number;
  mes: number; // 1-12
  celulaId?: string;
  celulaNome: string;
  setor: SetorTipo;
  area: AreaTipo;
  tipo: TipoTransacao;
  categoria: string;
  descricao: string;
  valorPix: number;
  valorEspecie: number;
  valorTotal: number;
  metodo: MetodoPagamento;
  status: 'PENDENTE' | 'CONFIRMADO' | 'CANCELADO';
  origem: 'SHAREPOINT_LIST' | 'MANUAL_APP' | 'IMPORT_EXCEL';
  dataSincronizacao?: string;
  observacoes?: string;

  // Campos canônicos da lista BD_Relatorio do SharePoint
  ID?: number | string;
  DataCelula?: string;
  ValorOferta?: number;
  OfertaEspecie?: number;
  Célula?: string;
  Membros?: number;
  Convidados?: number;
  Criancas?: number;
  Total?: number;
  Setor?: string;
  Area?: string;
  LíderCelula?: string;
  LíderSetor?: string;
  NumSemana?: number;
  MembrosPresentes?: number;
  Supervisao?: boolean;
  ID_CELULA?: string | number;
  SETOR_RECEB?: boolean | string;
  TESOURARIA_RECEB: boolean; // Coluna essencial: só é somado no total se true!
  DATA_TESOURARIA?: string;
  ID_TESOUREIRO?: string | number;
  MES?: number | string;
  Criado?: string;
}

export interface MembroItem {
  id: string | number;
  ID?: string | number;
  Title?: string;
  nome: string;
  login: string;
  email?: string;
  senha?: string;
  cargo?: string;
  celula?: string;
  setor?: string;
  area?: string;
  telefone?: string;
  status?: 'Ativo' | 'Inativo';
}

export interface SharePointAuthUser {
  nome: string;
  email: string;
  cargo: string;
  avatarUrl?: string;
  conectadoEm: string;
}

export interface SharePointConfig {
  siteUrl: string;
  listName: string;
  listId: string;
  status: 'CONECTADO' | 'DESCONECTADO' | 'SINCRONIZANDO' | 'ERRO';
  ultimaSincronizacao: string;
  totalItensSincronizados: number;
  autoSync: boolean;
  tenantId?: string;
  clientId?: string;
  usuarioConectado?: SharePointAuthUser | null;
  authMethod?: 'MICROSOFT_ACCOUNT' | 'APP_SECRET' | 'CUSTOM_API';
}

export interface FiltrosFluxoCaixa {
  periodoPredefinido: 'hoje' | 'esta_semana' | 'este_mes' | 'mes_anterior' | 'trimestre' | 'ano_2026' | 'ano_2025' | 'todos' | 'personalizado';
  dataInicio: string;
  dataFim: string;
  ano: number;
  mes: number | 'todos';
  setor: SetorTipo | 'todos';
  area: AreaTipo | 'todos';
  celula: string | 'todos';
  tipo: 'todos' | 'ENTRADA' | 'SAIDA';
  categoria: string | 'todos';
  metodo: 'todos' | MetodoPagamento;
  status: 'todos' | 'CONFIRMADO' | 'PENDENTE';
}

export type VisualizacaoAgrupamento = 'periodo' | 'setor' | 'area' | 'mensal' | 'anual' | 'categoria' | 'tabela-total';

export interface EtapaStats {
  etapa: string;
  concluidos: number;
  pendentes: number;
  percentual: number;
}

export interface MetricasTrilhoConjunto {
  totalMembros: number;
  percentualMedio: number;
  membrosCompletos: number;
  membrosNaoCompletos?: number;
  membrosComZero?: number;
  membrosEmAndamento: number;
  membrosNaoIniciados: number;
  etapasStats: EtapaStats[];
}

export interface MembroTrilhoItem {
  id: string | number;
  nome: string;
  login: string;
  setor: string;
  area: string;
  cargo: string;
  celula: string;
  etapasConcluidas: string[];
  totalConcluidas: number;
  totalEtapas: number;
  percentualConclusao: number;
  statusTrilho: 'Completo' | 'Em Andamento' | 'Não Iniciado';
}

export interface IndicadorTrilhoResponse {
  sucesso: boolean;
  totalMembrosGeral: number;
  totalCapacitacoesRegistros: number;
  etapas: string[];
  areasDisponiveis: string[];
  setoresDisponiveis: string[];
  geral: MetricasTrilhoConjunto;
  porArea: Record<string, MetricasTrilhoConjunto>;
  porSetor: Record<string, MetricasTrilhoConjunto>;
  membros: MembroTrilhoItem[];
}
