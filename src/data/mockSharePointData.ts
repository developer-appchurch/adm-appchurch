import { CelulaItem, LancamentoTesouraria, SetorTipo, AreaTipo, SharePointConfig, BDRelatorioItem, MembroItem } from '../types';

export const SETORES_DISPONIVEIS: SetorTipo[] = [
  'Amarelo 1',
  'Amarelo 2',
  'Amarelo 3',
  'Amarelo 4',
  'Laranja',
  'Roxo 1',
  'Roxo 2',
  'Roxo 3',
  'Verde',
  'Verde 2'
];

export const AREAS_DISPONIVEIS: AreaTipo[] = [
  'Área Central',
  'Área Norte',
  'Área Sul',
  'Área Leste',
  'Área Oeste'
];

export const CATEGORIAS_ENTRADA = [
  'Oferta de Célula (PIX)',
  'Oferta de Célula (Espécie)',
  'Dízimos dos Membros',
  'Ofertas de Culto Geral',
  'Doações Especiais / Missões',
  'Inscrições de Eventos e Retiros'
];

export const CATEGORIAS_SAIDA = [
  'Manutenção Predial e Reformas',
  'Material Didático e de Célula',
  'Energia Elétrica e Água',
  'Ajuda Ministerial e Pastoral',
  'Equipamentos de Som e Mídia',
  'Ação Social e Cestas Básicas',
  'Aluguel de Espaços e Templos',
  'Sistemas e Licenças de TI'
];

export const CELULAS_INICIAIS: CelulaItem[] = [
  // Células canônicas para referência de cadastro e supervisão
  { id: 'cel-1', nome: 'Alto Monte', lider: 'Carlos Eduardo', setor: 'Laranja', area: 'Área Central', diaSemana: 'Quarta-feira', horario: '20:00', ativo: true },
  { id: 'cel-2', nome: 'Refúgio', lider: 'Mariana Silva', setor: 'Laranja', area: 'Área Central', diaSemana: 'Quinta-feira', horario: '19:30', ativo: true },
  { id: 'cel-3', nome: 'Vencedores em Cristo', lider: 'Pr. André Souza', setor: 'Laranja', area: 'Área Central', diaSemana: 'Terça-feira', horario: '20:00', ativo: true },
  { id: 'cel-4', nome: 'Jeová Jireh', lider: 'Felipe Santos', setor: 'Laranja', area: 'Área Central', diaSemana: 'Quarta-feira', horario: '20:00', ativo: true },
  { id: 'cel-5', nome: 'Figueira', lider: 'Renata Oliveira', setor: 'Laranja', area: 'Área Central', diaSemana: 'Sexta-feira', horario: '20:00', ativo: true },
  { id: 'cel-6', nome: 'Leão de Judá', lider: 'Marcos Vinícius', setor: 'Amarelo 1', area: 'Área Norte', diaSemana: 'Quinta-feira', horario: '20:00', ativo: true },
  { id: 'cel-7', nome: 'Raiz de Davi', lider: 'Priscila Rocha', setor: 'Amarelo 1', area: 'Área Norte', diaSemana: 'Quarta-feira', horario: '19:30', ativo: true },
  { id: 'cel-8', nome: 'Aleluia', lider: 'Rodrigo Lima', setor: 'Amarelo 1', area: 'Área Norte', diaSemana: 'Terça-feira', horario: '20:00', ativo: true },
  { id: 'cel-9', nome: 'Chamas da Fé', lider: 'Beatriz Costa', setor: 'Amarelo 2', area: 'Área Norte', diaSemana: 'Sexta-feira', horario: '20:00', ativo: true },
  { id: 'cel-10', nome: 'Monte Sião', lider: 'Thiago Mendes', setor: 'Amarelo 2', area: 'Área Norte', diaSemana: 'Quarta-feira', horario: '19:45', ativo: true }
];

/**
 * Membros iniciais sincronizados com a lista BD_membros:
 * https://pazchurch.sharepoint.com/sites/PazSobral/Lists/BD_membros/
 */
export const MEMBROS_INICIAIS_BD_MEMBROS: MembroItem[] = [
  {
    id: 4,
    ID: 4,
    Title: 'Junio Fonteles',
    nome: 'Junio Fonteles',
    login: 'Jfonteles',
    email: 'juniosina@hotmail.com',
    senha: '12345',
    cargo: 'Líder de Setor',
    celula: 'Central',
    setor: 'Safira',
    area: 'Área Central',
    telefone: '(88) 99999-0004',
    status: 'Ativo'
  },
  {
    id: 1,
    ID: 1,
    Title: 'Admin Tesouraria',
    nome: 'Admin Tesouraria',
    login: 'admin',
    email: 'midia.sobral@paz.church',
    senha: 'teste',
    cargo: 'Tesoureiro',
    celula: 'Central',
    setor: 'Geral',
    area: 'Área Central',
    telefone: '(88) 99999-0001',
    status: 'Ativo'
  },
  {
    id: 2,
    ID: 2,
    Title: 'Developer AppChurch',
    nome: 'Developer AppChurch',
    login: 'developer.appchurch@gmail.com',
    email: 'developer.appchurch@gmail.com',
    senha: 'teste',
    cargo: 'Tesoureiro',
    celula: 'Central',
    setor: 'Safira',
    area: 'Área Central',
    telefone: '(88) 99999-0002',
    status: 'Ativo'
  },
  {
    id: 513,
    ID: 513,
    Title: 'Ana Karina',
    nome: 'Ana Karina',
    login: 'Karina',
    email: 'karinacerssobral@gmail.com',
    senha: 'karina123',
    cargo: 'Membro',
    celula: 'Central',
    setor: 'Safira',
    area: 'Área Central',
    telefone: '(88) 99999-0513',
    status: 'Ativo'
  },
  {
    id: 705,
    ID: 705,
    Title: 'Eluana Priscila',
    nome: 'Eluana Priscila',
    login: 'eluana',
    email: 'e.sobral@hotmail.com',
    senha: '123',
    cargo: 'Líder de Área',
    celula: 'Central',
    setor: 'Safira',
    area: 'Área Central',
    telefone: '(88) 99999-0705',
    status: 'Ativo'
  }
];

export const URL_SHAREPOINT_BD_MEMBROS = 'https://pazchurch.sharepoint.com/sites/PazSobral/Lists/BD_membros/';

/**
 * Converte item bruto retornado da lista BD_membros do SharePoint
 */
export function converterItemSharepointParaMembro(raw: any): MembroItem {
  const fields = raw.fields || raw;
  const id = fields.ID || fields.Id || fields.id || Math.random().toString(36).substr(2, 6);
  const title = String(fields.Title || fields.Nome || fields.nome || fields.Login || fields.login || '').trim();
  const nome = String(fields.Nome || fields.nome || fields.NomeCompleto || fields.Title || 'Membro').trim();
  const email = fields.Email || fields.email || fields.EMail || fields.Mail || '';
  
  // Deriva login caso não exista coluna específica
  let login = String(fields.Login || fields.login || fields.Usuario || fields.usuario || '').trim();
  if (!login) {
    if (email) {
      login = email.split('@')[0];
    } else {
      login = nome.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '.');
    }
  }

  return {
    id,
    ID: id,
    Title: title,
    nome,
    login,
    email,
    senha: fields.Senha || fields.senha,
    cargo: fields.Cargo || fields.cargo || fields.Funcao || fields.funcao || 'Membro',
    celula: fields.Celula || fields.celula || fields.Célula || '',
    setor: fields.Setor || fields.setor || '',
    area: fields.Area || fields.area || '',
    telefone: fields.Telefone || fields.telefone || fields.Celular || '',
    status: (fields.Status === 'Inativo' || fields.Ativo === false) ? 'Inativo' : 'Ativo'
  };
}

export const CONFIG_SHAREPOINT_PADRAO: SharePointConfig = {
  siteUrl: 'https://appchurch.sharepoint.com/sites/Tesouraria',
  listName: 'BD_Relatorio', // Nome exato da lista SharePoint fornecido pelo usuário
  listId: 'bd-relatorio-sharepoint-list-id',
  status: 'CONECTADO',
  ultimaSincronizacao: 'Aguardando sincronização da lista BD_Relatorio',
  totalItensSincronizados: 0,
  autoSync: true,
  tenantId: 'appchurch.onmicrosoft.com',
  clientId: '00000003-0000-0ff1-ce00-000000000000',
  authMethod: 'MICROSOFT_ACCOUNT',
  usuarioConectado: {
    nome: 'Admin',
    email: 'admin@teste.com',
    cargo: 'Administrador Geral',
    conectadoEm: new Date().toISOString().replace('T', ' ').slice(0, 19)
  }
};

/**
 * Normaliza um item vindo da lista BD_Relatorio do SharePoint
 * Colunas do BD_Relatorio:
 * ID, DataCelula, ValorOferta, OfertaEspecie, Célula, Membros, Convidados, Criancas, Total, Setor, Area,
 * LíderCelula, LíderSetor, NumSemana, MembrosPresentes, Supervisao(boolean), ID_CELULA, SETOR_RECEB,
 * TESOURARIA_RECEB (BOOLEAN), DATA_TESOURARIA, ID_TESOUREIRO, MES, Criado
 */
export function converterItemSharepointParaLancamento(raw: any): LancamentoTesouraria {
  // Lida com variações de maiúsculas/minúsculas e SharePoint REST API fields
  const fields = raw.fields || raw;

  const id = String(fields.ID || fields.Id || fields.id || `sp-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`);
  
  // Data da célula: verifica DataNascimento (formato DD/MM/YYYY do SharePoint), DataCelula, Data, Created
  let dataCelula = '';
  let dataBR = '';
  let ano = 0;
  let mes = 0;

  if (fields.DataNascimento && typeof fields.DataNascimento === 'string' && fields.DataNascimento.includes('/')) {
    const parts = fields.DataNascimento.trim().split('/');
    if (parts.length === 3) {
      const d = parts[0].padStart(2, '0');
      const m = parts[1].padStart(2, '0');
      const y = parts[2];
      dataCelula = `${y}-${m}-${d}`;
      dataBR = `${d}/${m}/${y}`;
      ano = parseInt(y, 10);
      mes = parseInt(m, 10);
    }
  }

  if (!dataCelula) {
    let rawData = fields.DataCelula || fields.dataCelula || fields.Data || fields.data || fields.Created || '';
    if (rawData && typeof rawData === 'string' && rawData.includes('T')) {
      rawData = rawData.split('T')[0];
    }
    dataCelula = rawData || new Date().toISOString().slice(0, 10);
    const parts = dataCelula.split('-');
    if (parts.length >= 1) ano = parseInt(parts[0], 10);
    if (parts.length >= 2) mes = parseInt(parts[1], 10);
    if (parts.length === 3) dataBR = `${parts[2]}/${parts[1]}/${parts[0]}`;
  }

  // Ano canônico baseado na data de criação do registro no SharePoint se não definido
  if (!ano && fields.Created && typeof fields.Created === 'string' && fields.Created.length >= 4) {
    ano = parseInt(fields.Created.slice(0, 4), 10);
  }
  if (!ano || isNaN(ano)) {
    ano = new Date().getFullYear();
  }

  // Valores de oferta:
  // IMPORTANTE: Na lista BD_Relatorio do SharePoint, a oferta via PIX é armazenada na coluna 'Bairro'
  // (ou 'ValorOferta'/'valorPix'), e a oferta em Espécie na coluna 'OfertaEspecie'!
  let valorPix = 0;
  if (fields.Bairro !== undefined && fields.Bairro !== null && String(fields.Bairro).trim() !== '' && !isNaN(Number(fields.Bairro))) {
    valorPix = Number(fields.Bairro);
  } else if (fields.ValorOferta !== undefined && fields.ValorOferta !== null) {
    valorPix = Number(fields.ValorOferta) || 0;
  } else if (fields.valorPix !== undefined) {
    valorPix = Number(fields.valorPix) || 0;
  }

  const valorEspecie = Number(fields.OfertaEspecie ?? fields.ofertaEspecie ?? fields.valorEspecie ?? 0) || 0;
  const valorTotal = Number((valorPix + valorEspecie).toFixed(2));

  // Célula, Liderança e Localização (trata encoding do SharePoint C_x00e9_lula e Nome)
  const celulaNome = String(fields.C_x00e9_lula || fields.Célula || fields.Celula || fields.celula || fields.celulaNome || 'Célula Sem Nome');
  const setor = (fields.Setor || fields.setor || 'Safira') as SetorTipo;
  const area = (fields.OData__x00c1_rea || fields.Area || fields.area || 'Área Central') as AreaTipo;
  const liderCelula = fields.Nome || fields.LíderCelula || fields.L_x00ed_derCelula || fields.LiderCelula || fields.liderCelula || '';
  const liderSetor = fields.LiderSetor || fields.LíderSetor || fields.liderSetor || '';

  // Semana e Mês
  let semanaNumero = 0;
  if (typeof fields.NumSemana === 'string') {
    const m = fields.NumSemana.match(/\d+/);
    semanaNumero = m ? parseInt(m[0], 10) : 0;
  } else if (fields.NumSemana !== undefined) {
    semanaNumero = Number(fields.NumSemana);
  }
  if (!mes) {
    mes = Number(fields.MES ?? fields.Mes ?? fields.mes ?? 0);
  }
  if (!mes && dataCelula) {
    const parts = dataCelula.split('-');
    if (parts.length >= 2) mes = parseInt(parts[1], 10);
  }

  // Frequência
  const membros = Number(fields.Membros ?? fields.membros ?? 0);
  const convidados = Number(fields.Convidados ?? fields.convidados ?? 0);
  const criancas = Number(fields.Criancas ?? fields.criancas ?? 0);
  const membrosPresentes = Number(fields.MembrosPresentes ?? fields.membrosPresentes ?? 0);
  const supervisao = Boolean(fields.Supervisao ?? fields.supervisao ?? false);
  const idCelula = fields.ID_CELULA || fields.id_celula || fields.idCelula || '';
  const setorReceb = fields.SETOR_RECEB ?? fields.setor_receb ?? fields.setorReceb ?? false;

  // TESOURARIA_RECEB (BOOLEAN) - Coluna essencial do usuário!
  const rawTesouraria = fields.TESOURARIA_RECEB ?? fields.tesouraria_receb ?? fields.tesourariaReceb ?? fields.status;
  const tesourariaReceb = (
    rawTesouraria === true ||
    rawTesouraria === 1 ||
    rawTesouraria === '1' ||
    String(rawTesouraria).trim().toLowerCase() === 'true' ||
    String(rawTesouraria).trim().toLowerCase() === 'sim' ||
    String(rawTesouraria).trim().toLowerCase() === 'yes' ||
    String(rawTesouraria).trim().toUpperCase() === 'CONFIRMADO'
  );

  const dataTesouraria = fields.DATA_TESOURARIA || fields.data_tesouraria || fields.dataTesouraria || (tesourariaReceb ? dataCelula : undefined);
  const idTesoureiro = fields.ID_TESOUREIRO || fields.id_tesoureiro || fields.idTesoureiro;
  const criado = fields.Criado || fields.Created || fields.criado || fields.dataSincronizacao;

  return {
    id,
    sharepointId: typeof fields.ID === 'number' ? fields.ID : parseInt(id, 10) || undefined,
    data: dataCelula,
    dataBR: dataBR || (dataCelula ? dataCelula.split('-').reverse().join('/') : ''),
    semanaNumero: semanaNumero || 38,
    ano: ano || 2026,
    mes: mes || 9,
    celulaNome,
    setor,
    area,
    tipo: 'ENTRADA',
    categoria: 'Oferta de Célula',
    descricao: `Envelope Célula ${celulaNome} - Sem. ${semanaNumero || ''}`,
    valorPix,
    valorEspecie,
    valorTotal,
    metodo: valorPix > 0 && valorEspecie === 0 ? 'PIX' : (valorEspecie > 0 && valorPix === 0 ? 'ESPECIE' : 'PIX'),
    status: tesourariaReceb ? 'CONFIRMADO' : 'PENDENTE',
    origem: 'SHAREPOINT_LIST',
    dataSincronizacao: new Date().toISOString().replace('T', ' ').slice(0, 19),

    // Campos canônicos do BD_Relatorio
    ID: fields.ID || id,
    DataCelula: dataCelula,
    ValorOferta: valorPix,
    OfertaEspecie: valorEspecie,
    Célula: celulaNome,
    Membros: membros,
    Convidados: convidados,
    Criancas: criancas,
    Total: valorTotal,
    Setor: setor,
    Area: area,
    LíderCelula: liderCelula,
    LíderSetor: liderSetor,
    NumSemana: semanaNumero,
    MembrosPresentes: membrosPresentes,
    Supervisao: supervisao,
    ID_CELULA: idCelula,
    SETOR_RECEB: setorReceb,
    TESOURARIA_RECEB: tesourariaReceb,
    DATA_TESOURARIA: dataTesouraria,
    ID_TESOUREIRO: idTesoureiro,
    MES: mes,
    Criado: criado
  };
}

/**
 * Utilitário universal para extrair listas de itens a partir de qualquer formato retornado
 * pelo Microsoft SharePoint (JSON REST API, Graph API, PowerAutomate, exportação CSV ou texto direto).
 */
export function parseSharePointPayload(input: any): any[] {
  if (!input) return [];

  // Se já for array
  if (Array.isArray(input)) {
    return input;
  }

  // Se for objeto estruturado
  if (typeof input === 'object') {
    if (Array.isArray(input.value)) return input.value;
    if (Array.isArray(input.d?.results)) return input.d.results;
    if (Array.isArray(input.d)) return input.d;
    if (Array.isArray(input.items)) return input.items;
    if (Array.isArray(input.data)) return input.data;
    if (Array.isArray(input.rows)) return input.rows;
    if (Array.isArray(input.records)) return input.records;
    if (Array.isArray(input.body?.value)) return input.body.value;
    if (Array.isArray(input.body)) return input.body;

    // Se for um único registro da lista (ex: { ID: 1, DataCelula: '...' })
    if (input.ID || input.Id || input.DataCelula || input.Célula || input.Total) {
      return [input];
    }
  }

  // Se for string (texto de arquivo ou textarea)
  if (typeof input === 'string') {
    const limpo = input.trim().replace(/^\uFEFF/, ''); // Remove UTF-8 BOM se existir

    // 1. Tenta parsear como JSON direto
    try {
      const parsed = JSON.parse(limpo);
      return parseSharePointPayload(parsed);
    } catch {
      // Continua para tentar outras estratégias
    }

    // 2. Tenta parsear como NDJSON (linhas com JSON em cada uma)
    const linhas = limpo.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (linhas.length > 0 && linhas[0].startsWith('{') && linhas[0].endsWith('}')) {
      const itens: any[] = [];
      for (const linha of linhas) {
        try {
          itens.push(JSON.parse(linha));
        } catch {
          // ignora linha malformada
        }
      }
      if (itens.length > 0) return itens;
    }

    // 3. Tenta parsear como CSV / TSV (separado por vírgula ou ponto-e-vírgula)
    if (linhas.length >= 2) {
      const separador = linhas[0].includes(';') ? ';' : (linhas[0].includes('\t') ? '\t' : ',');
      const cabecalhos = linhas[0].split(separador).map(h => h.trim().replace(/^["']|["']$/g, ''));
      
      const itensCsv: any[] = [];
      for (let i = 1; i < linhas.length; i++) {
        const valores = linhas[i].split(separador).map(v => v.trim().replace(/^["']|["']$/g, ''));
        if (valores.length >= 2) {
          const obj: Record<string, any> = {};
          cabecalhos.forEach((cab, idx) => {
            obj[cab] = valores[idx] ?? '';
          });
          itensCsv.push(obj);
        }
      }
      if (itensCsv.length > 0) return itensCsv;
    }
  }

  return [];
}

/**
 * Retorna lista vazia. Nenhum dado mocado padrão é criado,
 * permitindo que a aplicação pegue SOMENTE pelo SharePoint BD_Relatorio.
 */
export function gerarLancamentosIniciais(): LancamentoTesouraria[] {
  return [];
}
