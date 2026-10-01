import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { LancamentoTesouraria, FiltrosFluxoCaixa, MovimentacaoFluxoCaixa } from '../types';

export class ExportService {
  /**
   * Formata número para moeda brasileira
   */
  public static formatMoeda(valor: number): string {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    }).format(valor);
  }

  /**
   * Exporta dados para arquivo Excel (.xlsx) com múltiplas abas estruturadas
   */
  public static exportarParaExcel(
    lancamentos: LancamentoTesouraria[],
    filtros: FiltrosFluxoCaixa,
    nomeArquivo: string = 'Relatorio_Fluxo_Caixa_ADM_AppChurch'
  ): void {
    const wb = XLSX.utils.book_new();

    // 1. Aba: Resumo Executivo
    const totalEntradas = lancamentos
      .filter(l => l.tipo === 'ENTRADA')
      .reduce((sum, l) => sum + l.valorTotal, 0);

    const totalSaidas = lancamentos
      .filter(l => l.tipo === 'SAIDA')
      .reduce((sum, l) => sum + l.valorTotal, 0);

    const totalPix = lancamentos.reduce((sum, l) => sum + (l.valorPix || 0), 0);
    const totalEspecie = lancamentos.reduce((sum, l) => sum + (l.valorEspecie || 0), 0);
    const saldoLiquido = totalEntradas - totalSaidas;

    const dadosResumo = [
      ['ADM TESOURARIA APPCHURCH - RELATÓRIO EXECUTIVO DE FLUXO DE CAIXA'],
      ['Data de Geração:', new Date().toLocaleString('pt-BR')],
      ['Filtro Ano:', filtros.ano || 'Todos'],
      ['Filtro Mês:', filtros.mes === 'todos' ? 'Todos os Meses' : `Mês ${filtros.mes}`],
      ['Filtro Setor:', filtros.setor || 'Todos'],
      ['Filtro Área:', filtros.area || 'Todas'],
      ['Filtro Período:', `${filtros.dataInicio || 'Início'} até ${filtros.dataFim || 'Hoje'}`],
      [''],
      ['INDICADOR FINANCEIRO', 'VALOR (R$)', 'OBSERVAÇÃO'],
      ['Total de Entradas', totalEntradas, 'Receitas de Célula, Dízimos e Ofertas'],
      ['Total de Saídas', totalSaidas, 'Despesas e Custos Operacionais'],
      ['Saldo Líquido / Operacional', saldoLiquido, saldoLiquido >= 0 ? 'Superávit' : 'Déficit'],
      ['Total Arrecadado via PIX', totalPix, 'Entradas digitais'],
      ['Total Arrecadado em Espécie', totalEspecie, 'Dinheiro físico em envelopes'],
      ['Quantidade Total de Lançamentos', lancamentos.length, 'Registros validados']
    ];

    const wsResumo = XLSX.utils.aoa_to_sheet(dadosResumo);
    XLSX.utils.book_append_sheet(wb, wsResumo, 'Resumo Executivo');

    // 2. Aba: Lançamentos Detalhados
    const dadosDetalhados = lancamentos.map(item => ({
      'ID Lançamento': item.id,
      'SharePoint ID': item.sharepointId || '-',
      'Data': item.data,
      'Ano': item.ano,
      'Mês': item.mes,
      'Semana': item.semanaNumero,
      'Setor': item.setor,
      'Área': item.area,
      'Célula / Beneficiário': item.celulaNome,
      'Tipo': item.tipo === 'ENTRADA' ? 'Entrada (+)' : 'Saída (-)',
      'Categoria': item.categoria,
      'Descrição': item.descricao,
      'Valor PIX (R$)': item.valorPix,
      'Valor Espécie (R$)': item.valorEspecie,
      'Valor Total (R$)': item.valorTotal,
      'Forma Pagamento': item.metodo,
      'Status': item.status,
      'Origem Dado': item.origem === 'SHAREPOINT_LIST' ? 'SharePoint DB' : 'Manual'
    }));

    const wsDetalhes = XLSX.utils.json_to_sheet(dadosDetalhados);
    XLSX.utils.book_append_sheet(wb, wsDetalhes, 'Lançamentos Detalhados');

    // 3. Aba: Consolidado por Setor
    const setorMap = new Map<string, { pix: number; especie: number; total: number; qtd: number }>();
    lancamentos.forEach(l => {
      const atual = setorMap.get(l.setor) || { pix: 0, especie: 0, total: 0, qtd: 0 };
      atual.pix += (l.valorPix || 0);
      atual.especie += (l.valorEspecie || 0);
      atual.total += l.valorTotal;
      atual.qtd += 1;
      setorMap.set(l.setor, atual);
    });

    const dadosSetor = Array.from(setorMap.entries()).map(([setor, vals]) => ({
      'Setor': setor,
      'Qtd Lançamentos': vals.qtd,
      'Total PIX (R$)': Number(vals.pix.toFixed(2)),
      'Total Espécie (R$)': Number(vals.especie.toFixed(2)),
      'Total Geral (R$)': Number(vals.total.toFixed(2)),
      '% do Total': totalEntradas > 0 ? `${((vals.total / totalEntradas) * 100).toFixed(1)}%` : '0%'
    }));

    const wsSetor = XLSX.utils.json_to_sheet(dadosSetor);
    XLSX.utils.book_append_sheet(wb, wsSetor, 'Consolidado por Setor');

    // Salva o arquivo
    const dataHora = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `${nomeArquivo}_${dataHora}.xlsx`);
  }

  /**
   * Exporta relatório oficial formatado em PDF
   */
  public static exportarParaPDF(
    lancamentos: LancamentoTesouraria[],
    filtros: FiltrosFluxoCaixa,
    tituloPersonalizado: string = 'Relatório de Fluxo de Caixa'
  ): void {
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    const totalEntradas = lancamentos
      .filter(l => l.tipo === 'ENTRADA')
      .reduce((sum, l) => sum + l.valorTotal, 0);

    const totalSaidas = lancamentos
      .filter(l => l.tipo === 'SAIDA')
      .reduce((sum, l) => sum + l.valorTotal, 0);

    const totalPix = lancamentos.reduce((sum, l) => sum + (l.valorPix || 0), 0);
    const totalEspecie = lancamentos.reduce((sum, l) => sum + (l.valorEspecie || 0), 0);
    const saldo = totalEntradas - totalSaidas;

    // Header Background Bar
    doc.setFillColor(28, 32, 48); // #1c2030
    doc.rect(0, 0, 297, 24, 'F');

    // Title
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('ADM TESOURARIA APPCHURCH', 14, 11);

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(200, 210, 230);
    doc.text(`${tituloPersonalizado.toUpperCase()} | LEITURA TOTAL DE DADOS`, 14, 18);

    // Generation timestamp & filters
    doc.setFontSize(8);
    doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, 230, 11);
    doc.text(`Origem: Banco de Dados SharePoint`, 230, 18);

    // KPI Boxes
    const kpiY = 28;
    const boxW = 52;
    const boxH = 16;

    const kpis = [
      { label: 'TOTAL ENTRADAS', val: this.formatMoeda(totalEntradas), color: [34, 197, 94] },
      { label: 'TOTAL SAÍDAS', val: this.formatMoeda(totalSaidas), color: [239, 68, 68] },
      { label: 'SALDO LÍQUIDO', val: this.formatMoeda(saldo), color: saldo >= 0 ? [14, 165, 233] : [239, 68, 68] },
      { label: 'TOTAL PIX', val: this.formatMoeda(totalPix), color: [100, 116, 139] },
      { label: 'TOTAL ESPÉCIE', val: this.formatMoeda(totalEspecie), color: [100, 116, 139] }
    ];

    kpis.forEach((kpi, idx) => {
      const bx = 14 + idx * (boxW + 4);
      doc.setFillColor(245, 247, 250);
      doc.roundedRect(bx, kpiY, boxW, boxH, 2, 2, 'F');
      doc.setDrawColor(220, 226, 235);
      doc.roundedRect(bx, kpiY, boxW, boxH, 2, 2, 'S');

      doc.setFontSize(7);
      doc.setTextColor(100, 116, 139);
      doc.setFont('helvetica', 'bold');
      doc.text(kpi.label, bx + 4, kpiY + 5);

      doc.setFontSize(11);
      doc.setTextColor(kpi.color[0], kpi.color[1], kpi.color[2]);
      doc.text(kpi.val, bx + 4, kpiY + 12);
    });

    // Subtitle filters applied
    doc.setFontSize(8);
    doc.setTextColor(71, 85, 105);
    doc.setFont('helvetica', 'italic');
    doc.text(
      `Filtros: Ano [${filtros.ano || 'Todos'}] | Mês [${filtros.mes === 'todos' ? 'Todos' : filtros.mes}] | Setor [${filtros.setor}] | Área [${filtros.area}] | Registros exibidos: ${lancamentos.length}`,
      14,
      48
    );

    // Table Data
    const tableRows = lancamentos.map(l => [
      l.data,
      l.setor,
      l.area,
      l.celulaNome,
      l.categoria,
      l.tipo === 'ENTRADA' ? '+ Entrada' : '- Saída',
      l.valorPix > 0 ? this.formatMoeda(l.valorPix) : '-',
      l.valorEspecie > 0 ? this.formatMoeda(l.valorEspecie) : '-',
      this.formatMoeda(l.valorTotal),
      l.status === 'CONFIRMADO' ? 'OK' : 'Pendente'
    ]);

    autoTable(doc, {
      startY: 52,
      head: [['Data', 'Setor', 'Área', 'Célula / Origem', 'Categoria', 'Tipo', 'PIX', 'Espécie', 'Total', 'Status']],
      body: tableRows,
      theme: 'grid',
      headStyles: {
        fillColor: [30, 41, 59],
        textColor: [255, 255, 255],
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'left'
      },
      styles: {
        fontSize: 7.5,
        cellPadding: 2,
        lineColor: [226, 232, 240],
        lineWidth: 0.2
      },
      columnStyles: {
        0: { cellWidth: 20 },
        1: { cellWidth: 22 },
        2: { cellWidth: 24 },
        3: { cellWidth: 38 },
        4: { cellWidth: 42 },
        5: { cellWidth: 20 },
        6: { cellWidth: 26, halign: 'right' },
        7: { cellWidth: 26, halign: 'right' },
        8: { cellWidth: 28, halign: 'right', fontStyle: 'bold' },
        9: { cellWidth: 20, halign: 'center' }
      },
      didDrawPage: (data) => {
        // Footer with page numbering
        const pageCount = (doc as any).internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(148, 163, 184);
        doc.text(
          `ADM Tesouraria AppChurch - Página ${data.pageNumber} de ${pageCount}`,
          14,
          200
        );

        // Signatures line on bottom
        doc.setDrawColor(180, 190, 205);
        doc.line(160, 198, 220, 198);
        doc.line(230, 198, 285, 198);
        doc.setFontSize(6.5);
        doc.text('Admin - Tesouraria', 165, 202);
        doc.text('Pastor Presidente / Conselho', 235, 202);
      }
    });

    const dataHora = new Date().toISOString().slice(0, 10);
    doc.save(`Fluxo_Caixa_AppChurch_${dataHora}.pdf`);
  }

  /**
   * Exporta a lista de movimentações de BD_FluxoCaixa para Excel
   */
  public static exportarFluxoCaixaExcel(
    movimentacoes: MovimentacaoFluxoCaixa[],
    info: { ano?: number; mes?: string },
    nomeArquivo: string = 'Fluxo_Caixa_AppChurch'
  ): void {
    const wb = XLSX.utils.book_new();

    const totalEntradas = movimentacoes
      .filter(m => m.tipo === 'ENTRADA')
      .reduce((sum, m) => sum + m.valor, 0);

    const totalSaidas = movimentacoes
      .filter(m => m.tipo === 'SAIDA')
      .reduce((sum, m) => sum + m.valor, 0);

    const saldo = totalEntradas - totalSaidas;

    const resumo = [
      ['APPCHURCH - DEMONSTRATIVO DE FLUXO DE CAIXA'],
      ['Data de Exportação:', new Date().toLocaleString('pt-BR')],
      ['Ano de Referência:', info.ano || 'Todos'],
      ['Mês de Referência:', info.mes || 'Todos'],
      [''],
      ['INDICADOR', 'VALOR (R$)', 'STATUS'],
      ['Total de Entradas', totalEntradas, 'Receitas'],
      ['Total de Saídas', totalSaidas, 'Despesas'],
      ['Saldo Líquido', saldo, saldo >= 0 ? 'Superávit' : 'Déficit'],
      ['Total de Movimentações', movimentacoes.length, 'Registros']
    ];

    const wsResumo = XLSX.utils.aoa_to_sheet(resumo);
    XLSX.utils.book_append_sheet(wb, wsResumo, 'Resumo');

    const detalhado = movimentacoes.map(m => ({
      'ID': m.id,
      'Data': m.dataBR || m.data,
      'Ano': m.ano,
      'Mês': m.mes,
      'Tipo': m.tipo === 'ENTRADA' ? 'Entrada (+)' : 'Saída (-)',
      'Descrição': m.descricao,
      'Categoria': m.categoria,
      'Forma de Pagamento': m.formaPagamento || 'PIX',
      'Valor (R$)': m.valor,
      'Status': m.status || 'Confirmado',
      'Observações': m.observacao || ''
    }));

    const wsDetalhes = XLSX.utils.json_to_sheet(detalhado);
    XLSX.utils.book_append_sheet(wb, wsDetalhes, 'Movimentações');

    XLSX.writeFile(wb, `${nomeArquivo}.xlsx`);
  }

  /**
   * Exporta a lista de movimentações de BD_FluxoCaixa para PDF
   */
  public static exportarFluxoCaixaPDF(
    movimentacoes: MovimentacaoFluxoCaixa[],
    info: { ano?: number; mes?: string },
    nomeArquivo: string = 'Fluxo_Caixa_AppChurch'
  ): void {
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    const totalEntradas = movimentacoes
      .filter(m => m.tipo === 'ENTRADA')
      .reduce((sum, m) => sum + m.valor, 0);

    const totalSaidas = movimentacoes
      .filter(m => m.tipo === 'SAIDA')
      .reduce((sum, m) => sum + m.valor, 0);

    const saldo = totalEntradas - totalSaidas;

    // Cabeçalho
    doc.setFillColor(24, 28, 43);
    doc.rect(0, 0, 297, 24, 'F');

    doc.setFontSize(14);
    doc.setTextColor(255, 255, 255);
    doc.text('APPCHURCH - DEMONSTRATIVO DE FLUXO DE CAIXA', 14, 15);

    // Cards resumo
    doc.setFontSize(10);
    doc.setTextColor(30, 41, 59);
    doc.text(`Ano: ${info.ano || 'Todos'}  |  Entradas: ${this.formatMoeda(totalEntradas)}  |  Saídas: ${this.formatMoeda(totalSaidas)}  |  Saldo: ${this.formatMoeda(saldo)}`, 14, 34);

    const tableRows = movimentacoes.map(m => [
      m.dataBR || m.data,
      m.tipo === 'ENTRADA' ? '+ Entrada' : '- Saída',
      m.descricao,
      m.categoria,
      m.formaPagamento || 'PIX',
      this.formatMoeda(m.valor),
      m.status || 'Confirmado'
    ]);

    autoTable(doc, {
      startY: 40,
      head: [['Data', 'Tipo', 'Descrição / Histórico', 'Categoria', 'Pagamento', 'Valor (R$)', 'Status']],
      body: tableRows,
      theme: 'grid',
      headStyles: {
        fillColor: [30, 41, 59],
        textColor: [255, 255, 255],
        fontSize: 8,
        fontStyle: 'bold'
      },
      styles: {
        fontSize: 8,
        cellPadding: 2
      },
      columnStyles: {
        0: { cellWidth: 24 },
        1: { cellWidth: 22 },
        2: { cellWidth: 70 },
        3: { cellWidth: 50 },
        4: { cellWidth: 28 },
        5: { cellWidth: 35, halign: 'right', fontStyle: 'bold' },
        6: { cellWidth: 25, halign: 'center' }
      }
    });

    doc.save(`${nomeArquivo}.pdf`);
  }
}
