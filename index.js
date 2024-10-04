import { saveAs } from 'file-saver';
import * as ExcelJS from 'exceljs';

// Funções auxiliares para formatação
const formatDate = (date) => {
  // Preserva a data original como string no formato "YYYY-MM-DD"
  const [year, month, day] = date.split('-');
  return `${day}/${month}/${year}`; // Retorna no formato DD/MM/YYYY
};

const formatCurrency = (value) => {
  return `R$ ${value.toFixed(2).replace('.', ',')}`;
};

const button = document.querySelector('.button');
button.addEventListener('click', function () {
  generateExcet();
  // RunExcelJSExport();
});

async function generateExcet() {
  const data = {
    filename: 'relatorio_fluxo_caixa_diario',
    title: 'Relatório de fluxo de caixa diário',
    columns: ['Data', 'Recebimentos', 'Pagamentos', 'Saldo final'],
    rows: [
      ['2024-10-01', 'Pagamento de cartão de crédito', 'R$ 100,00'],
      ['2024-10-02', 'Pagamento de fatura', 'R$ 200,00'],
      ['2024-10-03', 'Recebimento de transferência bancária', 'R$ 300,00'],
    ],
  };

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Relatório Financeiro');

  const _addHeaderRow = (worksheet, options) => {
    const row = worksheet.addRow([options.title]);
    worksheet.mergeCells(row.number, 1, row.number, options.columns.length);

    row.getCell(1).alignment = {
      horizontal: 'center',
      vertical: 'middle',
    };

    row.height = 30;
    row.font = {
      bold: true,
      color: { argb: '000000' },
      size: 16,
      name: 'Arial',
    };

    return row;
  };

  const _addTitlesRow = (worksheet, options) => {
    const row = worksheet.addRow(options.columns);
    row.eachCell(
      (cell) => (cell.alignment = { horizontal: 'center', vertical: 'middle' })
    );
    row.height = 20;
    row.font = {
      bold: true,
      color: { argb: '000000' },
      size: 10,
      name: 'Arial',
    };

    return row;
  };

  const _addContentRow = (worksheet, options) => {
    options.rows.forEach((r) => {
      const row = worksheet.addRow(r);
      row.eachCell(
        (cell) =>
          (cell.alignment = { horizontal: 'center', vertical: 'middle' })
      );
      row.height = 25;
      row.font = {
        bold: true,
        color: { argb: '000000' },
        size: 12,
        name: 'Arial',
      };
    });

    return row;
  };

  const headerRow = _addHeaderRow(worksheet, data);
  const titlesRow = _addTitlesRow(worksheet, data);
  const contentRow = _addContentRow(worksheet, data);

  // Gerar o arquivo Excel e baixá-lo
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), 'relatorio_financeiro_agrupado.xlsx');
}

async function RunExcelJSExport() {
  const data = [
    {
      enabled: true,
      deleted: false,
      date: '2024-10-01',
      receivements: 0,
      payments: 0,
      balance: 0,
      details: [],
    },
    {
      enabled: true,
      deleted: false,
      date: '2024-10-02',
      receivements: 100,
      payments: 0,
      balance: 100,
      details: [
        {
          amountPaid: 100,
          paymentHour: '14:41:02.2',
          paymentDate: '2024-10-02',
          bankAccountDescription: 'BOLETO - CAIXA ECONOMICA FEDERAL',
          chargeType: 'CREDIT_CARD',
          documentNumber: '00001/BO1-01',
          movementType: 'CARD_ADMINISTRATOR_RECEIVEMENT',
          operation: 'CREDIT',
        },
      ],
    },
    {
      enabled: true,
      deleted: false,
      date: '2024-10-03',
      receivements: 0,
      payments: 100,
      balance: 100,
      details: [],
    },
    {
      enabled: true,
      deleted: false,
      date: '2024-08-04',
      receivements: 14620.73,
      payments: 0,
      balance: 14620.73,
      details: [
        {
          amountPaid: 14620.73,
          paymentHour: '07:23:56.767',
          paymentDate: '2024-08-05',
          bankAccountDescription: 'BOLETO - BANCO DO BRASIL S.A.',
          chargeType: 'CHECK',
          documentNumber: '1',
          movementType: 'CHECK_RECEIVED',
          operation: 'CREDIT',
        },
        {
          amountPaid: 10,
          paymentHour: '15:25:25.895',
          paymentDate: '2024-08-02',
          bankAccountDescription: 'BOLETO - BANCO DO BRASIL S.A.',
          chargeType: 'CHECK',
          documentNumber: '100',
          movementType: 'BILL_INSTALLMENT_PAYABLE',
          operation: 'DEBIT',
        },
      ],
    },
  ];

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Relatório Financeiro');

  // Cabeçalhos
  worksheet.addRow(['Data', 'Recebimentos', 'Pagamentos', 'Saldo final']);

  for (const entry of data) {
    // Linha principal (sem outline)
    const parentRow = worksheet.addRow([
      formatDate(entry.date),
      formatCurrency(entry.receivements),
      formatCurrency(entry.payments),
      formatCurrency(entry.balance),
    ]);

    parentRow.height = 20;

    // Linhas de receitas (outline level 1)
    if (entry.details.some((e) => e.operation === 'CREDIT')) {
      // Adicionando a linha de título "Receitas"
      const receivementsTitle = worksheet.addRow(['Receitas']);
      receivementsTitle.outlineLevel = 1;

      // Mesclando células da linha de título para "Receitas" e aplicando estilo
      worksheet.mergeCells(
        receivementsTitle.number,
        1,
        receivementsTitle.number,
        6
      );
      receivementsTitle.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'e9f4e8' }, // Cor de fundo #e9f4e8
      };
      receivementsTitle.getCell(1).border = {
        top: { style: 'none', color: { argb: '000000' } }, // Borda superior preta
        left: { style: 'thin', color: { argb: '000000' } }, // Borda esquerda preta
        bottom: { style: 'thin', color: { argb: '000000' } }, // Borda inferior preta
        right: { style: 'thin', color: { argb: '000000' } }, // Borda direita preta
      };
      receivementsTitle.getCell(1).alignment = {
        horizontal: 'center',
        vertical: 'middle',
      };

      // Adicionando os detalhes dos recebimentos
      entry.details
        .filter((e) => e.operation === 'CREDIT')
        .forEach((detail, index) => {
          const detailRow = worksheet.addRow([
            detail.paymentHour.substring(0, 5),
            detail.bankAccountDescription,
            formatCurrency(detail.amountPaid),
            detail.movementType,
            detail.documentNumber,
          ]);
          detailRow.outlineLevel = 2; // Definir outlineLevel para as linhas detalhadas

          detailRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'e9f4e8' }, // Cor de fundo #e9f4e8
          };
        });
    }

    // Linhas de despesas (outline level 1)
    if (entry.details.some((e) => e.operation === 'DEBIT')) {
      // Adicionando a linha de título "Despesas"
      const paymentsTitle = worksheet.addRow(['Despesas']);
      paymentsTitle.outlineLevel = 1;

      // Mesclando células da linha de título para "Despesas" e aplicando estilo
      worksheet.mergeCells(paymentsTitle.number, 1, paymentsTitle.number, 6);
      paymentsTitle.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'f3e7e7' }, // Cor de fundo #f3e7e7
      };
      paymentsTitle.getCell(1).alignment = {
        horizontal: 'center',
        vertical: 'middle',
      };

      // Adicionando os detalhes dos pagamentos
      entry.details
        .filter((e) => e.operation === 'DEBIT')
        .forEach((detail) => {
          const detailRow = worksheet.addRow([
            detail.paymentHour.substring(0, 5),
            detail.bankAccountDescription,
            formatCurrency(detail.amountPaid),
            detail.movementType,
            detail.documentNumber,
          ]);
          detailRow.outlineLevel = 2; // Definir outlineLevel para as linhas detalhadas
          detailRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'f3e7e7' }, // Cor de fundo #f3e7e7
          };
        });
    }
  }

  // Habilitar a visualização de agrupamento no Excel
  worksheet.properties.outlineLevelRow = 1;

  // Gerar o arquivo Excel e baixá-lo
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), 'relatorio_financeiro_agrupado.xlsx');
}
