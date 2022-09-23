const _ = require('lodash');
const Excel = require('exceljs');

const numeral = require('numeral');
const moment = require('moment');

const TRANSACTION_TYPE = {
  PURCHASE: 'Purchase',
  REFUND: 'Refund'
};

const getFormattedMoney = value => {
  const precisionFormat = _.repeat('0', 1);
  const scaleFormat = _.repeat('0', 2);

  return numeral(value).format(`${precisionFormat}.${scaleFormat}`);
};

const getFormattedDate = value => {
  const format = 'YYYY-MM-DD HH:mm:ss Z';

  try {
    return moment(value).locale('id').utcOffset(7).format(format);
  } catch (error) {
    throw new Error(`Invalid Indonesia date of ${value}`);
  }
};

const findTransactionWithTerminalId = transactions => {
  return _.find(
    transactions,
    transactionDetail => {
      const terminalId = _.get(transactionDetail, 'terminalId');

      return !_.isNil(terminalId);
    }
  );
};

const addBorder = column => {
  column.eachCell(cell => {
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
};

const getFixedColumns = () => {
  return [
    { header: 'NO', key: 'no', width: 5 },
    { header: 'MERCHANT NAME', key: 'merchantName', width: 25 },
    { header: 'TRANSACTION DATE', key: 'transactionDate', width: 25 },
    { header: 'TRANSIDMERCHANT', key: 'transactionId', width: 25 },
    { header: 'CUSTOMER NAME', key: 'customerName', width: 25 },
    { header: 'AMOUNT', key: 'amount', width: 25 },
    { header: 'FEE', key: 'fee', width: 25 },
    { header: 'TAX', key: 'feeTax', width: 25 },
    { header: 'MERCHANT SUPPORT', key: 'merchantSupport', width: 25 },
    { header: 'PAY TO MERCHANT', key: 'payToMerchant', width: 25 },
    { header: 'PAY OUT DATE', key: 'payoutDate', width: 25 },
    { header: 'TRANSACTION TYPE', key: 'transactionType', width: 25 },
    { header: 'TENURE', key: 'transactionLoanTenure', width: 25 },
  ];
};

const getAdditionalColumns = options => {
  let additionalColumns = [];

  if (_.get(options, 'transactionType', '') === TRANSACTION_TYPE.REFUND) {
    additionalColumns.push({ header: 'REFUND_IDS_SEPARATED_BY_SEMICOLON', key: 'refundIds', width: 30 });
  }

  if (_.get(options, 'hasTransactionWithTerminalId', false)) {
    additionalColumns.push({ header: 'TERMINAL_ID', key: 'terminalId', width: 30 });
  }

  return additionalColumns;
};

const createTransactionColumns = (options) => {
  return _.concat(
    getFixedColumns(),
    getAdditionalColumns(options)
  );
};

/**
 * param 'options' is used to add extra column on a specified row that contains certain info,
 *
 *
 * @param Object options: {
 *   transactionType
 *   hasTransactionWithTerminalId // this property is used only by YouTap merchants for their internal reconciliation purposes
 * }
 */

const mapTransactionDetailToRow = ({ transactionDetail, index, options }) => {
  let transactionDetailRow = {
    no: index + 1,
    merchantName: transactionDetail.merchantName,
    // because there are merchants that have parent-child relationship
    // we need to use merchantName to precisely display the child merchant's name where the transaction took place
    transactionId: transactionDetail.transactionId,
    transactionDate: getFormattedDate(transactionDetail.transactionDate),
    customerName: transactionDetail.customerName,
    amount: getFormattedMoney(transactionDetail.amount),
    fee: getFormattedMoney(transactionDetail.fee),
    feeTax: getFormattedMoney(_.get(transactionDetail, 'feeTax', 0)),
    merchantSupport: getFormattedMoney(_.get(transactionDetail, 'merchantSupport', 0)),
    payToMerchant: getFormattedMoney(transactionDetail.payToMerchant),
    payoutDate: getFormattedDate(transactionDetail.payoutDate),
    transactionType: transactionDetail.transactionType,
    transactionLoanTenure: transactionDetail.transactionLoanTenure,
  };

  if (_.get(options, 'transactionType', '') === TRANSACTION_TYPE.REFUND) {
    _.assign(transactionDetailRow, {
      refundIds: _.join(transactionDetail.refundIds, ';')
    });
  }

  if (_.get(options, 'hasTransactionWithTerminalId', false)) {
    _.assign(transactionDetailRow, {
      terminalId: _.get(transactionDetail, 'terminalId')
    });
  }

  return transactionDetailRow;
};

const addWorksheet = (workbook, {
  worksheetName,
  transactionDetails,
  options
}) => {
  const worksheet = workbook.addWorksheet(worksheetName);

  const headerRowNumber = 1;

  // Table header font set to bold
  worksheet.getRow(headerRowNumber).font = { bold: true };

  const generatedColumns = createTransactionColumns(options);
  worksheet.columns = generatedColumns;

  transactionDetails.forEach((transactionDetail, index) => {
    const row = mapTransactionDetailToRow({
      transactionDetail,
      index,
      options
    });

    worksheet.addRow(row);
  });

  _.forEach(generatedColumns, column => {
    // Add border for all cell on table
    addBorder(worksheet.getColumn(column.key));
  });

  return workbook;
};

const constructMerchantSettlementReport = ({
  purchaseDetails,
  refundDetails
}) => {
  const workbook = new Excel.Workbook();

  const purchaseTransactionDetails = _.map(purchaseDetails, purchaseDetail => ({
    ...purchaseDetail,
    transactionType: TRANSACTION_TYPE.PURCHASE
  }));

  addWorksheet(workbook, {
    worksheetName: 'Transaction', // A little weird name, but that's how it is
    transactionDetails: purchaseTransactionDetails,
    options: {
      hasTransactionWithTerminalId: !_.isNil(findTransactionWithTerminalId(purchaseTransactionDetails)),
      transactionType: TRANSACTION_TYPE.PURCHASE
    }
  });

  const refundTransactionDetails = _.map(refundDetails, refundDetail => ({
    ...refundDetail,
    transactionType: TRANSACTION_TYPE.REFUND
  }));

  addWorksheet(workbook, {
    worksheetName: 'Refund',
    transactionDetails: refundTransactionDetails,
    options: {
      hasTransactionWithTerminalId: !_.isNil(findTransactionWithTerminalId(refundTransactionDetails)),
      transactionType: TRANSACTION_TYPE.REFUND
    }
  });

  const disbursementLedgers = _
    .chain([
      ...purchaseTransactionDetails,
      ...refundTransactionDetails,
    ])
    .sortBy('transactionDate')
    .value();

  addWorksheet(workbook, {
    worksheetName: 'Ledger',
    transactionDetails: disbursementLedgers,
    options: {
      hasTransactionWithTerminalId: false // for ledgers, we DO NOT need to display terminal_id
    }
  });

  return workbook;
};

const getExcelAttachment = async ({
  workbook,
  filename
}) => {
  const fileBuffer = await workbook.xlsx.writeBuffer();
  const fileBase64 = fileBuffer.toString('base64');

  return {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    name: `${filename}.xlsx`,
    content: fileBase64
  };
};

const processor = () => {
  const source = {}; // Payment summary json content goes here

  const { purchaseReport, refundReport } = source;

  return {
    // To get merchant name, email, report date & disbursement amount; using data from any reports between these two should be fine
    // However, using purchaseReport is preferable because there might not be any refund and there should be at least 1 purchase
    merchantName: purchaseReport.merchantName,
    merchantEmail: purchaseReport.merchantEmail,
    reportDate: purchaseReport.date,
    disbursedAmount: purchaseReport.merchantDisbursementAmount,
    purchaseDetails: _.get(purchaseReport, 'merchantDisbursementDetails'),
    totalPurchaseCount: _.get(purchaseReport, 'merchantDisbursementDetails').length,
    totalPurchaseAmount: _.get(purchaseReport, 'merchantDisbursementDetailsTotalAmount'),
    refundDetails: _.get(refundReport, 'merchantDisbursementDetails'),
    totalRefundCount: _.get(refundReport, 'merchantDisbursementDetails').length,
    totalRefundAmount: _.get(refundReport, 'merchantDisbursementDetailsTotalAmount'),
    // isRequiredToNotifyFinanceTeam: _.toLower(process.env.NODE_ENV) === 'prod' &&
    //   _.includes(MERCHANT_SLUGS_TO_NOTIFY_FINANCE_TEAM, purchaseReport.merchantSlug),
    // product: PRODUCT.INDODANA_PAYLATER,
    // useCase: USE_CASE.MERCHANT_DISBURSEMENT_COMPLETED,
  };
};

(async (res) => {
  const {
    reportDate,
    purchaseDetails,
    refundDetails,
  } = processor();

  const formattedReportDate = moment(reportDate).format('DD-MM-YYYY');
  const type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const fileName = `Indodana Disbursement Report ${formattedReportDate}.xlsx`;
  const workbook = constructMerchantSettlementReport({
    purchaseDetails,
    refundDetails
  });

  console.log('res ==> ', res);
  //
  // res.setHeader('Content-Type', type);
  // res.setHeader('Content-Disposition', 'attachment; filename=' + `${fileName}`);

  await workbook.xlsx.writeFile(fileName);
  console.log('File write done........');
})();
