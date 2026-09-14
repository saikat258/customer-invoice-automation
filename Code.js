const TEMPLATE_DOC_ID = '1wgtEoL6feqtZtb_lwdraJIibupCJXiHoM3coXOESCvc';
const SHEET_NAME = 'master_customer_details';

// Set either the folder ID or the folder path you want PDFs saved into.
// Prefer folder ID because it is stable and does not depend on a Drive path.
const INVOICE_FOLDER_ID = '1YE5isKtLep9AlqJS2ZZijGv4gaM2Jbbh';
const INVOICE_FOLDER_PATH = 'Invoices/Customer Invoices';

const PRODUCT_MAP = {
  'Panchmukhi HanumanJi': { itemCode: 'PHJ01', unitPrice: 1200 },
  'Little KanhaJi': { itemCode: 'LKC01', unitPrice: 900 },
  'Customized Solo Figurine': { itemCode: 'CSF01', unitPrice: 1200 },
  'Customized Couple Figurine': { itemCode: 'CCF01', unitPrice: 1700 }
};

function normalize(text) {
  return String(text || '')
    .trim()
    .toLowerCase()
    .replace(/[^\w\s]/g, '')   // remove punctuation (so 'Payment Completed?' -> 'payment completed')
    .replace(/\s+/g, ' ')
    .trim();
}

function getHeaderIndex(headers, candidate) {
  const target = normalize(candidate);
  return headers.findIndex(h => normalize(h) === target);
}

function getColumnNumber(headers, candidates) {
  for (const candidate of candidates) {
    const idx = getHeaderIndex(headers, candidate);
    if (idx !== -1) return idx + 1;
  }
  return 0;
}

function getValue(rowValues, headers, candidates) {
  for (const candidate of candidates) {
    const idx = getHeaderIndex(headers, candidate);
    if (idx !== -1) return rowValues[idx];
  }
  return '';
}

function ensureColumn(sheet, headers, headerName) {
  const idx = getHeaderIndex(headers, headerName);
  if (idx !== -1) return idx + 1;
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(headerName);
  headers.push(headerName);
  return newCol;
}

function buildEmailBody(customerName, invoiceAttached, failureMessage) {
  if (invoiceAttached) {
    return `Hello ${customerName},\n\nThank you for placing an order with us.\n\nPlease find your invoice attached.\n\nThank you.\n\n*** This is a system generated mail. Please contact us on @inkfinitystudio.3d on Instagram or mail us at payel.majumder.asn@gmail.com`;
  } else {
    return `Hello ${customerName},\n\n${failureMessage}\n\n*** This is a system generated mail. Please contact us on @inkfinitystudio.3d on Instagram or mail us at payel.majumder.asn@gmail.com`;
  }
}

function getInvoiceFolder() {
  if (INVOICE_FOLDER_ID) {
    return DriveApp.getFolderById(INVOICE_FOLDER_ID);
  }

  const pathParts = INVOICE_FOLDER_PATH.split('/').map(part => part.trim()).filter(Boolean);
  let currentFolder = DriveApp.getRootFolder();

  for (const part of pathParts) {
    let folder = null;
    const folders = currentFolder.getFoldersByName(part);
    while (folders.hasNext()) {
      folder = folders.next();
      break;
    }

    if (!folder) {
      folder = currentFolder.createFolder(part);
    }

    currentFolder = folder;
  }

  return currentFolder;
}

function createInvoicePdf(invoiceNumber, data) {
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  const invoiceFolder = getInvoiceFolder();
  const copyFile = templateFile.makeCopy(`Invoice ${invoiceNumber}`, invoiceFolder);
  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();

  body.replaceText('{{invoiceNumber}}', invoiceNumber);
  body.replaceText('{{customerName}}', data.customerName);
  body.replaceText('{{date}}', data.date);
  body.replaceText('{{email}}', data.email);
  body.replaceText('{{amount}}', data.formattedAmount);
  body.replaceText('{{billingAddress}}', data.billingAddress);
  body.replaceText('{{description}}', data.product);
  body.replaceText('{{qty}}', String(data.quantity));
  body.replaceText('{{unitPrice}}', data.formattedUnitPrice);
  body.replaceText('{{total}}', data.formattedAmount);
  body.replaceText('{{itemCode}}', data.itemCode);

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copyFile.getId()).getAs(MimeType.PDF);
  const pdfFile = invoiceFolder.createFile(pdfBlob).setName(`Invoice ${invoiceNumber}.pdf`);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  copyFile.setTrashed(true);

  return pdfFile;
}

function sendEmail(email, subject, body, pdfFile) {
  const options = { name: 'InkFinity Studios' };
  if (pdfFile) options.attachments = [pdfFile.getAs(MimeType.PDF)];
  GmailApp.sendEmail(email, subject, body, options);
}

function getBillingAddressValue(rowValues, headers) {
  const address1 = String(getValue(rowValues, headers, ['Billing Address (with Pincode):', 'Address', 'Customer Address', 'Address Line 1', 'Address 1']) || '').trim();
  const address2 = String(getValue(rowValues, headers, ['Address Line 2', 'Address 2', 'Apartment', 'Flat', 'Street Address']) || '').trim();

  if (!address1 && !address2) return '';
  if (address1 && address2) return `${address1}, ${address2}`;
  return address1 || address2;
}

function buildInvoiceData(rowValues, headers) {
  const product = String(getValue(rowValues, headers, ['Product'])).trim();
  const qtyRaw = String(getValue(rowValues, headers, ['Quantity', 'Qty']) || '1').trim();
  let quantity = parseInt(qtyRaw, 10);
  if (isNaN(quantity) || quantity < 1) quantity = 1;

  const productInfo = PRODUCT_MAP[product] || { itemCode: '', unitPrice: 0 };
  const amount = productInfo.unitPrice * quantity;

  return {
    customerName: String(getValue(rowValues, headers, ['Customer Name']) || 'Customer').trim(),
    email: String(getValue(rowValues, headers, ['Email'])).trim(),
    billingAddress: getBillingAddressValue(rowValues, headers),
    product,
    quantity,
    itemCode: productInfo.itemCode,
    unitPrice: productInfo.unitPrice,
    formattedUnitPrice: productInfo.unitPrice.toFixed(2),
    amount,
    formattedAmount: amount.toFixed(2),
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

/**
 * Core row processing used by both onFormSubmit and onEdit(payment)
 * If paymentYes = true -> generate PDF + send invoice
 * If paymentYes = false -> send failure message (no PDF)
 */
function processRowPayment(sheet, rowNumber, paymentYes) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = buildInvoiceData(rowValues, headers);

  if (!data.email) {
    Logger.log(`Skipping row ${rowNumber}: email empty`);
    return;
  }

  const invoiceNumberCol = getColumnNumber(headers, ['Invoice Number']);
  let invoiceNumber = invoiceNumberCol ? String(getValue(rowValues, headers, ['Invoice Number'])).trim() : '';

  let pdfFile = null;
  let invoiceLink = '';

  if (paymentYes) {
    if (!invoiceNumber) {
      invoiceNumber = `INV-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}-${rowNumber}`;
      sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Invoice Number')).setValue(invoiceNumber);
    }
    pdfFile = createInvoicePdf(invoiceNumber, data);
    invoiceLink = pdfFile.getUrl();
    sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Invoice PDF Link')).setValue(invoiceLink);
    sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Status')).setValue('Invoice Sent');
  } else {
    sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Status')).setValue('Payment Not Completed');
  }

  // common updates
  sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Item Code')).setValue(data.itemCode);
  sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Amount')).setValue(data.amount);
  sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Invoice Sent Date')).setValue(new Date());
  sheet.getRange(rowNumber, ensureColumn(sheet, headers, 'Follow Up Sent')).setValue('Yes');

  const subject = paymentYes ? `Your Invoice for InkFinity order IFS${invoiceNumber}` : 'Order could not be placed';
  const body = paymentYes
    ? buildEmailBody(data.customerName, true)
    : buildEmailBody(data.customerName, false, 'Your Order cannot be placed because you have not completed the payment.');

  sendEmail(data.email, subject, body, pdfFile);
}

/**
 * Form submit handler - reads Payment Completed? from the row and acts accordingly
 */
function onFormSubmit(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  let rowNumber = 0;
  if (e && e.range) rowNumber = e.range.getRow();
  if (!rowNumber) rowNumber = sheet.getLastRow();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  // look for 'Payment Completed' (with possible '?') and normalize
  const paymentValRaw = getValue(values, headers, ['Payment Completed', 'Payment Completed?', 'Payment']);
  const paymentYes = String(paymentValRaw || '').trim().toLowerCase() === 'yes';

  Logger.log(`onFormSubmit row ${rowNumber} paymentCompleted=${paymentValRaw}`);

  processRowPayment(sheet, rowNumber, paymentYes);
}

/**
 * onEdit handler - triggers only when the 'Payment' column is edited
 * Must be installed as an installable trigger to allow GmailApp usage.
 */
function onEdit(e) {
  if (!e || !e.range) {
    Logger.log('onEdit skipped: no event object');
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getRow() === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const paymentCol = getColumnNumber(headers, ['Payment', 'Payment Completed', 'Payment Completed?']);
  if (paymentCol === 0) {
    Logger.log('onEdit: Payment column not found');
    return;
  }

  if (e.range.getColumn() !== paymentCol) {
    // not a Payment edit
    return;
  }

  // read new value from sheet
  const rowNumber = e.range.getRow();
  const rowValues = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const paymentVal = String(getValue(rowValues, headers, ['Payment', 'Payment Completed', 'Payment Completed?'] || '') || '').trim().toLowerCase();

  Logger.log(`onEdit row ${rowNumber} payment changed to "${paymentVal}"`);

  if (paymentVal === 'yes') {
    processRowPayment(sheet, rowNumber, true);
  } else if (paymentVal === 'no') {
    processRowPayment(sheet, rowNumber, false);
  } else {
    Logger.log('onEdit: payment value not yes/no — no action');
  }
}