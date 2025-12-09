import jsPDF from 'jspdf';
import { format } from 'date-fns';

interface InvoiceData {
  invoice_number: string;
  created_at: string;
  amount: number;
  status: string;
  entity?: {
    name: string;
  } | null;
  permit?: {
    title: string;
    permit_number?: string;
    permit_type?: string;
  } | null;
  inspection?: {
    inspection_type: string;
    province?: string;
    number_of_days?: number;
    scheduled_date: string;
  } | null;
  intent_registration?: {
    activity_description: string;
    status: string;
  } | null;
  invoice_type?: string;
  item_code?: string | null;
  item_description?: string | null;
}

export function generateInvoicePdf(invoice: InvoiceData): void {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 15;
  let yPos = 15;

  // Helper functions
  const addText = (text: string, x: number, y: number, options?: { fontSize?: number; fontStyle?: 'normal' | 'bold'; color?: [number, number, number]; align?: 'left' | 'center' | 'right' }) => {
    doc.setFontSize(options?.fontSize || 10);
    doc.setFont('helvetica', options?.fontStyle || 'normal');
    if (options?.color) {
      doc.setTextColor(...options.color);
    } else {
      doc.setTextColor(0, 0, 0);
    }
    
    if (options?.align === 'right') {
      doc.text(text, x, y, { align: 'right' });
    } else if (options?.align === 'center') {
      doc.text(text, x, y, { align: 'center' });
    } else {
      doc.text(text, x, y);
    }
  };

  // Load and add PNG Emblem
  const img = new Image();
  img.crossOrigin = 'anonymous';
  img.src = '/images/png-emblem.png';
  
  // Add emblem placeholder position
  const emblemSize = 22;
  const emblemX = margin;
  const emblemY = yPos;
  
  // Try to add emblem - it will be added when image loads
  try {
    doc.addImage('/images/png-emblem.png', 'PNG', emblemX, emblemY, emblemSize, emblemSize);
  } catch (e) {
    // Emblem not available, continue without it
  }

  // Header - Authority Name (positioned after emblem)
  const headerX = margin + emblemSize + 8;
  addText('Conservation & Environment', headerX, yPos + 6, { 
    fontSize: 13, 
    fontStyle: 'bold'
  });
  addText('Protection Authority', headerX, yPos + 12, { 
    fontSize: 13, 
    fontStyle: 'bold'
  });
  
  addText('Tower 1, Dynasty Twin Tower', headerX, yPos + 18, { fontSize: 9 });
  addText('Savannah Heights, Waigani', headerX, yPos + 22, { fontSize: 9 });
  addText('P.O. Box 6601/BOROKO, NCD', headerX, yPos + 26, { fontSize: 9 });
  addText('Papua New Guinea', headerX, yPos + 30, { fontSize: 9 });

  // Invoice Details - Right side (two-column layout: labels and values)
  const labelX = pageWidth - margin - 70;
  const valueX = pageWidth - margin;
  
  addText('Invoice:', labelX, yPos + 4, { fontSize: 9 });
  addText(invoice.invoice_number, valueX, yPos + 4, { fontSize: 9, fontStyle: 'bold', color: [200, 0, 0], align: 'right' });
  
  addText('Date:', labelX, yPos + 10, { fontSize: 9 });
  addText(format(new Date(invoice.created_at), 'dd/MM/yyyy'), valueX, yPos + 10, { fontSize: 9, align: 'right' });
  
  addText('Your Ref:', labelX, yPos + 16, { fontSize: 9 });
  addText('', valueX, yPos + 16, { fontSize: 9, align: 'right' });
  
  addText('Contact:', labelX, yPos + 22, { fontSize: 9 });
  addText('Kavau Diagoro, Manager Revenue', valueX, yPos + 22, { fontSize: 9, align: 'right' });
  
  addText('Telephone:', labelX, yPos + 28, { fontSize: 9 });
  addText('(675) 3014665/3014614', valueX, yPos + 28, { fontSize: 9, align: 'right' });
  
  addText('Email:', labelX, yPos + 34, { fontSize: 9 });
  addText('revenuemanager@cepa.gov.pg', valueX, yPos + 34, { fontSize: 9, color: [0, 128, 0], align: 'right' });

  yPos += 42;

  // Divider line
  doc.setDrawColor(200, 200, 200);
  doc.line(margin, yPos, pageWidth - margin, yPos);
  yPos += 8;

  // Client Information Box
  doc.setDrawColor(200, 200, 200);
  doc.rect(margin, yPos, pageWidth - 2 * margin, 18);
  yPos += 6;
  addText('Client:', margin + 4, yPos, { fontSize: 10, fontStyle: 'bold' });
  yPos += 6;
  addText(invoice.entity?.name || 'N/A', margin + 4, yPos, { fontSize: 10 });
  yPos += 14;

  // Build description with associated context
  const getAssociatedDescription = () => {
    if (invoice.intent_registration) {
      return `for Associated Intent Registration\n${invoice.entity?.name || ''} ${invoice.intent_registration.activity_description}`;
    }
    if (invoice.permit) {
      return `for Associated Permit\n${invoice.entity?.name || ''} ${invoice.permit.title}`;
    }
    if (invoice.inspection) {
      return `for Associated Inspection\n${invoice.entity?.name || ''} ${invoice.inspection.inspection_type}`;
    }
    return '';
  };

  const associatedContext = getAssociatedDescription();
  const baseDescription = invoice.item_description || (invoice.invoice_type === 'inspection_fee' 
    ? `Inspection Fee - ${invoice.inspection?.inspection_type || 'Field Inspection'}` 
    : invoice.permit?.title || 'Permit Application Fee');
  
  const fullDescription = associatedContext 
    ? `${baseDescription} - ${associatedContext}`
    : baseDescription;

  // Invoice Items Table
  const colWidths = [22, 28, 60, 35, 18, 35]; // Adjusted widths: Quantity, Item Code, Description, Unit Price, Disc%, Total
  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const tableStartX = margin;
  const rowHeight = 10;
  
  // Table header
  doc.setFillColor(230, 230, 230);
  doc.rect(tableStartX, yPos, tableWidth, rowHeight, 'F');
  doc.setDrawColor(180, 180, 180);
  doc.rect(tableStartX, yPos, tableWidth, rowHeight);
  
  let colX = tableStartX;
  const headers = ['QUANTITY', 'ITEM CODE', 'DESCRIPTION', 'UNIT PRICE(ex. GST)', 'DISC %', 'TOTAL PRICE(ex. GST)'];
  headers.forEach((header, i) => {
    doc.rect(colX, yPos, colWidths[i], rowHeight);
    const headerFontSize = i >= 3 ? 7 : 8; // Smaller font for price columns
    addText(header, colX + colWidths[i] / 2, yPos + 6, { fontSize: headerFontSize, fontStyle: 'bold', align: 'center' });
    colX += colWidths[i];
  });
  yPos += rowHeight;

  // Table data row
  const dataRowHeight = rowHeight * 2.5;
  doc.rect(tableStartX, yPos, tableWidth, dataRowHeight);
  
  // Draw column borders for the data row
  let borderX = tableStartX;
  colWidths.forEach(width => {
    doc.rect(borderX, yPos, width, dataRowHeight);
    borderX += width;
  });
  
  // Add data - centered quantity
  addText('1', tableStartX + colWidths[0] / 2, yPos + 8, { fontSize: 9, align: 'center' });
  
  // Item code
  addText(invoice.item_code || 'F1', tableStartX + colWidths[0] + colWidths[1] / 2, yPos + 8, { fontSize: 9, align: 'center' });
  
  // Description - left aligned with padding
  const descColX = tableStartX + colWidths[0] + colWidths[1];
  const descriptionLines = doc.splitTextToSize(fullDescription, colWidths[2] - 6);
  let descY = yPos + 8;
  descriptionLines.slice(0, 3).forEach((line: string) => {
    addText(line, descColX + 3, descY, { fontSize: 8 });
    descY += 4;
  });
  
  // Price columns - right aligned
  const priceColX = tableStartX + colWidths[0] + colWidths[1] + colWidths[2];
  addText(`K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, priceColX + colWidths[3] - 3, yPos + 8, { fontSize: 9, align: 'right' });
  addText('0', priceColX + colWidths[3] + colWidths[4] / 2, yPos + 8, { fontSize: 9, align: 'center' });
  addText(`K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, priceColX + colWidths[3] + colWidths[4] + colWidths[5] - 3, yPos + 8, { fontSize: 9, align: 'right' });
  
  yPos += dataRowHeight;

  // Empty rows (reduced)
  for (let i = 0; i < 2; i++) {
    doc.rect(tableStartX, yPos, tableWidth, rowHeight);
    borderX = tableStartX;
    colWidths.forEach(width => {
      doc.rect(borderX, yPos, width, rowHeight);
      borderX += width;
    });
    yPos += rowHeight;
  }

  yPos += 8;

  // Totals section - Right aligned
  const totalsWidth = 85;
  const totalsX = pageWidth - margin - totalsWidth;
  
  const totalsData = [
    { label: 'Subtotal:', value: `K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}` },
    { label: 'Freight (ex. GST):', value: 'K0.00' },
    { label: 'GST:', value: 'K0.00' },
    { label: 'Total (inc. GST):', value: `K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, bold: true },
    { label: 'Paid to Date:', value: invoice.status === 'paid' ? `K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}` : 'K0.00' },
    { label: 'Balance Due:', value: invoice.status === 'paid' ? 'K0.00' : `K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, bold: true }
  ];

  totalsData.forEach(item => {
    doc.rect(totalsX, yPos, totalsWidth, 7);
    addText(item.label, totalsX + 3, yPos + 5, { fontSize: 9, fontStyle: item.bold ? 'bold' : 'normal' });
    addText(item.value, totalsX + totalsWidth - 3, yPos + 5, { fontSize: 9, fontStyle: item.bold ? 'bold' : 'normal', align: 'right' });
    yPos += 7;
  });

  yPos += 10;

  // Payment Terms
  addText('PAYMENT TERMS:', margin, yPos, { fontSize: 10, fontStyle: 'bold' });
  yPos += 6;
  addText('Cheque payable to:', margin + 5, yPos, { fontSize: 9 });
  yPos += 5;
  addText('CONSERVATION & ENVIRONMENT PROTECTION AUTHORITY', margin + 5, yPos, { fontSize: 10, fontStyle: 'bold' });
  yPos += 12;

  // Bank Accounts
  const bankWidth = (pageWidth - 2 * margin - 10) / 3;
  const banks = [
    { name: 'RECURRENT ACCOUNT', account: '7003101749', bsb: '088-294' },
    { name: 'OPERATIONAL ACCOUNT', account: '7003101905', bsb: '088-294' },
    { name: 'DOF-CEPA REVENUE', account: '7012975319', bsb: '088-294' }
  ];

  let bankX = margin;
  banks.forEach(bank => {
    doc.rect(bankX, yPos, bankWidth, 32);
    addText(bank.name, bankX + bankWidth / 2, yPos + 6, { fontSize: 8, fontStyle: 'bold', align: 'center' });
    addText('BANK: BANK SOUTH PACIFIC', bankX + 3, yPos + 12, { fontSize: 7 });
    addText(`ACCOUNT: ${bank.account}`, bankX + 3, yPos + 17, { fontSize: 7 });
    addText('BRANCH: PORT MORESBY', bankX + 3, yPos + 22, { fontSize: 7 });
    addText(`BSB: ${bank.bsb}`, bankX + 3, yPos + 27, { fontSize: 7 });
    bankX += bankWidth + 5;
  });

  yPos += 40;

  // Page number
  addText('Page 1 of 1', pageWidth - margin, yPos, { fontSize: 8, align: 'right' });

  // Save the PDF
  doc.save(`Invoice_${invoice.invoice_number}.pdf`);
}
