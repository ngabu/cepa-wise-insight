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
  let yPos = 20;

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

  // Header - Authority Name
  addText('Conservation & Environment Protection Authority', pageWidth / 2, yPos, { 
    fontSize: 14, 
    fontStyle: 'bold', 
    align: 'center' 
  });
  yPos += 6;
  
  addText('Tower 1, Dynasty Twin Tower, Savannah Heights, Waigani', pageWidth / 2, yPos, { 
    fontSize: 9, 
    align: 'center' 
  });
  yPos += 4;
  addText('P.O. Box 6601/BOROKO, NCD, Papua New Guinea', pageWidth / 2, yPos, { 
    fontSize: 9, 
    align: 'center' 
  });
  yPos += 10;

  // Invoice Title
  addText('TAX INVOICE', pageWidth / 2, yPos, { 
    fontSize: 16, 
    fontStyle: 'bold', 
    align: 'center' 
  });
  yPos += 12;

  // Invoice Details - Right side
  const rightColX = pageWidth - margin;
  doc.setFontSize(9);
  
  addText('Invoice:', rightColX - 50, yPos, { fontSize: 9, fontStyle: 'bold' });
  addText(invoice.invoice_number, rightColX, yPos, { fontSize: 9, fontStyle: 'bold', color: [200, 0, 0], align: 'right' });
  yPos += 5;
  
  addText('Date:', rightColX - 50, yPos, { fontSize: 9, fontStyle: 'bold' });
  addText(format(new Date(invoice.created_at), 'dd/MM/yyyy'), rightColX, yPos, { fontSize: 9, align: 'right' });
  yPos += 5;
  
  addText('Contact:', rightColX - 50, yPos, { fontSize: 9, fontStyle: 'bold' });
  addText('Kavau Diagoro, Manager Revenue', rightColX, yPos, { fontSize: 9, align: 'right' });
  yPos += 5;
  
  addText('Telephone:', rightColX - 50, yPos, { fontSize: 9, fontStyle: 'bold' });
  addText('(675) 3014665/3014614', rightColX, yPos, { fontSize: 9, align: 'right' });
  yPos += 5;
  
  addText('Email:', rightColX - 50, yPos, { fontSize: 9, fontStyle: 'bold' });
  addText('revenuemanager@cepa.gov.pg', rightColX, yPos, { fontSize: 9, color: [0, 128, 0], align: 'right' });
  yPos += 12;

  // Client Information Box
  doc.setDrawColor(200, 200, 200);
  doc.rect(margin, yPos, pageWidth - 2 * margin, 20);
  yPos += 5;
  addText('Client:', margin + 3, yPos, { fontSize: 10, fontStyle: 'bold' });
  yPos += 6;
  addText(invoice.entity?.name || 'N/A', margin + 3, yPos, { fontSize: 10 });
  yPos += 18;

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
  const tableStartY = yPos;
  const colWidths = [15, 25, 70, 30, 15, 30]; // Quantity, Item Code, Description, Unit Price, Disc, Total
  const rowHeight = 8;
  
  // Table header
  doc.setFillColor(240, 240, 240);
  doc.rect(margin, yPos, pageWidth - 2 * margin, rowHeight, 'F');
  doc.setDrawColor(200, 200, 200);
  doc.rect(margin, yPos, pageWidth - 2 * margin, rowHeight);
  
  let colX = margin;
  const headers = ['QTY', 'ITEM CODE', 'DESCRIPTION', 'UNIT PRICE', 'DISC%', 'TOTAL'];
  headers.forEach((header, i) => {
    doc.rect(colX, yPos, colWidths[i], rowHeight);
    addText(header, colX + 2, yPos + 5, { fontSize: 8, fontStyle: 'bold' });
    colX += colWidths[i];
  });
  yPos += rowHeight;

  // Table data row
  colX = margin;
  doc.rect(margin, yPos, pageWidth - 2 * margin, rowHeight * 3);
  
  // Draw column borders for the data row
  let borderX = margin;
  colWidths.forEach(width => {
    doc.rect(borderX, yPos, width, rowHeight * 3);
    borderX += width;
  });
  
  // Add data
  addText('1', margin + 5, yPos + 5, { fontSize: 9 });
  addText(invoice.item_code || 'PERMIT-FEE', margin + colWidths[0] + 2, yPos + 5, { fontSize: 8 });
  
  // Split description for multi-line
  const descriptionLines = doc.splitTextToSize(fullDescription, colWidths[2] - 4);
  let descY = yPos + 5;
  descriptionLines.slice(0, 3).forEach((line: string) => {
    addText(line, margin + colWidths[0] + colWidths[1] + 2, descY, { fontSize: 8 });
    descY += 4;
  });
  
  const priceColX = margin + colWidths[0] + colWidths[1] + colWidths[2];
  addText(`K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, priceColX + colWidths[3] - 2, yPos + 5, { fontSize: 9, align: 'right' });
  addText('0', priceColX + colWidths[3] + colWidths[4] - 2, yPos + 5, { fontSize: 9, align: 'right' });
  addText(`K${invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`, priceColX + colWidths[3] + colWidths[4] + colWidths[5] - 2, yPos + 5, { fontSize: 9, align: 'right' });
  
  yPos += rowHeight * 3;

  // Empty rows
  for (let i = 0; i < 2; i++) {
    doc.rect(margin, yPos, pageWidth - 2 * margin, rowHeight);
    borderX = margin;
    colWidths.forEach(width => {
      doc.rect(borderX, yPos, width, rowHeight);
      borderX += width;
    });
    yPos += rowHeight;
  }

  yPos += 5;

  // Totals section - Right aligned
  const totalsX = pageWidth - margin - 80;
  const totalsWidth = 80;
  
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
    addText(item.label, totalsX + 2, yPos + 5, { fontSize: 9, fontStyle: item.bold ? 'bold' : 'normal' });
    addText(item.value, totalsX + totalsWidth - 2, yPos + 5, { fontSize: 9, fontStyle: item.bold ? 'bold' : 'normal', align: 'right' });
    yPos += 7;
  });

  yPos += 10;

  // Payment Terms
  addText('PAYMENT TERMS:', margin, yPos, { fontSize: 10, fontStyle: 'bold' });
  yPos += 6;
  addText('Cheque payable to:', margin + 10, yPos, { fontSize: 9 });
  yPos += 5;
  addText('CONSERVATION & ENVIRONMENT PROTECTION AUTHORITY', margin + 10, yPos, { fontSize: 10, fontStyle: 'bold' });
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
    doc.rect(bankX, yPos, bankWidth, 35);
    addText(bank.name, bankX + bankWidth / 2, yPos + 6, { fontSize: 8, fontStyle: 'bold', align: 'center' });
    addText('BANK: BANK SOUTH PACIFIC', bankX + 3, yPos + 13, { fontSize: 7 });
    addText(`ACCOUNT: ${bank.account}`, bankX + 3, yPos + 18, { fontSize: 7 });
    addText('BRANCH: PORT MORESBY', bankX + 3, yPos + 23, { fontSize: 7 });
    addText(`BSB: ${bank.bsb}`, bankX + 3, yPos + 28, { fontSize: 7 });
    addText('SWIFTCODE: BOSPPGPM', bankX + 3, yPos + 33, { fontSize: 7 });
    bankX += bankWidth + 5;
  });

  yPos += 45;

  // Page number
  addText('Page 1 of 1', pageWidth - margin, yPos, { fontSize: 8, align: 'right' });

  // Save the PDF
  doc.save(`Invoice_${invoice.invoice_number}.pdf`);
}
