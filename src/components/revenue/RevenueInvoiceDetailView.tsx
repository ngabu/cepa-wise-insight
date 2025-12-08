import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { ArrowLeft, Download, Printer, Ban, Loader2, AlertCircle } from 'lucide-react';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { format } from 'date-fns';
import emblem from '@/assets/png-emblem.png';
import { Invoice } from './types';

interface RevenueInvoiceDetailViewProps {
  invoice: Invoice & {
    inspection?: {
      inspection_type: string;
      province?: string;
      number_of_days?: number;
      scheduled_date: string;
    };
    intent_registration?: {
      activity_description: string;
      status: string;
    };
    verification_status?: string;
    verification_notes?: string;
    invoice_type?: string;
  };
  onBack: () => void;
  onSuspend?: () => void;
  isSuspending?: boolean;
}

export function RevenueInvoiceDetailView({ 
  invoice, 
  onBack, 
  onSuspend,
  isSuspending = false 
}: RevenueInvoiceDetailViewProps) {
  
  const handlePrint = () => {
    window.print();
  };

  const handleDownload = () => {
    console.log('Download invoice');
  };

  const canSuspend = !invoice.source_dashboard || invoice.source_dashboard === 'revenue';
  const isNotSuspended = invoice.status !== 'suspended';
  const isNotPaid = invoice.status !== 'paid';

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'paid': return 'bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-200';
      case 'draft': return 'bg-gray-100 text-gray-800 dark:bg-gray-900 dark:text-gray-200';
      case 'sent': return 'bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200';
      case 'overdue': return 'bg-red-100 text-red-800 dark:bg-red-900 dark:text-red-200';
      case 'cancelled': return 'bg-gray-100 text-gray-800 dark:bg-gray-900 dark:text-gray-200';
      case 'suspended': return 'bg-orange-100 text-orange-800 dark:bg-orange-900 dark:text-orange-200';
      default: return 'bg-muted text-muted-foreground';
    }
  };

  // Calculate invoice items based on type
  const invoiceItems = [
    {
      quantity: 1,
      itemCode: invoice.invoice_type === 'inspection_fee' ? 'INSP-FEE' : 'PERMIT-FEE',
      description: invoice.invoice_type === 'inspection_fee' 
        ? `Inspection Fee - ${invoice.inspection?.inspection_type || 'Field Inspection'}` 
        : invoice.permit?.title || 'Permit Application Fee',
      unitPrice: invoice.amount,
      disc: 0,
      totalPrice: invoice.amount
    }
  ];

  return (
    <div className="space-y-4">
      {/* Action Buttons */}
      <div className="flex justify-between items-center print:hidden">
        <Button variant="outline" onClick={onBack}>
          <ArrowLeft className="w-4 h-4 mr-2" />
          Back to Invoices
        </Button>
        <div className="flex gap-2">
          <Button variant="outline" onClick={handleDownload}>
            <Download className="w-4 h-4 mr-2" />
            Download
          </Button>
          <Button variant="outline" onClick={handlePrint}>
            <Printer className="w-4 h-4 mr-2" />
            Print
          </Button>
          {canSuspend && isNotSuspended && isNotPaid && onSuspend && (
            <Button 
              variant="outline"
              className="text-orange-600 hover:text-orange-700 hover:bg-orange-50 border-orange-300"
              onClick={onSuspend}
              disabled={isSuspending}
            >
              {isSuspending ? (
                <Loader2 className="w-4 h-4 mr-2 animate-spin" />
              ) : (
                <Ban className="w-4 h-4 mr-2" />
              )}
              {isSuspending ? 'Suspending...' : 'Suspend Invoice'}
            </Button>
          )}
        </div>
      </div>

      {/* External Invoice Warning */}
      {!canSuspend && (
        <Alert variant="destructive" className="print:hidden">
          <AlertCircle className="h-4 w-4" />
          <AlertTitle>External Invoice</AlertTitle>
          <AlertDescription>
            This invoice was created on the <strong className="capitalize">{invoice.source_dashboard}</strong> dashboard. 
            To suspend this invoice, please go to the {invoice.source_dashboard} dashboard.
          </AlertDescription>
        </Alert>
      )}

      {/* Invoice Document */}
      <Card className="p-8 bg-white">
        {/* Header with Emblem and Authority Info */}
        <div className="flex justify-between items-start mb-8">
          <div className="flex items-start gap-4">
            <img src={emblem} alt="CEPA Emblem" className="w-24 h-24 object-contain" />
            <div>
              <h1 className="text-lg font-bold text-foreground">
                Conservation & Environment Protection Authority
              </h1>
              <div className="text-sm text-muted-foreground mt-2 space-y-1">
                <p>Tower 1, Dynasty Twin Tower</p>
                <p>Savannah Heights, Waigani</p>
                <p>P.O. Box 6601/BOROKO, NCD</p>
                <p>Papua New Guinea</p>
              </div>
            </div>
          </div>
          
          <div className="text-right space-y-2">
            <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-sm">
              <span className="font-semibold text-muted-foreground">Invoice:</span>
              <span className="text-red-600 font-semibold">{invoice.invoice_number}</span>
              
              <span className="font-semibold text-muted-foreground">Date:</span>
              <span className="font-semibold">{format(new Date(invoice.created_at), 'MMM dd, yyyy')}</span>
              
              <span className="font-semibold text-muted-foreground">Due Date:</span>
              <span className="font-semibold">{format(new Date(invoice.due_date), 'MMM dd, yyyy')}</span>
              
              <span className="font-semibold text-muted-foreground">Status:</span>
              <Badge className={getStatusColor(invoice.status)}>
                {invoice.status}
              </Badge>
              
              <span className="font-semibold text-muted-foreground">Payment:</span>
              <Badge variant="outline" className="capitalize">
                {invoice.payment_status?.replace('_', ' ') || 'pending'}
              </Badge>
              
              <span className="font-semibold text-muted-foreground">Source:</span>
              <Badge variant="secondary" className="capitalize text-xs">
                {invoice.source_dashboard || 'revenue'}
              </Badge>
            </div>
          </div>
        </div>

        {/* Client Information */}
        <div className="mb-8 p-4 border border-border">
          <div className="font-semibold text-foreground mb-2">Client:</div>
          <div className="text-sm text-muted-foreground">
            <p className="font-semibold text-foreground">{invoice.entity?.name || 'N/A'}</p>
            <p className="capitalize">{invoice.entity?.entity_type || 'Individual'}</p>
          </div>
        </div>

        {/* Associated Records */}
        {(invoice.permit || invoice.inspection || invoice.intent_registration) && (
          <div className="mb-8 grid grid-cols-1 md:grid-cols-2 gap-4">
            {invoice.permit && (
              <div className="p-4 border border-border rounded-lg bg-muted/30">
                <div className="font-semibold text-foreground mb-2">Associated Permit</div>
                <p className="font-medium">{invoice.permit.title}</p>
                <p className="text-sm text-muted-foreground">
                  {invoice.permit.permit_number && `${invoice.permit.permit_number} • `}
                  {invoice.permit.permit_type}
                </p>
              </div>
            )}
            
            {invoice.inspection && (
              <div className="p-4 border border-border rounded-lg bg-muted/30">
                <div className="font-semibold text-foreground mb-2">Associated Inspection</div>
                <p className="font-medium">{invoice.inspection.inspection_type}</p>
                <p className="text-sm text-muted-foreground">
                  {invoice.inspection.province && `${invoice.inspection.province} • `}
                  {invoice.inspection.number_of_days} day(s)
                </p>
                <p className="text-sm text-muted-foreground">
                  Scheduled: {format(new Date(invoice.inspection.scheduled_date), 'MMM dd, yyyy')}
                </p>
              </div>
            )}

            {invoice.intent_registration && (
              <div className="p-4 border border-border rounded-lg bg-muted/30">
                <div className="font-semibold text-foreground mb-2">Associated Intent Registration</div>
                <p className="font-medium line-clamp-2">{invoice.intent_registration.activity_description}</p>
                <Badge variant="outline" className="mt-1 capitalize">
                  {invoice.intent_registration.status}
                </Badge>
              </div>
            )}
          </div>
        )}

        {/* Invoice Items Table */}
        <div className="mb-8">
          <table className="w-full border-collapse border border-border">
            <thead>
              <tr className="bg-muted/50">
                <th className="border border-border p-2 text-left text-sm font-semibold">QUANTITY</th>
                <th className="border border-border p-2 text-left text-sm font-semibold">ITEM CODE</th>
                <th className="border border-border p-2 text-left text-sm font-semibold">DESCRIPTION</th>
                <th className="border border-border p-2 text-right text-sm font-semibold">UNIT PRICE(ex. GST)</th>
                <th className="border border-border p-2 text-right text-sm font-semibold">DISC %</th>
                <th className="border border-border p-2 text-right text-sm font-semibold">TOTAL PRICE(ex. GST)</th>
              </tr>
            </thead>
            <tbody>
              {invoiceItems.map((item, index) => (
                <tr key={index}>
                  <td className="border border-border p-2 text-center">{item.quantity}</td>
                  <td className="border border-border p-2">{item.itemCode}</td>
                  <td className="border border-border p-2">{item.description}</td>
                  <td className="border border-border p-2 text-right">
                    K{item.unitPrice.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="border border-border p-2 text-right">{item.disc}</td>
                  <td className="border border-border p-2 text-right">
                    K{item.totalPrice.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                  </td>
                </tr>
              ))}
              {/* Empty rows for spacing */}
              {Array(4).fill(0).map((_, index) => (
                <tr key={`empty-${index}`}>
                  <td className="border border-border p-2 h-12">&nbsp;</td>
                  <td className="border border-border p-2">&nbsp;</td>
                  <td className="border border-border p-2">&nbsp;</td>
                  <td className="border border-border p-2">&nbsp;</td>
                  <td className="border border-border p-2">&nbsp;</td>
                  <td className="border border-border p-2">&nbsp;</td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Totals */}
          <div className="flex justify-end mt-4">
            <div className="w-96 space-y-2">
              <div className="flex justify-between border border-border p-2">
                <span className="font-semibold">Subtotal:</span>
                <span>K{invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
              </div>
              <div className="flex justify-between border border-border p-2">
                <span className="font-semibold">Freight (ex. GST):</span>
                <span>K0.00</span>
              </div>
              <div className="flex justify-between border border-border p-2">
                <span className="font-semibold">GST:</span>
                <span>K0.00</span>
              </div>
              <div className="flex justify-between border border-border p-2 bg-muted/30">
                <span className="font-bold">Total (inc. GST):</span>
                <span className="font-bold">K{invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
              </div>
              <div className="flex justify-between border border-border p-2">
                <span className="font-semibold">Paid to Date:</span>
                <span>K{invoice.status === 'paid' ? invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 }) : '0.00'}</span>
              </div>
              <div className="flex justify-between border border-border p-2 bg-muted/50">
                <span className="font-bold text-lg">Balance Due:</span>
                <span className="font-bold text-lg">
                  K{invoice.status === 'paid' ? '0.00' : invoice.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Payment Terms */}
        <div className="mb-8">
          <div className="font-bold text-foreground mb-2">PAYMENT TERMS:</div>
          <div className="text-sm text-muted-foreground ml-8">
            Cheque payable to:
            <div className="font-semibold text-foreground mt-1">
              CONSERVATION & ENVIRONMENT PROTECTION AUTHORITY
            </div>
          </div>
        </div>

        {/* Bank Accounts */}
        <div className="grid grid-cols-3 gap-4 text-sm mb-8">
          <div className="border border-border p-4">
            <div className="font-bold text-center mb-3">RECURRENT ACCOUNT</div>
            <div className="space-y-1">
              <div className="flex justify-between">
                <span className="font-semibold">BANK:</span>
                <span>BANK SOUTH PACIFIC</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">ACCOUNT:</span>
                <span>7003101749</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BRANCH:</span>
                <span>PORT MORESBY</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BSB:</span>
                <span>088-294</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">SWIFTCODE:</span>
                <span>BOSPPGPM</span>
              </div>
            </div>
          </div>

          <div className="border border-border p-4">
            <div className="font-bold text-center mb-3">OPERATIONAL ACCOUNT</div>
            <div className="space-y-1">
              <div className="flex justify-between">
                <span className="font-semibold">BANK:</span>
                <span>BANK SOUTH PACIFIC</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">ACCOUNT:</span>
                <span>7003101905</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BRANCH:</span>
                <span>PORT MORESBY</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BSB:</span>
                <span>088-294</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">SWIFTCODE:</span>
                <span>BOSPPGPM</span>
              </div>
            </div>
          </div>

          <div className="border border-border p-4">
            <div className="font-bold text-center mb-3">DOF-CEPA REVENUE ACCOUNT</div>
            <div className="space-y-1">
              <div className="flex justify-between">
                <span className="font-semibold">BANK:</span>
                <span>BANK SOUTH PACIFIC</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">ACCOUNT:</span>
                <span>7012975319</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BRANCH:</span>
                <span>PORT MORESBY</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">BSB:</span>
                <span>088-294</span>
              </div>
              <div className="flex justify-between">
                <span className="font-semibold">SWIFTCODE:</span>
                <span>BOSPPGPM</span>
              </div>
            </div>
          </div>
        </div>

        {/* Assigned Officer (Revenue Staff Info) */}
        {invoice.assigned_officer && (
          <div className="mb-4 text-sm">
            <span className="font-semibold">Assigned Officer:</span>{' '}
            {invoice.assigned_officer.full_name || 'N/A'} ({invoice.assigned_officer.email})
          </div>
        )}

        {/* Page Number */}
        <div className="text-right text-sm text-muted-foreground">
          Page 1 of 1
        </div>
      </Card>
    </div>
  );
}
