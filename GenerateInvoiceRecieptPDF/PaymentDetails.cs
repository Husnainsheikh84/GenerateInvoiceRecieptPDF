using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateInvoiceRecieptPDF
{
    class InvoiceDetails
    {
        public string InvoiceDetailID { get; set; }
        public string PaymentHistoryNumber { get; set; }
        public string ServiceName { get; set; }
        public string ItemDescription { get; set; }
        public string InvoiceNO { get; set; }
        public string Amount { get; set; }
        public string TaxCode { get; set; }
        public string TaxAmount { get; set; }
        public string TotalAmount { get; set; }
        public string GLNo { get; set; }
     

    }
}
