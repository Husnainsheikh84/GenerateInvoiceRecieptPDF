using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateInvoiceRecieptPDF
{
    class Invoice
    {
        public string InvoiceNO { get; set; }
        public string PaymentHistoryNumber { get; set; }
        public string Date { get; set; }
        public string CreatedBY { get; set; }
        public string TransactionID { get; set; }
        public string EntityName { get; set; }
        public string Address { get; set; }
        public string TelephoneNumber { get; set; }        
        public string SubTotal { get; set; }
        public string VAT { get; set; }
        public string TotalAmount { get; set; }
        public string InvoiceStatus { get; set; }       

        public List<InvoiceDetails> InvoiceDetails;
    }
}
