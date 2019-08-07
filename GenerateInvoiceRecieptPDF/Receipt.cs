using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateInvoiceRecieptPDF
{
    class Receipt
    {
        public string DocNo { get; set; }
        public string ReceiptDate { get; set; }
        public string Amount { get; set; }
        public string ReceivedFrom { get; set; }
        public string TheSumOfDhs { get; set; }
        public string Being { get; set; }        
        public string InvoiceNo { get; set; }
        public string ReceiptStatus { get; set; }
        public string PaymentMode { get; set; }
        public string PaymentReferenceNO { get; set; }
    }
}
