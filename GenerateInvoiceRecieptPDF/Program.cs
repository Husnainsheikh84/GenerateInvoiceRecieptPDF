using iTextSharp.text.html.simpleparser;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using System.Reflection;
using List = Microsoft.SharePoint.Client.List;
using iTextSharp.text.pdf.draw;

namespace GenerateInvoiceRecieptPDF
{
    class Program
    {
        static string spurl = ConfigurationManager.AppSettings["Sharepoint_URL"].ToString();
        static string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
        static string pwd = ConfigurationManager.AppSettings["Password"].ToString();
        static string imagePath = ConfigurationManager.AppSettings["ImagePath"].ToString();

        static string imagePath18001 = ConfigurationManager.AppSettings["imagePath18001"].ToString();
        static string imagePath9001 = ConfigurationManager.AppSettings["imagePath9001"].ToString();
        static string imagePath14001 = ConfigurationManager.AppSettings["imagePath14001"].ToString();

        static void Main(string[] args)
        {
            Console.WriteLine("--- Process Started at " + DateTime.Now + " ---");
            ProcessPDFforPendingInvoicesReceipt();
            //CreatePDFFileWithArabicText();
            //SupportPDF("");
            Console.WriteLine("--- Process Ends at " + DateTime.Now + " ---");
            Console.ReadKey();
        }

        private static void ProcessPDFforPendingInvoicesReceipt()
        {
            try
            {
                #region Invoice
                //Console.WriteLine("---Fetching Invoice List Starts " + DateTime.Now + " ---");
                //List<Invoice> invoiceList = new List<Invoice>();
                //invoiceList = GetInvoiceList();
                //Console.WriteLine("---Fetching Invoice List Ends " + DateTime.Now + " ---");
                //Console.WriteLine("---Merge Invoice Html and Generate PDF " + DateTime.Now + " ---");

                //foreach (Invoice objInvoice in invoiceList)
                //{
                //    string strMergedHtml = MergeInvoiceFieldData(GetInvoiceHtml(), objInvoice);

                //    GenerateInvoicePDF(strMergedHtml, objInvoice.InvoiceNO);
                //}
                //Console.WriteLine("---Merge Invoice Html and Generate PDF " + DateTime.Now + " ---");

                #endregion

                #region Receipt
                Console.WriteLine("---Fetching Receipt List Starts " + DateTime.Now + " ---");
                List<Receipt> receiptList = new List<Receipt>();
                receiptList = GetReceiptList();
                Console.WriteLine("---Fetching Receipt List Ends " + DateTime.Now + " ---");
                Console.WriteLine("---Merge Receipt Html and Generate PDF " + DateTime.Now + " ---");

                foreach (Receipt objReceipt in receiptList)
                {
                    //string strReceiptMergedHtml = MergeReceiptFieldData(GetReceiptHtml(), objReceipt);
                    //GenerateReceiptPDF(strReceiptMergedHtml, objReceipt.DocNo);

                    string strReceiptMergedHtml = MergeReceiptFieldData(GetReceiptHtml(), objReceipt);
                    GenerateReceiptPDF(strReceiptMergedHtml, objReceipt);
                }

                Console.WriteLine("---Merge Receipt Html and Generate PDF " + DateTime.Now + " ---");
                #endregion
            }
            catch (Exception Ex)
            {
                Console.WriteLine("Application Encounterd an error during process, please contact andministrator. " + Ex.Message);
            }
        }

        #region Invoice Code
        private static List<Invoice> GetInvoiceList()
        {
            SecureString Password = new SecureString();
            foreach (char c in pwd)
                Password.AppendChar(c);
            List<Invoice> InvoiceList = new List<Invoice>();

            try
            {
                using (ClientContext context = new ClientContext(spurl))
                {

                    CamlQuery query = new CamlQuery();
                    query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Invoice_x0020_Status' /><Value Type='Text'>Not Generated</Value></Eq></Where></Query></View>";

                    context.Credentials = new NetworkCredential(UserName, Password);
                    Microsoft.SharePoint.Client.List invoiceSharePointList = context.Web.Lists.GetByTitle("Invoice");
                    ListItemCollection invoiceListItem = invoiceSharePointList.GetItems(query);
                    //Folder folder = invoiceSharePointList.RootFolder;
                    context.Load(invoiceListItem);
                    context.Load(invoiceSharePointList.RootFolder);
                    context.ExecuteQuery();


                    if (invoiceListItem.Count> 0)
                    {

                        foreach (ListItem itmPass in invoiceListItem)
                        {

                            Invoice invoiceObject = new Invoice();
                            invoiceObject.InvoiceNO = Convert.ToString(itmPass["ID"]);
                            invoiceObject.Date = Convert.ToString(itmPass["Created"]);
                            invoiceObject.CreatedBY = Convert.ToString(itmPass["Author"]);
                            invoiceObject.TransactionID = Convert.ToString(itmPass["TransactionID"]);
                            invoiceObject.EntityName = Convert.ToString(itmPass["EntityName"]);
                            invoiceObject.Address = Convert.ToString(itmPass["Address"]);
                            invoiceObject.TelephoneNumber = Convert.ToString(itmPass["TelephoneNumber"]);
                            invoiceObject.PaymentHistoryNumber = Convert.ToString(itmPass["PaymentHistoryNumber"]);
                            invoiceObject.SubTotal = Convert.ToString(itmPass["Sub_x0020_Total"]);
                            invoiceObject.VAT = Convert.ToString(itmPass["VAT"]);
                            invoiceObject.TotalAmount = Convert.ToString(itmPass["Total_x0020_Amount"]);
                            invoiceObject.InvoiceStatus = Convert.ToString(itmPass["Invoice_x0020_Status"]);
                            

                            invoiceObject.InvoiceDetails = GetInvoiceDetailList(invoiceObject.InvoiceNO);

                            InvoiceList.Add(invoiceObject);

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return InvoiceList;
        }

        private static void GenerateInvoicePDF(string strHtml, string invoiceNO)
        {
            try
            {
                StringReader sr1 = new StringReader(strHtml);
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
                HTMLWorker htmlworker = new HTMLWorker(pdfDoc);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                    pdfDoc.Open();
                    htmlworker.Parse(sr1);
                    pdfDoc.Close();

                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();
                    //System.IO.File.WriteAllBytes(@"D:\WORK\SharePoint\GenerateInvoiceRecieptPDF\GenerateInvoiceRecieptPDF\Image\test.pdf", bytes);

                    //Save document to Document Library
                    bool Status = UploadInvoiceDocument(spurl, "Invoice PDF", "Invoice PDF", invoiceNO, bytes);

                    if (Status)
                    {

                        using (ClientContext context = new ClientContext(spurl))
                        {
                            SecureString Password = new SecureString();
                            foreach (char c in pwd)
                                Password.AppendChar(c);
                            CamlQuery query = new CamlQuery();
                            query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + invoiceNO + "</Value></Eq></Where></Query></View>";

                            context.Credentials = new NetworkCredential(UserName, Password);
                            Microsoft.SharePoint.Client.List invoiceSharePointList = context.Web.Lists.GetByTitle("Invoice");
                            ListItemCollection invoiceListItem = invoiceSharePointList.GetItems(query);
                            context.Load(invoiceListItem);
                            context.ExecuteQuery();
                            var item = invoiceListItem[0];
                            item["Invoice_x0020_Status"] = "Generated";
                            item.Update();
                            context.ExecuteQuery();
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static bool UploadInvoiceDocument(string siteURL, string documentListName, string documentListURL, string documentName, byte[] documentStream)
        {
            bool status = true;
            try
            {
                using (ClientContext clientContext = new ClientContext(siteURL))
                {
                    //Get Document List
                    List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);

                    var fileCreationInformation = new FileCreationInformation();
                    //Assign to content byte[] i.e. documentStream

                    fileCreationInformation.Content = documentStream;
                    //Allow owerwrite of document

                    fileCreationInformation.Overwrite = true;
                    //Upload URL

                    fileCreationInformation.Url = siteURL + "/" + documentListURL + "/" + documentName + ".pdf";
                    Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(
                        fileCreationInformation);

                    ////Update the metadata for a field having name "DocType"
                    //uploadFile.ListItemAllFields["DocType"] = "Favourites";

                    var item = uploadFile.ListItemAllFields;
                    clientContext.Load(item);
                    clientContext.ExecuteQuery();//Not sure if this is needed

                    //Update value of the ListItem
                    item["Invoice_x0020_No"] = documentName;
                    //item.Update();
                    //Save to SharePoint
                    //clientContext.ExecuteQuery();


                    uploadFile.ListItemAllFields.Update();
                    clientContext.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return status;
        }

        private static string GetInvoiceHtml()
        {
            StringBuilder strHtml = new StringBuilder();
            strHtml.AppendLine("<Table width = '70%'><tr><td align = 'center' style = 'background-color:white'><img src = '" + imagePath + "' /></td></tr><tr><td> &nbsp;</td></tr>");
            strHtml.AppendLine("<tr><td align = 'center'> TRN:<TransactionID> </td></tr><tr><td> &nbsp;</td></tr><tr><td align = 'center'><b> TAX INVOICE </b></td></tr><tr><td>&nbsp;</td></tr>");
            strHtml.AppendLine("<tr><td style = 'padding-left:100px;'><Table cellspacing = '10px'><tr><td align = 'left'> Customer Details </td></tr><tr><td> Name </td>");
            strHtml.AppendLine("<td><u><b> <EntityName> </b></u></td></tr><tr><td> Address </td><td><Address>");
            strHtml.AppendLine("</td></tr><tr><td> &nbsp;</td></tr><tr><td> TEL NO.</td><td><TelephoneNumber></td></tr><tr><td> TRN </td>");
            strHtml.AppendLine("<td> <TransactionID> </td></tr></Table></td></tr><tr><td style = 'padding-left:100px;' align = 'center'>");
            strHtml.AppendLine("<Table width = '100%' cellspacing = '10px' border = '0'><tr><td align = 'left'> Invoice NO.:DCA /<InvoiceNO> </td><td align = 'right'> DATE:<Date> </td></tr>");
            strHtml.AppendLine("</Table></td></tr><tr><td style = 'padding-left:100px;' align = 'center'><Table width = '100%' border = '1' style = 'border-style:unset;margin-left:5px;border-spacing:0;'>");
            strHtml.AppendLine("<thead><tr><td align = 'center' width='70%'><b> Details </b></td><td align = 'left' width='30%'><b> Amount(Dh - Fils) </b></td></tr></thead>");
            strHtml.AppendLine("<tbody>");
            strHtml.AppendLine("<PaymentDetails>");
            //strHtml.AppendLine("<tr><td align = 'right'> Sub Total </td><td align = 'right'> @SubTotal </td></tr>");
            //strHtml.AppendLine("<tr><td align = 'right'> VAT </td><td align = 'right'> @VAT</td></tr><tr><td align = 'right'> Total Amount </td><td align = 'right'> @TotalAmount </td></tr>");
            strHtml.AppendLine("</tbody></Table></td></tr><tr><td> &nbsp;</td></tr><tr><td> &nbsp;</td></tr><tr><td style = 'padding-left:100px;' align = 'center'>");
            strHtml.AppendLine("Tel.:(+9717)2449111 - Fax:(+9717)2448861 - P.O.Box:501 - Ras Al Khaimah -United Arab Emirates</td></tr><tr>");
            strHtml.AppendLine("<td style = 'padding-left:100px;' align = 'center'><u> E - mail:finance @rakdca.gov.ae /Website:www.rakdca.gov.ae </u></td></tr></Table>");
            //GenerateInvoicePDF(strHtml.ToString());
            return strHtml.ToString();
        }
       
        private static List<InvoiceDetails> GetInvoiceDetailList(string InvoiceNO)
        {
            SecureString Password = new SecureString();
            foreach (char c in pwd)
                Password.AppendChar(c);
            List<InvoiceDetails> InvoiceDetailList = new List<InvoiceDetails>();
            try
            {

                using (ClientContext context = new ClientContext(spurl))
                {
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Invoice_x0020_NO_x002e_' /><Value Type='Text'>" + InvoiceNO + "</Value></Eq></Where></Query></View>";

                    context.Credentials = new NetworkCredential(UserName, Password);
                    Microsoft.SharePoint.Client.List paymentDetailSharePointList = context.Web.Lists.GetByTitle("Invoice Details");
                    ListItemCollection paymentDetailListItem = paymentDetailSharePointList.GetItems(query);
                    Folder folder = paymentDetailSharePointList.RootFolder;
                    context.Load(paymentDetailListItem);
                    context.Load(paymentDetailSharePointList.RootFolder);
                    context.ExecuteQuery();


                    if (paymentDetailListItem.Count> 0)
                    {

                        foreach (ListItem itmPass in paymentDetailListItem)
                        {

                            InvoiceDetails InvoiceDetailsObject = new InvoiceDetails();
                            InvoiceDetailsObject.InvoiceDetailID = itmPass["ID"].ToString();
                            InvoiceDetailsObject.InvoiceNO = itmPass["Invoice_x0020_NO_x002e_"].ToString();
                            InvoiceDetailsObject.PaymentHistoryNumber = itmPass["PaymentHistoryNumber"].ToString();
                            InvoiceDetailsObject.ServiceName = itmPass["Service_x0020_Name_x0020_"].ToString();
                            InvoiceDetailsObject.ItemDescription = itmPass["Item_x0020_Description"].ToString();                            
                            InvoiceDetailsObject.Amount = itmPass["Amount"].ToString();
                            InvoiceDetailsObject.TaxCode = itmPass["Tax_x0020_Code"].ToString();
                            InvoiceDetailsObject.TaxAmount = itmPass["Tax_x0020_Amount_x0020_"].ToString();
                            InvoiceDetailsObject.TotalAmount = itmPass["Total_x0020_Amount_x0020_"].ToString();
                            InvoiceDetailsObject.GLNo = itmPass["GLNo"].ToString();
                            InvoiceDetailList.Add(InvoiceDetailsObject);

                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }

            return InvoiceDetailList;
        }

        private static string MergeInvoiceFieldData(string strHtmlData, Invoice invoice)
        {
            StringBuilder strPaydetail = new StringBuilder();
            string MergedData = strHtmlData;
            FieldInfo[] fields = invoice.GetType().GetFields(BindingFlags.Public |
                                              BindingFlags.NonPublic |
                                              BindingFlags.Instance);
            foreach (FieldInfo FI in fields)
            {
                if (FI.FieldType.Name == "List`1")
                {
                    foreach (InvoiceDetails objInvoiceDetails in invoice.InvoiceDetails)
                    {
                        strPaydetail.AppendLine("<tr><td align='center' height='100px'>" + objInvoiceDetails.ServiceName + "</td><td align='right'>" + objInvoiceDetails.TotalAmount + "</td></tr>");
                    }

                    strPaydetail.AppendLine("<tr><td align = 'right'> Sub Total </td><td align = 'right'> " + invoice.SubTotal + " </td></tr>");
                    strPaydetail.AppendLine("<tr><td align = 'right'> VAT </td><td align = 'right'> " + invoice.VAT + "</td></tr>");
                    strPaydetail.AppendLine("<tr><td align = 'right'> Total Amount </td><td align = 'right'>" + invoice.SubTotal + " </td></tr>");
                }
                else
                {
                    MergedData = MergedData.Replace(FI.Name.Replace("k__BackingField", ""), Convert.ToString(FI.GetValue(invoice)));
                }
                MergedData = MergedData.Replace(FI.Name, strPaydetail.ToString());
            }


            return MergedData;
        }

        #endregion

        #region Receipt Code
        private static List<Receipt> GetReceiptList()
        {
            SecureString Password = new SecureString();
            foreach (char c in pwd)
                Password.AppendChar(c);
            List<Receipt> ReceiptList = new List<Receipt>();

            try
            {
                using (ClientContext context = new ClientContext(spurl))
                {
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ReceiptStatus' /><Value Type='Text'>Not Generated</Value></Eq></Where></Query></View>";

                    context.Credentials = new NetworkCredential(UserName, Password);
                    Microsoft.SharePoint.Client.List ReceiptSharePointList = context.Web.Lists.GetByTitle("Receipt");
                    ListItemCollection receiptListItem = ReceiptSharePointList.GetItems(query);
                    //Folder folder = ReceiptSharePointList.RootFolder;
                    context.Load(receiptListItem);
                    context.Load(ReceiptSharePointList.RootFolder);
                    context.ExecuteQuery();


                    if (receiptListItem.Count> 0)
                    {

                        foreach (ListItem itmPass in receiptListItem)
                        {

                            Receipt receiptObject = new Receipt();
                            receiptObject.DocNo = Convert.ToString(itmPass["ID"]);
                            receiptObject.ReceiptDate = Convert.ToDateTime(itmPass["Created"]).ToString("dd.MM.yyyy");
                            receiptObject.Amount = Convert.ToString(itmPass["Amount"]);
                            receiptObject.ReceivedFrom = Convert.ToString(itmPass["ReceivedFrom"]);
                            receiptObject.TheSumOfDhs = Convert.ToString(itmPass["TheSumOfDhs"]);
                            receiptObject.PaymentMode = Convert.ToString(itmPass["Payment_x0020_Mode"]);
                            receiptObject.PaymentReferenceNO = Convert.ToString(itmPass["Payment_x0020_Reference_x0020_No"]);
                            receiptObject.Being = Convert.ToString(itmPass["Being"]);                            
                            receiptObject.InvoiceNo = Convert.ToString(itmPass["InvoiceNo"]);
                            receiptObject.ReceiptStatus = Convert.ToString(itmPass["ReceiptStatus"]);
                            ReceiptList.Add(receiptObject);


                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;

            }

            return ReceiptList;
        }       

        private static string MergeReceiptFieldData(string strHtmlData, Receipt receipt)
        {
            StringBuilder strPaydetail = new StringBuilder();
            string MergedData = strHtmlData;
            FieldInfo[] fields = receipt.GetType().GetFields(BindingFlags.Public |
                                              BindingFlags.NonPublic |
                                              BindingFlags.Instance);
            foreach (FieldInfo FI in fields)
            {
               
                    MergedData = MergedData.Replace(FI.Name.Replace("k__BackingField", ""), Convert.ToString(FI.GetValue(receipt)));
                
            }


            return MergedData;
        }        

        private static string GetReceiptHtml()
        {
            StringBuilder strHtml = new StringBuilder();            
            strHtml.AppendLine("<html><head> <meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body style='font-family:ARIALUNI;'>");

            strHtml.AppendLine("<Table width='100%' border='0'><tr><td> &nbsp;</td></tr><tr><td align = 'center' style = 'background-color:white;'><img src = '" + imagePath + "' /></td></tr>");
            strHtml.AppendLine("<tr><td> &nbsp;</td></tr><tr><td align = 'center'  width='100%'><Table width = '100%'> ");
            strHtml.AppendLine("<tr><td width = '30%' align='left'><Table><tr><td><b> Doc.No:</b></td><td align='left'><DocNo></td></tr>");
            strHtml.AppendLine("<tr><td><b>Date:</b></td><td align='left'><ReceiptDate></td></tr></Table></td>");
            strHtml.AppendLine("<td width = '60%'><b> Receipt Voucher </b></td><td width = '10%'>");
            strHtml.AppendLine("<Table border = '1' style='border - style:unset; margin-left:5px; border - spacing:0;'><tr><td align = 'right'><b> AED </b></td><td><b>Fils</b></td></tr><tr><td align = 'right'><Amount></td><td></td></tr></Table>");
            strHtml.AppendLine("</td></tr></Table></td></tr><tr><td> &nbsp;</td></tr> ");
            strHtml.AppendLine("<tr><td> &nbsp;</td></tr><tr><td><Table width='100%'><tr><td><b> Received From:</b> <ReceivedFrom><b>: بأجهزة الكمبيوتر المحمولة لطلبة كلية</b><td></tr></Table></td></tr> ");
            strHtml.AppendLine("<tr><td><Table width='100%'><tr><td><b> The Sum of Dhs:</b> <TheSumOfDhs> <span style='text-align:right'><b>:بأجهزة الكمبيوتر المحمولة لطلبة كلية</b></Span><td></tr></Table></td></tr> ");
            strHtml.AppendLine("<tr><td><Table width='100%'><tr><td><b> Cash /Cheque No:</b> <ChequeNo> <b>:بأجهزة الكمبيوتر المحمولة لطلبة كلية</b><td></tr></Table></td></tr> ");
            strHtml.AppendLine("<tr><td><Table width='100%'><tr><td><b> Being:</b> <Being> <b>:بأجهزة الكمبيوتر المحمولة لطلبة كلية</b><td></tr></Table></td></tr> ");
            strHtml.AppendLine("<tr><td> &nbsp;</td></tr>");
            strHtml.AppendLine("<tr><td><Table width = '100%'><tr><td><b> Signature:</b> Husnain </td><td><b> Receiver's Signature:</b> Husnain</td></tr></Table>");
            strHtml.AppendLine("<tr><tr><td> &nbsp;</td></tr><td>");
            strHtml.AppendLine("<Table width = '100%'><tr><td><b> PS:Proceeds of the cheque(s) will be credited subjects to realization of the amounts </b></td>");
            strHtml.AppendLine("</tr></Table>");
            strHtml.AppendLine("</td></tr></Table>");

            strHtml.AppendLine("</body></html>");
            return strHtml.ToString();
        }
       

        private static void GenerateReceiptPDF(string strHtml, Receipt objReceipt)
        {
            try
            {              

                StringReader sr1 = new StringReader(strHtml);
                Document pdfDoc = new Document(PageSize.A4.Rotate(), 80f, 80f, -2f, 35f);
                pdfDoc.SetMargins(20,0,20,20);
                HTMLWorker htmlworker = new HTMLWorker(pdfDoc);                


                using (MemoryStream memoryStream = new MemoryStream())
                {
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                    pdfDoc.Open();


                    //
                    //PdfWriter writer = PdfWriter.GetInstance(document, new FileStream("D://SqlServer//receipt1.pdf", FileMode.Create));
                    //document.Open();
                    string fontpath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\times.ttf";
                    BaseFont basefont = BaseFont.CreateFont(fontpath, BaseFont.IDENTITY_H, true);
                    iTextSharp.text.Font arabicFont = new iTextSharp.text.Font(basefont, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);



                    var el = new Chunk();
                    iTextSharp.text.Font f2 = new iTextSharp.text.Font(basefont, el.Font.Size, el.Font.Style, el.Font.Color);
                    el.Font = f2;

                    iTextSharp.text.Font b2 = new iTextSharp.text.Font(basefont, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    iTextSharp.text.Font b1 = new iTextSharp.text.Font(basefont, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    iTextSharp.text.Font b3 = new iTextSharp.text.Font(basefont, 7, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    LineSeparator line = new LineSeparator(0, 100, null, Element.ALIGN_CENTER, -2);

                    PdfPTable table = new PdfPTable(4);
                    table.WidthPercentage = 100;


                    //table.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                    //var str = "نام : ";
                    //PdfPCell cell = new PdfPCell(new Phrase(10, str, el.Font));
                    //table.AddCell(cell);


                    //
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imagePath);
                    PdfPCell imageCell = new PdfPCell(jpg);
                    imageCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    imageCell.Colspan = 4;
                    imageCell.Border = 0;
                    table.AddCell(imageCell);




                    PdfPCell EmptyCell = new PdfPCell(new Phrase("\n"));
                    EmptyCell.Colspan = 4;
                    EmptyCell.Border = 0;
                    table.AddCell(EmptyCell);


                    PdfPTable pdfPTable1 = new PdfPTable(1);

                    PdfPCell CellDoc = new PdfPCell();
                    PdfPCell CellDate = new PdfPCell();
                    Phrase phraseDoc = new Phrase();
                    Phrase phraseDate = new Phrase();
                    var chunkDoc = new Chunk("Doc.No: ", b1);
                    var chunkDoc1 = new Chunk(objReceipt.DocNo, el.Font);

                    var chunkDate = new Chunk("Date: ", b1);
                    var chunkDate1 = new Chunk(objReceipt.ReceiptDate, el.Font);
                    chunkDate1.SetUnderline(0, -3);

                    var chunkDate2 = new Chunk(" ", b1);
                    var chunkDate3 = new Chunk("نام", b1);

                    phraseDoc.Add(0, chunkDoc);
                    phraseDoc.Add(1, chunkDoc1);

                    phraseDate.Add(0, chunkDate);
                    phraseDate.Add(1, chunkDate1);
                    phraseDate.Add(2, chunkDate2);
                    phraseDate.Add(3, chunkDate3);

                    CellDoc.AddElement(phraseDoc);
                    CellDoc.Border = 0;
                    pdfPTable1.AddCell(CellDoc);
                    CellDate.AddElement(phraseDate);
                    CellDate.Border = 0;
                    pdfPTable1.AddCell(CellDate);
                    PdfPCell cell1 = new PdfPCell(pdfPTable1);
                    cell1.Border = 0;
                    table.AddCell(cell1);



                    PdfPTable pdfPTable2 = new PdfPTable(1);

                    PdfPCell pdfCellCenterArabicHeader = new PdfPCell(new Phrase(10, "    نامنام    ", arabicFont));
                    pdfCellCenterArabicHeader.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCellCenterArabicHeader.Border = 0;
                    pdfPTable2.AddCell(pdfCellCenterArabicHeader);



                    Chunk chunkVoucher = new Chunk("Receipt Voucher");
                    chunkVoucher.SetUnderline(1, 10);
                    Phrase phaseVoucher = new Phrase();
                    phaseVoucher.Add(chunkVoucher);
                    PdfPCell pdfCellCenterHeader = new PdfPCell(phaseVoucher);
                    pdfCellCenterHeader.HorizontalAlignment = Element.ALIGN_CENTER;
                    pdfCellCenterHeader.Border = 0;
                    pdfPTable2.AddCell(pdfCellCenterHeader);

                    PdfPCell cell2 = new PdfPCell(pdfPTable2);
                    cell2.Border = 0;
                    cell2.Colspan = 2;
                    table.AddCell(cell2);


                    PdfPTable pdfPTable3 = new PdfPTable(3);
                    PdfPCell cell3 = new PdfPCell(pdfPTable3);
                    cell3.Border = 0;


                    Phrase phraseArabCurr = new Phrase();
                    Phrase phraseArabFil = new Phrase();
                    var chunkArabCurr = new Chunk("نامنام", b1);
                    var chunkArabFil = new Chunk("نامنام", b1);
                    phraseArabCurr.Add(0, chunkArabCurr);
                    phraseArabFil.Add(0, chunkArabFil);
                    PdfPCell cell4 = new PdfPCell(phraseArabCurr);
                    cell4.Colspan = 2;
                    cell4.Border = 0;
                    cell4.HorizontalAlignment = Element.ALIGN_RIGHT;
                    PdfPCell cell5 = new PdfPCell(phraseArabFil);
                    cell5.Border = 0;
                    cell5.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfPTable3.AddCell(cell4);
                    pdfPTable3.AddCell(cell5);

                    Phrase phraseEngCurr = new Phrase();
                    Phrase phraseEngFil = new Phrase();
                    var chunkEngCurr = new Chunk("AED ", b1);
                    var chunkEngFil = new Chunk("Fils", b1);
                    phraseEngCurr.Add(0, chunkEngCurr);
                    phraseEngFil.Add(0, chunkEngFil);
                    PdfPCell cell6 = new PdfPCell(phraseEngCurr);
                    cell6.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell6.Border = 0;
                    cell6.Colspan = 2;
                    PdfPCell cell7 = new PdfPCell(phraseEngFil);
                    cell7.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell7.Border = 0;
                    pdfPTable3.AddCell(cell6);
                    pdfPTable3.AddCell(cell7);

                    Phrase phraseAmountCurr = new Phrase();
                    Phrase phraseAmountFil = new Phrase();
                    var chunkAmountCurr = new Chunk(objReceipt.Amount, b1);
                    var chunkAmountFil = new Chunk("00", b1);
                    phraseAmountCurr.Add(0, chunkAmountCurr);
                    phraseAmountFil.Add(0, chunkAmountFil);
                    PdfPCell cell8 = new PdfPCell(phraseAmountCurr);
                    cell8.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell8.Colspan = 2;
                    PdfPCell cell9 = new PdfPCell(phraseAmountFil);
                    cell9.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfPTable3.AddCell(cell8);
                    pdfPTable3.AddCell(cell9);

                    table.AddCell(cell3);


                    //main cell 
                    PdfPCell cell10 = new PdfPCell(new Phrase("\n"));
                    cell10.Colspan = 4;
                    cell10.Border = 0;
                    table.AddCell(cell10);

                    //Received From Table
                    PdfPTable pdfPTable4 = new PdfPTable(7);
                    PdfPCell cell11 = new PdfPCell(pdfPTable4);
                    cell11.Colspan = 4;
                    cell11.Border = 0;
                    table.AddCell(cell11);



                    PdfPCell cell12 = new PdfPCell();
                    cell12.Border = 0;
                    Chunk ChunkPaymentRecFromName = new Chunk("Received From:", b1);
                    Phrase phasePaymentRecFrom = new Phrase();
                    phasePaymentRecFrom.Add(0, ChunkPaymentRecFromName);
                    cell12.AddElement(phasePaymentRecFrom);
                    pdfPTable4.AddCell(cell12);

                    PdfPCell cell12_ = new PdfPCell();
                    cell12_.Border = 0;
                    cell12_.Colspan = 5;
                    Chunk chunkPaymentRecFrom = new Chunk(objReceipt.ReceivedFrom, el.Font);
                    //chunkPaymentRecFrom.SetUnderline(0, -3);              
                    Phrase phasePaymentRecFrom_ = new Phrase();
                    phasePaymentRecFrom_.Add(0, chunkPaymentRecFrom);
                    phasePaymentRecFrom_.Add(1, line);
                    cell12_.AddElement(phasePaymentRecFrom_);
                    pdfPTable4.AddCell(cell12_);

                    PdfPCell cell13 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
                    cell13.VerticalAlignment = Element.ALIGN_BOTTOM;
                    cell13.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell13.Border = 0;
                    pdfPTable4.AddCell(cell13);





                    //The Sum of Dhs Table
                    PdfPTable pdfPTable5 = new PdfPTable(6);
                    PdfPCell cell15 = new PdfPCell(pdfPTable5);
                    cell15.Colspan = 4;
                    cell15.Border = 0;
                    table.AddCell(cell15);

                    PdfPCell cell16 = new PdfPCell();
                    cell16.Border = 0;
                    Chunk ChunkSumName = new Chunk("The Sum of DHS:", b1);
                    Phrase phaseSum = new Phrase();
                    phaseSum.Add(0, ChunkSumName);
                    cell16.AddElement(phaseSum);
                    pdfPTable5.AddCell(cell16);

                    PdfPCell cell16_ = new PdfPCell();
                    cell16_.Border = 0;
                    cell16_.Colspan = 4;
                    Phrase phaseSum_ = new Phrase();
                    Chunk ChunkSumValue = new Chunk(objReceipt.TheSumOfDhs, el.Font);
                    phaseSum_.Add(0, ChunkSumValue);
                    phaseSum_.Add(1, line);
                    cell16_.AddElement(phaseSum_);
                    pdfPTable5.AddCell(cell16_);

                    PdfPCell cell17 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
                    cell17.VerticalAlignment = Element.ALIGN_BOTTOM;
                    cell17.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell17.Border = 0;
                    pdfPTable5.AddCell(cell17);




                    //The Sum of Dhs Table
                    PdfPTable pdfPTable6 = new PdfPTable(7);
                    PdfPCell cell19 = new PdfPCell(pdfPTable6);
                    cell19.Colspan = 4;
                    cell19.Border = 0;
                    table.AddCell(cell19);


                    PdfPCell cell20 = new PdfPCell();
                    cell20.Border = 0;
                    Chunk ChunkPaymode = new Chunk("Cash / Cheque:", b1);
                    Phrase phasePaymode = new Phrase();
                    phasePaymode.Add(0, ChunkPaymode);
                    cell20.AddElement(phasePaymode);
                    pdfPTable6.AddCell(cell20);

                    PdfPCell cell20_ = new PdfPCell();
                    cell20_.Border = 0;
                    cell20_.Colspan = 5;
                    Chunk ChunkPayValue = new Chunk(objReceipt.PaymentReferenceNO, el.Font);
                    Phrase phasePaymode_ = new Phrase();
                    phasePaymode_.Add(0, ChunkPayValue);
                    phasePaymode_.Add(1, line);
                    cell20_.AddElement(phasePaymode_);
                    pdfPTable6.AddCell(cell20_);

                    PdfPCell cell21 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
                    cell21.VerticalAlignment = Element.ALIGN_BOTTOM;
                    cell21.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell21.Border = 0;
                    pdfPTable6.AddCell(cell21);




                    //The Sum of Dhs Table
                    PdfPTable pdfPTable7 = new PdfPTable(7);
                    PdfPCell cell23 = new PdfPCell(pdfPTable7);
                    cell23.Colspan = 4;
                    cell23.Border = 0;
                    table.AddCell(cell23);

                    PdfPCell cell24 = new PdfPCell();
                    cell24.Border = 0;
                    Chunk ChunkBeingName = new Chunk("Being:", b1);
                    Phrase phaseBeingName = new Phrase();
                    phaseBeingName.Add(0, ChunkBeingName);
                    cell24.AddElement(phaseBeingName);
                    pdfPTable7.AddCell(cell24);

                    PdfPCell cell24_ = new PdfPCell();
                    cell24_.Border = 0;
                    cell24_.Colspan = 5;
                    Chunk ChunkBeingValue = new Chunk(objReceipt.Being, el.Font);
                    Phrase phaseBeingName_ = new Phrase();
                    phaseBeingName_.Add(0, ChunkBeingValue);
                    phaseBeingName_.Add(1, line);
                    cell24_.AddElement(phaseBeingName_);
                    pdfPTable7.AddCell(cell24_);

                    PdfPCell cell25 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
                    cell25.VerticalAlignment = Element.ALIGN_BOTTOM;
                    cell25.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell25.Border = 0;
                    pdfPTable7.AddCell(cell25);


                    PdfPTable pdfPTable7_ = new PdfPTable(7);
                    PdfPCell cell25_ = new PdfPCell(pdfPTable7_);
                    cell25_.Colspan = 4;
                    cell25_.Border = 0;
                    PdfPCell cellBlankLine1 = new PdfPCell();
                    cellBlankLine1.Border = 0;
                    pdfPTable7_.AddCell(cellBlankLine1);
                    PdfPCell cellBlankLine2 = new PdfPCell();
                    cellBlankLine2.Border = 0;
                    cellBlankLine2.Colspan = 5;
                    Phrase phraseBlankLine = new Phrase();
                    phraseBlankLine.Add(0, line);
                    cellBlankLine2.AddElement(phraseBlankLine);
                    pdfPTable7_.AddCell(cellBlankLine2);
                    PdfPCell cellBlankLine3 = new PdfPCell();
                    cellBlankLine3.Border = 0;
                    pdfPTable7_.AddCell(cellBlankLine3);
                    table.AddCell(cell25_);

                    //main cell 
                    PdfPCell cell26 = new PdfPCell(new Phrase("\n"));
                    cell26.Colspan = 4;
                    cell26.Border = 0;
                    table.AddCell(cell26);



                    PdfPCell cell27 = new PdfPCell();
                    cell27.Colspan = 4;
                    cell27.Border = 0;
                    table.AddCell(cell27);


                    //Signature Table
                    PdfPTable pdfPTable8 = new PdfPTable(5);
                    PdfPTable pdfPTable9 = new PdfPTable(5);
                    PdfPCell cell28 = new PdfPCell(pdfPTable8);
                    cell28.Colspan = 2;
                    cell28.Border = 0;
                    PdfPCell cell29 = new PdfPCell(pdfPTable9);
                    cell29.Colspan = 2;
                    cell29.Border = 0;
                    table.AddCell(cell28);
                    table.AddCell(cell29);


                    PdfPCell cell30 = new PdfPCell();
                    cell30.Border = 0;
                    Chunk ChunkSigName = new Chunk("Signature:", b1);
                    Phrase phraseSig = new Phrase();
                    phraseSig.Add(0, ChunkSigName);
                    cell30.AddElement(phraseSig);

                    PdfPCell cell30_ = new PdfPCell();
                    cell30_.Colspan = 3;
                    cell30_.Border = 0;
                    cell30_.HorizontalAlignment = Element.ALIGN_CENTER;
                    Phrase phraseSig_ = new Phrase();
                    Chunk ChunkSignature = new Chunk("", el.Font);
                    phraseSig_.Add(0, ChunkSignature);
                    //LineSeparator line = new LineSeparator(1, 100, null, Element.ALIGN_CENTER, -2);
                    phraseSig_.Add(1, line);
                    cell30_.AddElement(phraseSig_);


                    Chunk ChunkArabSigName = new Chunk(":المحمولة", b1);
                    PdfPCell cell31 = new PdfPCell();
                    cell31.Border = 0;
                    cell31.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell31.AddElement(ChunkArabSigName);
                    pdfPTable8.AddCell(cell30);
                    pdfPTable8.AddCell(cell30_);
                    pdfPTable8.AddCell(cell31);



                    PdfPCell cell32 = new PdfPCell();
                    cell32.Border = 0;
                    cell32.Colspan = 2;
                    Chunk ChunkRecSig = new Chunk("Receiver's Signature:", b1);
                    Phrase phraseReceiverSignature = new Phrase();
                    phraseReceiverSignature.Add(0, ChunkRecSig);
                    cell32.AddElement(phraseReceiverSignature);

                    PdfPCell cell32_ = new PdfPCell();
                    cell32_.Border = 0;
                    cell32_.Colspan = 2;
                    Chunk ChunkRecSignature = new Chunk("", el.Font);
                    Phrase phraseReceiverSignature_ = new Phrase();
                    phraseReceiverSignature_.Add(0, ChunkRecSignature);
                    phraseReceiverSignature_.Add(1, line);
                    cell32_.AddElement(phraseReceiverSignature_);
                    Chunk ChunkArabReceivedSignature = new Chunk(":المحمولة", b1);
                    PdfPCell cell33 = new PdfPCell();
                    cell33.Border = 0;
                    cell33.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell33.AddElement(ChunkArabReceivedSignature);
                    pdfPTable9.AddCell(cell32);
                    pdfPTable9.AddCell(cell32_);
                    pdfPTable9.AddCell(cell33);
                                                                          
                    //main cell 
                    PdfPCell cell34 = new PdfPCell(new Phrase("\n"));
                    cell34.Colspan = 4;
                    cell34.Border = 0;
                    table.AddCell(cell34);


                    PdfPCell cell35 = new PdfPCell(new Phrase(10, "PS:Proceeds of the cheque(s) will be credited subjects to realization of the amounts ", b3));
                    cell35.Border = 0;
                    cell35.Colspan = 2;
                    cell35.HorizontalAlignment = Element.ALIGN_LEFT;
                    table.AddCell(cell35);


                    PdfPCell cell36 = new PdfPCell(new Phrase(10, "المحمولة لطلبةالمحمولة لطلبةالمحمولة لطلبةالمحمولة لطلبة", b3));
                    cell36.Border = 0;
                    cell36.Colspan = 2;
                    cell36.HorizontalAlignment = Element.ALIGN_RIGHT;
                    table.AddCell(cell36);


                    //main cell 
                    PdfPCell cell37 = new PdfPCell(new Phrase("\n"));
                    cell37.Colspan = 4;
                    cell37.Border = 0;
                    table.AddCell(cell37);

                    PdfPTable imageTable = new PdfPTable(6);

                    iTextSharp.text.Image image18001 = iTextSharp.text.Image.GetInstance(imagePath18001);
                    iTextSharp.text.Image image9001 = iTextSharp.text.Image.GetInstance(imagePath9001);
                    iTextSharp.text.Image image14001 = iTextSharp.text.Image.GetInstance(imagePath14001);

                    PdfPCell imageCell1 = new PdfPCell(image18001);
                    imageCell1.Colspan = 2;
                    imageCell1.Border = 0;
                    imageCell1.HorizontalAlignment = Element.ALIGN_CENTER;

                    PdfPCell imageCell2 = new PdfPCell(image9001);
                    imageCell2.Colspan = 2;
                    imageCell2.Border = 0;
                    imageCell2.HorizontalAlignment = Element.ALIGN_CENTER;

                    PdfPCell imageCell3 = new PdfPCell(image14001);
                    imageCell3.Colspan = 2;
                    imageCell3.Border = 0;
                    imageCell3.HorizontalAlignment = Element.ALIGN_CENTER;

                    PdfPCell imageCell0 = new PdfPCell();

                    imageTable.AddCell(imageCell1);
                    imageTable.AddCell(imageCell2);
                    imageTable.AddCell(imageCell3);
                    imageCell0.AddElement(imageTable);
                    imageCell0.HorizontalAlignment = Element.ALIGN_CENTER;
                    imageCell0.Colspan = 4;
                    imageCell0.Border = 0;
                    table.AddCell(imageCell0);


                    //
                    //htmlworker.Parse(sr1);
                    
                    pdfDoc.Close();
                    byte[] bytes = memoryStream.ToArray();
                    memoryStream.Close();

                    //Save document to Document Library
                    bool Status = UploadReceiptDocument(spurl, "Receipt PDF", "Receipt PDF", objReceipt.DocNo, bytes);

                    if (Status)
                    {
                        using (ClientContext context = new ClientContext(spurl))
                        {
                            SecureString Password = new SecureString();
                            foreach (char c in pwd)
                                Password.AppendChar(c);
                            CamlQuery query = new CamlQuery();
                            query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + objReceipt.DocNo + "</Value></Eq></Where></Query></View>";

                            context.Credentials = new NetworkCredential(UserName, Password);
                            Microsoft.SharePoint.Client.List receiptSharePointList = context.Web.Lists.GetByTitle("Receipt");
                            ListItemCollection receiptListItem = receiptSharePointList.GetItems(query);
                            context.Load(receiptListItem);
                            context.ExecuteQuery();
                            var item = receiptListItem[0];
                            item["ReceiptStatus"] = "Not Generated";
                            item.Update();
                            context.ExecuteQuery();
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
}

        private static bool UploadReceiptDocument(string siteURL, string documentListName, string documentListURL, string documentName, byte[] documentStream)
        {
            bool status = true;
            try
            {
               
                using (ClientContext clientContext = new ClientContext(siteURL))
                {
                    //Get Document List
                    List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);

                    var fileCreationInformation = new FileCreationInformation();
                    //Assign to content byte[] i.e. documentStream

                    fileCreationInformation.Content = documentStream;
                    //Allow owerwrite of document

                    fileCreationInformation.Overwrite = true;
                    //Upload URL

                    fileCreationInformation.Url = siteURL + "/" + documentListURL + "/" + documentName + ".pdf";
                    Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(
                        fileCreationInformation);

                    ////Update the metadata for a field having name "DocType"
                    //uploadFile.ListItemAllFields["DocType"] = "Favourites";

                    var item = uploadFile.ListItemAllFields;
                    clientContext.Load(item);
                    clientContext.ExecuteQuery();//Not sure if this is needed

                    //Update value of the ListItem
                    item["DocNo"] = documentName;
                    //item.Update();
                    //Save to SharePoint
                    //clientContext.ExecuteQuery();


                    uploadFile.ListItemAllFields.Update();
                    clientContext.ExecuteQuery();
                   
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return status;
        }
        #endregion


        private static void CreatePDFFileWithArabicText()
        {
            FileStream fs = null;
            Document document = null;
            PdfWriter writer = null;
            PdfReader reader = null;
            string str1 = "مرحبا العالم";
            //string str1 = "Hello World";
            try
            {
                string sourceFile = "D:/SqlServer/receipt.pdf";
                string newFile = "D:/SqlServer/receipt1.pdf";
                // open the reader
                reader = new PdfReader(sourceFile);
                Rectangle size = reader.GetPageSizeWithRotation(1);
                int pageCount = reader.NumberOfPages;
                document = new Document(size);
                // open the writer
                fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                // loop through every page of source document
                for (int i = 1; i <= pageCount; i++)
                {
                    //// Read the pdf content
                    PdfContentByte cb = writer.DirectContent;
                    // Get new page size.
                    document.SetPageSize(reader.GetPageSizeWithRotation(1));
                    document.NewPage();
                    float h1 = document.PageSize.Height;
                    //Insert text into the third page of document.
                    if (i == 3)
                    {
                        // select the font properties
                        string fontpath = Environment.GetEnvironmentVariable("SystemRoot") +
                        "\\fonts\\tahoma.ttf";
                        BaseFont basefont = BaseFont.CreateFont
                        (fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font tahomaFont = new Font(basefont, 10, Font.NORMAL, BaseColor.RED);
                        //set the direction of text.
                        ColumnText ct = new ColumnText(writer.DirectContent);
                        ct.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                        //set the position of text in page.
                        ct.SetSimpleColumn(100, 100, 500, 800, 24, Element.ALIGN_RIGHT);
                        var chunk = new Chunk(str1, tahomaFont);
                        ct.AddElement(chunk);
                    }
                    PdfImportedPage page = writer.GetImportedPage(reader, i);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // close the streams 
                //document.Close();
                //fs.Close();
                //writer.Close();
                //reader.Close();
            }
            
        }

        //private static void SupportPDF(string sFilePDF)
        //{
        //    Document document = new Document();
        //    try
        //    {
        //        PdfWriter writer = PdfWriter.GetInstance(document, new FileStream("D://SqlServer//receipt1.pdf", FileMode.Create));
        //        document.Open();
        //        string fontpath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\times.ttf";
        //        BaseFont basefont = BaseFont.CreateFont(fontpath, BaseFont.IDENTITY_H, true);
        //        iTextSharp.text.Font arabicFont = new iTextSharp.text.Font(basefont, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                

        //        var el = new Chunk();
        //        iTextSharp.text.Font f2 = new iTextSharp.text.Font(basefont, el.Font.Size,el.Font.Style,el.Font.Color);
        //        el.Font = f2;

        //        iTextSharp.text.Font b2 = new iTextSharp.text.Font(basefont, 14, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
        //        iTextSharp.text.Font b1 = new iTextSharp.text.Font(basefont, 10, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
        //        iTextSharp.text.Font b3 = new iTextSharp.text.Font(basefont, 7, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
        //        LineSeparator line = new LineSeparator(0, 100, null, Element.ALIGN_CENTER, -2);

        //        PdfPTable table = new PdfPTable(4);
        //        table.WidthPercentage = 100;


        //        //table.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
        //        //var str = "نام : ";
        //        //PdfPCell cell = new PdfPCell(new Phrase(10, str, el.Font));
        //        //table.AddCell(cell);


        //        //
        //        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imagePath);
        //        PdfPCell imageCell = new PdfPCell(jpg);
        //        imageCell.HorizontalAlignment = Element.ALIGN_CENTER;
        //        imageCell.Colspan = 4;
        //        imageCell.Border = 0;
        //        table.AddCell(imageCell);

               


        //        PdfPCell EmptyCell = new PdfPCell(new Phrase("\n"));
        //        EmptyCell.Colspan = 4;
        //        EmptyCell.Border = 0;
        //        table.AddCell(EmptyCell);


        //        PdfPTable pdfPTable1 = new PdfPTable(1);                

        //        PdfPCell CellDoc = new PdfPCell();
        //        PdfPCell CellDate = new PdfPCell();
        //        Phrase phraseDoc = new Phrase();
        //        Phrase phraseDate = new Phrase();
        //        var chunkDoc = new Chunk("Doc.No: ", b1);
        //        var chunkDoc1 = new Chunk("1500000003", el.Font);

        //        var chunkDate = new Chunk("Date: ", b1);
        //        var chunkDate1 = new Chunk("15.3.2015", el.Font);
        //        chunkDate1.SetUnderline(0, -3);

        //        var chunkDate2 = new Chunk(" ", b1);
        //        var chunkDate3 = new Chunk("نام", b1);

        //        phraseDoc.Add(0, chunkDoc);
        //        phraseDoc.Add(1, chunkDoc1);

        //        phraseDate.Add(0, chunkDate);
        //        phraseDate.Add(1, chunkDate1);
        //        phraseDate.Add(2, chunkDate2);
        //        phraseDate.Add(3, chunkDate3);

        //        CellDoc.AddElement(phraseDoc);
        //        CellDoc.Border = 0;
        //        pdfPTable1.AddCell(CellDoc);
        //        CellDate.AddElement(phraseDate);
        //        CellDate.Border = 0;
        //        pdfPTable1.AddCell(CellDate);
        //        PdfPCell cell1 = new PdfPCell(pdfPTable1);
        //        cell1.Border = 0;
        //        table.AddCell(cell1);



        //        PdfPTable pdfPTable2 = new PdfPTable(1);

        //        PdfPCell  pdfCellCenterArabicHeader = new PdfPCell(new Phrase(10, "    نامنام    ", arabicFont));
        //        pdfCellCenterArabicHeader.HorizontalAlignment = Element.ALIGN_CENTER;
        //        pdfCellCenterArabicHeader.Border = 0;               
        //        pdfPTable2.AddCell(pdfCellCenterArabicHeader);
                

               
        //        Chunk chunkVoucher = new Chunk("Receipt Voucher");
        //        chunkVoucher.SetUnderline(1, 10);
        //        Phrase phaseVoucher = new Phrase();
        //        phaseVoucher.Add(chunkVoucher);
        //        PdfPCell pdfCellCenterHeader = new PdfPCell(phaseVoucher);               
        //        pdfCellCenterHeader.HorizontalAlignment = Element.ALIGN_CENTER;
        //        pdfCellCenterHeader.Border = 0;               
        //        pdfPTable2.AddCell(pdfCellCenterHeader);

        //        PdfPCell cell2 = new PdfPCell(pdfPTable2);
        //        cell2.Border = 0;
        //        cell2.Colspan = 2;
        //        table.AddCell(cell2);


        //        PdfPTable pdfPTable3 = new PdfPTable(3);
        //        PdfPCell cell3 = new PdfPCell(pdfPTable3);
        //        cell3.Border = 0;


        //        Phrase phraseArabCurr = new Phrase();
        //        Phrase phraseArabFil = new Phrase();
        //        var chunkArabCurr = new Chunk("نامنام", b1);
        //        var chunkArabFil = new Chunk("نامنام", b1);
        //        phraseArabCurr.Add(0, chunkArabCurr);
        //        phraseArabFil.Add(0, chunkArabFil);
        //        PdfPCell cell4 = new PdfPCell(phraseArabCurr);
        //        cell4.Colspan = 2;
        //        cell4.Border = 0;
        //        cell4.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        PdfPCell cell5 = new PdfPCell(phraseArabFil);
        //        cell5.Border = 0;
        //        cell5.HorizontalAlignment = Element.ALIGN_LEFT;
        //        pdfPTable3.AddCell(cell4);
        //        pdfPTable3.AddCell(cell5);

        //        Phrase phraseEngCurr = new Phrase();
        //        Phrase phraseEngFil = new Phrase();
        //        var chunkEngCurr = new Chunk("AED ", b1);
        //        var chunkEngFil = new Chunk("Fils", b1);
        //        phraseEngCurr.Add(0, chunkEngCurr);
        //        phraseEngFil.Add(0, chunkEngFil);
        //        PdfPCell cell6 = new PdfPCell(phraseEngCurr);
        //        cell6.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell6.Border = 0;
        //        cell6.Colspan = 2;
        //        PdfPCell cell7 = new PdfPCell(phraseEngFil);
        //        cell7.HorizontalAlignment = Element.ALIGN_LEFT;
        //        cell7.Border = 0;
        //        pdfPTable3.AddCell(cell6);
        //        pdfPTable3.AddCell(cell7);

        //        Phrase phraseAmountCurr = new Phrase();
        //        Phrase phraseAmountFil = new Phrase();
        //        var chunkAmountCurr = new Chunk("456 ", b1);
        //        var chunkAmountFil = new Chunk("00", b1);
        //        phraseAmountCurr.Add(0, chunkAmountCurr);
        //        phraseAmountFil.Add(0, chunkAmountFil);
        //        PdfPCell cell8 = new PdfPCell(phraseAmountCurr);
        //        cell8.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell8.Colspan = 2;
        //        PdfPCell cell9 = new PdfPCell(phraseAmountFil);
        //        cell9.HorizontalAlignment = Element.ALIGN_LEFT;
        //        pdfPTable3.AddCell(cell8);
        //        pdfPTable3.AddCell(cell9);

        //        table.AddCell(cell3);


        //        //main cell 
        //        PdfPCell cell10 = new PdfPCell(new Phrase("\n"));
        //        cell10.Colspan = 4;
        //        cell10.Border = 0;
        //        table.AddCell(cell10);

        //        //Received From Table
        //        PdfPTable pdfPTable4 = new PdfPTable(7);
        //        PdfPCell cell11 = new PdfPCell(pdfPTable4);                
        //        cell11.Colspan = 4;
        //        cell11.Border = 0;
        //        table.AddCell(cell11);



        //        PdfPCell cell12 = new PdfPCell();
        //        cell12.Border = 0;               
        //        Chunk ChunkPaymentRecFromName = new Chunk("Received From:", b1);
        //        Phrase phasePaymentRecFrom = new Phrase();
        //        phasePaymentRecFrom.Add(0, ChunkPaymentRecFromName);
        //        cell12.AddElement(phasePaymentRecFrom);
        //        pdfPTable4.AddCell(cell12);

        //        PdfPCell cell12_ = new PdfPCell();
        //        cell12_.Border = 0;
        //        cell12_.Colspan = 5;                
        //        Chunk chunkPaymentRecFrom = new Chunk(" M/s ADNOC DISTRIBUTION",el.Font);
        //        //chunkPaymentRecFrom.SetUnderline(0, -3);              
        //        Phrase phasePaymentRecFrom_ = new Phrase();
        //        phasePaymentRecFrom_.Add(0,chunkPaymentRecFrom);
        //        phasePaymentRecFrom_.Add(1, line);
        //        cell12_.AddElement(phasePaymentRecFrom_);
        //        pdfPTable4.AddCell(cell12_);

        //        PdfPCell cell13 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
        //        cell13.VerticalAlignment = Element.ALIGN_BOTTOM;
        //        cell13.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell13.Border = 0;
        //        pdfPTable4.AddCell(cell13);



              

        //        //The Sum of Dhs Table
        //        PdfPTable pdfPTable5 = new PdfPTable(6);
        //        PdfPCell cell15 = new PdfPCell(pdfPTable5);
        //        cell15.Colspan = 4;
        //        cell15.Border = 0;
        //        table.AddCell(cell15);

        //        PdfPCell cell16 = new PdfPCell();
        //        cell16.Border = 0;
        //        Chunk ChunkSumName = new Chunk("The Sum of DHS:", b1);
        //        Phrase phaseSum = new Phrase();
        //        phaseSum.Add(0, ChunkSumName);
        //        cell16.AddElement(phaseSum);
        //        pdfPTable5.AddCell(cell16);

        //        PdfPCell cell16_ = new PdfPCell();
        //        cell16_.Border = 0;
        //        cell16_.Colspan = 4;
        //        Phrase phaseSum_ = new Phrase();
        //        Chunk ChunkSumValue = new Chunk(" M/s FIVE HUNDRED AND FIFTY ", el.Font);
        //        phaseSum_.Add(0, ChunkSumValue);
        //        phaseSum_.Add(1, line);
        //        cell16_.AddElement(phaseSum_);
        //        pdfPTable5.AddCell(cell16_);

        //        PdfPCell cell17 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
        //        cell17.VerticalAlignment = Element.ALIGN_BOTTOM;
        //        cell17.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell17.Border = 0;
        //        pdfPTable5.AddCell(cell17);



              
        //        //The Sum of Dhs Table
        //        PdfPTable pdfPTable6 = new PdfPTable(7);
        //        PdfPCell cell19 = new PdfPCell(pdfPTable6);
        //        cell19.Colspan = 4;
        //        cell19.Border = 0;
        //        table.AddCell(cell19);


        //        PdfPCell cell20 = new PdfPCell();
        //        cell20.Border = 0;
        //        Chunk ChunkPaymode = new Chunk("Cash / Cheque:", b1);
        //        Phrase phasePaymode = new Phrase();
        //        phasePaymode.Add(0, ChunkPaymode);
        //        cell20.AddElement(phasePaymode);
        //        pdfPTable6.AddCell(cell20);

        //        PdfPCell cell20_ = new PdfPCell();
        //        cell20_.Border = 0;
        //        cell20_.Colspan = 5;                
        //        Chunk ChunkPayValue = new Chunk(" M/s FIVE HUNDRED AND FIFTY ", el.Font);               
        //        Phrase phasePaymode_ = new Phrase();
        //        phasePaymode_.Add(0, ChunkPayValue);
        //        phasePaymode_.Add(1, line);
        //        cell20_.AddElement(phasePaymode_);
        //        pdfPTable6.AddCell(cell20_);

        //        PdfPCell cell21 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
        //        cell21.VerticalAlignment = Element.ALIGN_BOTTOM;
        //        cell21.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell21.Border = 0;
        //        pdfPTable6.AddCell(cell21);


                

        //        //The Sum of Dhs Table
        //        PdfPTable pdfPTable7 = new PdfPTable(7);
        //        PdfPCell cell23 = new PdfPCell(pdfPTable7);
        //        cell23.Colspan = 4;
        //        cell23.Border = 0;
        //        table.AddCell(cell23);

        //        PdfPCell cell24 = new PdfPCell();
        //        cell24.Border = 0;
        //        Chunk ChunkBeingName = new Chunk("Being:", b1);
        //        Phrase phaseBeingName = new Phrase();
        //        phaseBeingName.Add(0, ChunkBeingName);
        //        cell24.AddElement(phaseBeingName);
        //        pdfPTable7.AddCell(cell24);

        //        PdfPCell cell24_ = new PdfPCell();
        //        cell24_.Border = 0;
        //        cell24_.Colspan = 5;               
        //        Chunk ChunkBeingValue = new Chunk(" pmt received against the invoice DCA/1 ", el.Font);                
        //        Phrase phaseBeingName_ = new Phrase();
        //        phaseBeingName_.Add(0, ChunkBeingValue);
        //        phaseBeingName_.Add(1, line);
        //        cell24_.AddElement(phaseBeingName_);
        //        pdfPTable7.AddCell(cell24_);

        //        PdfPCell cell25 = new PdfPCell(new Phrase(10, ":المحمولة لطلبة", b1));
        //        cell25.VerticalAlignment = Element.ALIGN_BOTTOM;
        //        cell25.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell25.Border = 0;
        //        pdfPTable7.AddCell(cell25);


        //        PdfPTable pdfPTable7_ = new PdfPTable(7);
        //        PdfPCell cell25_ = new PdfPCell(pdfPTable7_);
        //        cell25_.Colspan = 4;
        //        cell25_.Border = 0;
        //        PdfPCell cellBlankLine1 = new PdfPCell();
        //        cellBlankLine1.Border = 0;
        //        pdfPTable7_.AddCell(cellBlankLine1);
        //        PdfPCell cellBlankLine2 = new PdfPCell();
        //        cellBlankLine2.Border = 0;
        //        cellBlankLine2.Colspan = 5;
        //        Phrase phraseBlankLine = new Phrase();
        //        phraseBlankLine.Add(0, line);
        //        cellBlankLine2.AddElement(phraseBlankLine);
        //        pdfPTable7_.AddCell(cellBlankLine2);
        //        PdfPCell cellBlankLine3 = new PdfPCell();
        //        cellBlankLine3.Border = 0;
        //        pdfPTable7_.AddCell(cellBlankLine3);
        //        table.AddCell(cell25_);

        //        //main cell 
        //        PdfPCell cell26 = new PdfPCell(new Phrase("\n"));
        //        cell26.Colspan = 4;
        //        cell26.Border = 0;
        //        table.AddCell(cell26);



        //        PdfPCell cell27 = new PdfPCell();
        //        cell27.Colspan = 4;
        //        cell27.Border = 0;
        //        table.AddCell(cell27);


        //        //Signature Table
        //        PdfPTable pdfPTable8 = new PdfPTable(5);
        //        PdfPTable pdfPTable9 = new PdfPTable(5);
        //        PdfPCell cell28 = new PdfPCell(pdfPTable8);
        //        cell28.Colspan = 2;
        //        cell28.Border = 0;
        //        PdfPCell cell29 = new PdfPCell(pdfPTable9);
        //        cell29.Colspan = 2;
        //        cell29.Border = 0;
        //        table.AddCell(cell28);
        //        table.AddCell(cell29);

                               
        //        PdfPCell cell30 = new PdfPCell();
        //        cell30.Border = 0;
        //        Chunk ChunkSigName = new Chunk("Signature:", b1);
        //        Phrase phraseSig = new Phrase();
        //        phraseSig.Add(0, ChunkSigName);
        //        cell30.AddElement(phraseSig);

        //        PdfPCell cell30_ = new PdfPCell();
        //        cell30_.Colspan = 3;
        //        cell30_.Border = 0;
        //        cell30_.HorizontalAlignment = Element.ALIGN_CENTER;
        //        Phrase phraseSig_ = new Phrase();
        //        Chunk ChunkSignature = new Chunk("", el.Font);
        //        phraseSig_.Add(0, ChunkSignature);
        //        //LineSeparator line = new LineSeparator(1, 100, null, Element.ALIGN_CENTER, -2);
        //        phraseSig_.Add(1, line);
        //        cell30_.AddElement(phraseSig_);


        //        Chunk ChunkArabSigName = new Chunk(":المحمولة", b1);
        //        PdfPCell cell31 = new PdfPCell();
        //        cell31.Border = 0;
        //        cell31.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell31.AddElement(ChunkArabSigName);
        //        pdfPTable8.AddCell(cell30);
        //        pdfPTable8.AddCell(cell30_);
        //        pdfPTable8.AddCell(cell31);



        //        PdfPCell cell32 = new PdfPCell();
        //        cell32.Border = 0;
        //        cell32.Colspan =2 ;
        //        Chunk ChunkRecSig = new Chunk("Receiver's Signature:", b1);
        //        Phrase phraseReceiverSignature = new Phrase();
        //        phraseReceiverSignature.Add(0, ChunkRecSig);
        //        cell32.AddElement(phraseReceiverSignature);

        //        PdfPCell cell32_ = new PdfPCell();
        //        cell32_.Border = 0;
        //        cell32_.Colspan = 2;                
        //        Chunk ChunkRecSignature = new Chunk("",el.Font);
        //        Phrase phraseReceiverSignature_ = new Phrase();               
        //        phraseReceiverSignature_.Add(0, ChunkRecSignature);
        //        phraseReceiverSignature_.Add(1, line);               
        //        cell32_.AddElement(phraseReceiverSignature_);
        //        Chunk ChunkArabReceivedSignature = new Chunk(":المحمولة", b1);
        //        PdfPCell cell33 = new PdfPCell();
        //        cell33.Border = 0;
        //        cell33.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        cell33.AddElement(ChunkArabReceivedSignature);                
        //        pdfPTable9.AddCell(cell32);
        //        pdfPTable9.AddCell(cell32_);
        //        pdfPTable9.AddCell(cell33);


                


        //        //main cell 
        //        PdfPCell cell34 = new PdfPCell(new Phrase("\n"));
        //        cell34.Colspan = 4;
        //        cell34.Border = 0;
        //        table.AddCell(cell34);


        //        PdfPCell cell35 = new PdfPCell(new Phrase(10, "PS:Proceeds of the cheque(s) will be credited subjects to realization of the amounts ", b3));
        //        cell35.Border = 0;
        //        cell35.Colspan = 2;
        //        cell35.HorizontalAlignment = Element.ALIGN_LEFT;
        //        table.AddCell(cell35);


        //        PdfPCell cell36 = new PdfPCell(new Phrase(10, "المحمولة لطلبةالمحمولة لطلبةالمحمولة لطلبةالمحمولة لطلبة", b3));
        //        cell36.Border = 0;
        //        cell36.Colspan = 2;
        //        cell36.HorizontalAlignment = Element.ALIGN_RIGHT;
        //        table.AddCell(cell36);


        //        //main cell 
        //        PdfPCell cell37 = new PdfPCell(new Phrase("\n"));
        //        cell37.Colspan = 4;
        //        cell37.Border = 0;
        //        table.AddCell(cell37);

        //        PdfPTable imageTable = new PdfPTable(6);

        //        iTextSharp.text.Image image18001 = iTextSharp.text.Image.GetInstance(imagePath18001);
        //        iTextSharp.text.Image image9001 = iTextSharp.text.Image.GetInstance(imagePath9001);
        //        iTextSharp.text.Image image14001 = iTextSharp.text.Image.GetInstance(imagePath14001);

        //        PdfPCell imageCell1 = new PdfPCell(image18001);
        //        imageCell1.Colspan = 2;
        //        imageCell1.Border = 0;
        //        imageCell1.HorizontalAlignment = Element.ALIGN_CENTER;

        //        PdfPCell imageCell2 = new PdfPCell(image9001);
        //        imageCell2.Colspan = 2;
        //        imageCell2.Border = 0;
        //        imageCell2.HorizontalAlignment = Element.ALIGN_CENTER;

        //        PdfPCell imageCell3 = new PdfPCell(image14001);
        //        imageCell3.Colspan = 2;
        //        imageCell3.Border = 0;
        //        imageCell3.HorizontalAlignment = Element.ALIGN_CENTER;

        //        PdfPCell imageCell0 = new PdfPCell();

        //        imageTable.AddCell(imageCell1);
        //        imageTable.AddCell(imageCell2);
        //        imageTable.AddCell(imageCell3);
        //        imageCell0.AddElement(imageTable);
        //        imageCell0.HorizontalAlignment = Element.ALIGN_CENTER;
        //        imageCell0.Colspan = 4;
        //        imageCell0.Border = 0;
        //        table.AddCell(imageCell0);             
               

        //        document.Add(table);

        //        document.Close();

        //    }
        //    catch (DocumentException de)
        //    {
        //        //              this.Message = de.Message;
        //    }
        //    catch (IOException ioe)
        //    {
        //        //                this.Message = ioe.Message;
        //    }

        //    // step 5: we close the document
        //    document.Close();


        //}
    }
}


