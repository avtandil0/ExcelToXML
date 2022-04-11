using ExcelToXML.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ExcelToXML.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment Environment;
        private IConfiguration Configuration;



        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment _environment, IConfiguration _configuration)
        {
            _logger = logger;
            Environment = _environment;
            Configuration = _configuration;

        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("FileUpload")]
        public async Task<IActionResult> FileUpload(List<IFormFile> files)
        {
            long size = files.Sum(f => f.Length);

            var filePaths = new List<string>();
            var file = files[0];
            // if (file == null || file.Length == 0)
            //     return new Result(false, 0, "File Not Found");

            string fileExtension = Path.GetExtension(file.FileName);
            // if (fileExtension != ".xls" && fileExtension != ".xlsx")
            //     return new Result(false, 0, "File Not Found");

            string wwwPath = this.Environment.WebRootPath;
            string contentPath = this.Environment.ContentRootPath;
           
            string rootFolder = Path.Combine(contentPath, "UploadExcels");
            if (!Directory.Exists(rootFolder))
            {
                Directory.CreateDirectory(rootFolder);
            }
            
            var fileName = file.FileName;
            var filePath = Path.Combine(rootFolder, fileName);
            var fileLocation = new FileInfo(filePath);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(fileStream);
            }

            // if (file.Length <= 0)
            //     return new Result(false, 0, "File Not Found");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;



            FileStreamResult xmlDoc;
            using (ExcelPackage package = new ExcelPackage(fileLocation))
            {
                var sheet = package.Workbook.Worksheets.FirstOrDefault();

                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();

                int totalRows = workSheet.Dimension.Rows;
                var rowLength = workSheet.Dimension.End.Row;


                var entryNumber = getEntryNumber();

                xmlDoc = importToXML(workSheet);


            }
            
            
            return xmlDoc;
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        public string getEntryNumber()
        {

            //string connectionString =
            // "Data Source=(local);Initial Catalog=Northwind;"
            // + "Integrated Security=true";

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            var dagbknr = "202";
            string entryNumber = "" ;
            string queryString =
                "select max (bkstnr) from gbkmut ";
            //   + "where dagbknr = @dagbknr ";

            string queryString1 =
               "select top 1 vatnumber,crdnr,debnr from cicmpy";

           
            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@dagbknr", dagbknr);

                SqlCommand command1 = new SqlCommand(queryString1, connection);

              
                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        entryNumber = reader[0].ToString();
                    }

                    reader.Close();

                    SqlDataReader reader1 = command1.ExecuteReader();
                    while (reader1.Read())
                    {
                        entryNumber = reader1[1].ToString();
                    }
                    reader1.Close();
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                }
                //Console.ReadLine();
            }

            return entryNumber;
        }

        public string getInvoiceNumber()
        {
            
            //+1
            
            //741011 W K D


            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");


            var dagbknr = "202";
            string invoiceNumber = "";
            //select faktuurnr from gbkmut where dagbknr='202' დაიწყება 802-ით '80200001'
            string queryString =
                "select faktuurnr from gbkmut where faktuurnr like '%802' order by faktuurnr desc";
            

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        invoiceNumber = reader[0].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                }
                //Console.ReadLine();
            }

            if(invoiceNumber != "")
            {
                var nextInvoiceNumber = Int32.Parse(invoiceNumber);
                // nextInvoiceNumber++;
                return nextInvoiceNumber.ToString();
            }
            return "80200001";
        }


        public string getCreditorFromDB(string identityNumber)
        {

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");


            string crdnr="", debnr = "";
           
            string queryString = "select top 1 crdnr,debnr from cicmpy where vatnumber = @identityNumber ";


            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@identityNumber", identityNumber);

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        crdnr = reader[0].ToString();
                        debnr = reader[1].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                }
                //Console.ReadLine();
            }

            if(crdnr != "")
            {
                return crdnr;
            }
            return debnr;
        }

        public FileStreamResult importToXML(ExcelWorksheet workSheet)
        {
            MemoryStream ms = new MemoryStream();
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;


            using (XmlWriter xw = XmlWriter.Create(ms, xws))
            {
                XmlDocument doc = new XmlDocument();
                XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                doc.AppendChild(docNode);
                XmlElement employeeDataNode = doc.CreateElement("eExact");
                (employeeDataNode).SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                (employeeDataNode).SetAttribute("xsi:noNamespaceSchemaLocation", "eExact-Schema.xsd");
                doc.AppendChild(employeeDataNode);

                //GLEntries
                XmlNode GLEntriesNode = doc.CreateElement("GLEntries");
                doc.DocumentElement.AppendChild(GLEntriesNode);

                var GLEntryNode = getGLEntryNode(doc, workSheet);

                var rowLength = workSheet.Dimension.End.Row;

               

                var invoiceNumber = getInvoiceNumber();

                XmlNode BankStatement = doc.CreateElement("BankStatement");
                ((XmlElement)BankStatement).SetAttribute("number", "24020023");

                XmlNode Date = doc.CreateElement("Date");
                Date.AppendChild(doc.CreateTextNode("2022-02-03"));
                BankStatement.AppendChild(Date);

                XmlNode GLOffset = doc.CreateElement("GLOffset");
                ((XmlElement)GLOffset).SetAttribute("code", "   129000");
                BankStatement.AppendChild(GLOffset);

                var comIndex = 0;
                var invNumber = Int32.Parse(invoiceNumber);
                for (int i = 14; i <= rowLength; i++)
                {
                    var j = i;
                    if ( workSheet.Cells[i, 7].Value.ToString() == "COM" )
                    {
                        if (comIndex == 0)
                        {
                            comIndex = i;
                        }
                        else
                        {
                           // worksheet.Cells[i, 4].Value?.ToString())
                           //doc.SelectSingleNode()
                        }
                    }
                    var commonId = Guid.NewGuid();
                    invNumber++;
                    var FinEntryLine = getFinEntryLine(j, doc, workSheet, invNumber.ToString(), commonId);
                    GLEntryNode.AppendChild(FinEntryLine);

                    // var BankStatementLine = getBankStatement(i, doc, workSheet, commonId);
                    // BankStatement.AppendChild(BankStatementLine);
                }

                // var PaymentTerms = getPaymentTerms(doc);
                // GLEntryNode.AppendChild(PaymentTerms);
                
                

                // GLEntryNode.AppendChild(BankStatement);


                GLEntriesNode.AppendChild(GLEntryNode);

                doc.WriteTo(xw);

            }
            ms.Position = 0;
            return File(ms, "text/xml", "Sample.xml");

        }


        public string getGLAccountInEntryLine(string code)
        {
            if (code == "COM")
            {
                return "741011";
            }
            if (code == "CCO")
            {
                return "129000";
            }

            //თუ დებიტორი მაშინ  return '141010'
            return "311010";
        }
        public XmlNode getBankStatement(int i,XmlDocument doc, ExcelWorksheet worksheet, Guid commonId)
        {
          

            //////////////////////////////////////////////////////////////ციკლი
            XmlNode BankStatementLine = doc.CreateElement("BankStatementLine");
            ((XmlElement)BankStatementLine).SetAttribute("type", "Z");
            ((XmlElement)BankStatementLine).SetAttribute("termType", "S");
            ((XmlElement)BankStatementLine).SetAttribute("status", "J");
            ((XmlElement)BankStatementLine).SetAttribute("entry", "24020023");
            ((XmlElement)BankStatementLine).SetAttribute("lineNo", (i-13).ToString());
            //საერთო paymant banktransactionId
            ((XmlElement)BankStatementLine).SetAttribute("ID", String.Format("{{{0}}}", commonId.ToString()));
            ((XmlElement)BankStatementLine).SetAttribute("statementType", "B");
            ((XmlElement)BankStatementLine).SetAttribute("paymentType", "B");

            XmlNode Description = doc.CreateElement("Description");
            Description.AppendChild(doc.CreateTextNode(worksheet.Cells[i, 6].Value.ToString()));
            BankStatementLine.AppendChild(Description);

            var d = worksheet.Cells[i, 1].Value.ToString();
            var dateFormated = DateTime.Parse(d).ToString("yyyy-MM-dd");
            
            XmlNode ValueDate = doc.CreateElement("ValueDate");
            ValueDate.AppendChild(doc.CreateTextNode(dateFormated));
            BankStatementLine.AppendChild(ValueDate);

            XmlNode ReportingDate = doc.CreateElement("ReportingDate");
            ReportingDate.AppendChild(doc.CreateTextNode(dateFormated));
            BankStatementLine.AppendChild(ReportingDate);

            XmlNode StatementDate = doc.CreateElement("StatementDate");
            StatementDate.AppendChild(doc.CreateTextNode(dateFormated));
            BankStatementLine.AppendChild(StatementDate);

            var gLAccountCode = getGLAccountInEntryLine(worksheet.Cells[i, 7].Value.ToString());
            
            XmlNode GLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)GLAccount).SetAttribute("code", gLAccountCode);
            if (gLAccountCode == "741011")
            {
                ((XmlElement)GLAccount).SetAttribute("type", "W");
                ((XmlElement)GLAccount).SetAttribute("subtype", "K");
                ((XmlElement)GLAccount).SetAttribute("side", "D");
            }
            else
            {
                ((XmlElement)GLAccount).SetAttribute("type", "B");
                ((XmlElement)GLAccount).SetAttribute("subtype", "B");
                ((XmlElement)GLAccount).SetAttribute("side", "D");
            }
           

            XmlNode GLDescription = doc.CreateElement("Description");
            GLDescription.AppendChild(doc.CreateTextNode("&#1031;^&#1036;^&#1110;&#1027;&#164;&#166;&#1106;&#1029;&#1031; `^&#1108;&#1026;&#1107; GEL 3406000029"));
            GLAccount.AppendChild(GLDescription);

            BankStatementLine.AppendChild(GLAccount);

            XmlNode OwnBankAccount = doc.CreateElement("OwnBankAccount");
            ((XmlElement)OwnBankAccount).SetAttribute("code", "202");
            ((XmlElement)OwnBankAccount).SetAttribute("type", "R");

            XmlNode OwnBankAccountDescription = doc.CreateElement("Description");
            OwnBankAccountDescription.AppendChild(doc.CreateTextNode("&#1031;^&#1036;.`^&#1108;&#1026;&#1107; GEL 3406000029"));
            OwnBankAccount.AppendChild(OwnBankAccountDescription);

            XmlNode OwnBankAccountCurrency = doc.CreateElement("Currency");
            ((XmlElement)OwnBankAccountCurrency).SetAttribute("code", "GEL");

            OwnBankAccount.AppendChild(OwnBankAccountCurrency);

            XmlNode OwnBankAccountJournal = doc.CreateElement("Journal");
            ((XmlElement)OwnBankAccountJournal).SetAttribute("code", "202");
            ((XmlElement)OwnBankAccountJournal).SetAttribute("type", "B");

            OwnBankAccount.AppendChild(OwnBankAccountJournal);

            XmlNode OwnBankAccountGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)OwnBankAccountGLAccount).SetAttribute("code", "  121003");
            ((XmlElement)OwnBankAccountGLAccount).SetAttribute("type", "B");
            ((XmlElement)OwnBankAccountGLAccount).SetAttribute("subtype", "B");
            ((XmlElement)OwnBankAccountGLAccount).SetAttribute("side", "D");

            XmlNode OwnBankAccountGLAccountGLDescription = doc.CreateElement("Description");
            OwnBankAccountGLAccountGLDescription.AppendChild(doc.CreateTextNode("&#1031;^&#1036;^&#1110;&#1027;&#164;&#166;&#1106;&#1029;&#1031; `^&#1108;&#1026;&#1107; GEL 3406000029"));
            OwnBankAccountGLAccount.AppendChild(OwnBankAccountGLAccountGLDescription);

            OwnBankAccount.AppendChild(OwnBankAccountGLAccount);

            XmlNode GLPaymentInTransit = doc.CreateElement("GLPaymentInTransit");
            ((XmlElement)GLPaymentInTransit).SetAttribute("code", "   999001");
            OwnBankAccount.AppendChild(GLPaymentInTransit);

            XmlNode Country = doc.CreateElement("Country");
            ((XmlElement)Country).SetAttribute("code", "GE");
            OwnBankAccount.AppendChild(Country);

            XmlNode BankName = doc.CreateElement("BankName");
            BankName.AppendChild(doc.CreateTextNode("Other banks"));
            OwnBankAccount.AppendChild(BankName);

            XmlNode BankCreditor = doc.CreateElement("BankCreditor");
            BankCreditor.AppendChild(doc.CreateTextNode("                 294"));
            OwnBankAccount.AppendChild(BankCreditor);

            BankStatementLine.AppendChild(OwnBankAccount);

            XmlNode BankAccount = doc.CreateElement("BankAccount");
            ((XmlElement)BankAccount).SetAttribute("code", "TEST");

            XmlNode BankAccountType = doc.CreateElement("BankAccountType");
            ((XmlElement)BankAccountType).SetAttribute("code", "DEF");

            XmlNode BankAccountTypeDescription = doc.CreateElement("Description");
            BankAccountTypeDescription.AppendChild(doc.CreateTextNode(""));

            BankAccountType.AppendChild(BankAccountTypeDescription);
            BankAccount.AppendChild(BankAccountType);

            XmlNode Currency = doc.CreateElement("Currency");
            ((XmlElement)Currency).SetAttribute("code", "");

            BankAccount.AppendChild(Currency);
            BankStatementLine.AppendChild(BankAccount);

            XmlNode Creditor = doc.CreateElement("Creditor");
            var cr = getCreditorCode(worksheet.Cells[i, 7].Value.ToString(), worksheet.Cells[i, 16].Value.ToString());
            ((XmlElement)Creditor).SetAttribute("code", cr);
            ((XmlElement)Creditor).SetAttribute("number", cr);
            BankStatementLine.AppendChild(Creditor);

            XmlNode TransactionNumber = doc.CreateElement("TransactionNumber");
            TransactionNumber.AppendChild(doc.CreateTextNode("913"));
            BankStatementLine.AppendChild(TransactionNumber);

            //////////////
            XmlNode Amount = doc.CreateElement("Amount");

            XmlNode AmountCurrency = doc.CreateElement("Currency");
            ((XmlElement)AmountCurrency).SetAttribute("code", "GEL");

            Amount.AppendChild(AmountCurrency);

            XmlNode Debit = doc.CreateElement("Debit");
            Debit.AppendChild(doc.CreateTextNode("0.0"));
            Amount.AppendChild(Debit);

            XmlNode Credit = doc.CreateElement("Credit");
            Credit.AppendChild(doc.CreateTextNode("19.89"));
            Amount.AppendChild(Credit);


            BankStatementLine.AppendChild(Amount);


            XmlNode ForeignAmount = doc.CreateElement("ForeignAmount");

            XmlNode ForeignAmountCurrency = doc.CreateElement("Currency");
            ((XmlElement)ForeignAmountCurrency).SetAttribute("code", "GEL");

            ForeignAmount.AppendChild(ForeignAmountCurrency);

            XmlNode ForeignAmountDebit = doc.CreateElement("Debit");
            ForeignAmountDebit.AppendChild(doc.CreateTextNode("0.0"));
            ForeignAmount.AppendChild(ForeignAmountDebit);

            XmlNode ForeignAmountCredit = doc.CreateElement("Credit");
            ForeignAmountCredit.AppendChild(doc.CreateTextNode("19.89"));
            ForeignAmount.AppendChild(ForeignAmountCredit);

            XmlNode Rate = doc.CreateElement("Rate");
            Rate.AppendChild(doc.CreateTextNode("1"));

            ForeignAmount.AppendChild(Rate);

            BankStatementLine.AppendChild(ForeignAmount);
            
            ///
            
            XmlNode Reference = doc.CreateElement("Reference");
            Reference.AppendChild(doc.CreateTextNode(""));
            
            BankStatementLine.AppendChild(Reference);
            
            XmlNode YourRef = doc.CreateElement("YourRef");
            YourRef.AppendChild(doc.CreateTextNode("11207821"));
            
            BankStatementLine.AppendChild(YourRef);
            
            XmlNode InvoiceNumber = doc.CreateElement("InvoiceNumber");
            InvoiceNumber.AppendChild(doc.CreateTextNode("11207821"));
            
            BankStatementLine.AppendChild(InvoiceNumber);
            
            XmlNode IsBlocked = doc.CreateElement("IsBlocked");
            IsBlocked.AppendChild(doc.CreateTextNode("0"));
            
            BankStatementLine.AppendChild(IsBlocked);
            
            XmlNode PaymentTermIDs = doc.CreateElement("PaymentTermIDs");
            XmlNode PaymentTermID = doc.CreateElement("PaymentTermID");
            PaymentTermID.AppendChild(doc.CreateTextNode("{5E3224EF-910E-43C1-9381-97E22C43363E}"));
            PaymentTermIDs.AppendChild(PaymentTermID);
            BankStatementLine.AppendChild(PaymentTermIDs);
            
            XmlNode BankStatementLineGLOffset = doc.CreateElement("GLOffset");
            ((XmlElement)BankStatementLineGLOffset).SetAttribute("code", "   000002");
            
            BankStatementLine.AppendChild(BankStatementLineGLOffset);
            

            //////////////////////////////////////////////////////////////ციკლი

            return BankStatementLine;

        }

        public XmlNode getPaymentTerms(XmlDocument doc)
        {
            // დაიჯამება და ერთი იქნება თუ COM არის
            //დებეტი -> კრედიტში
            //რეფერენს yourreferenc 802
            XmlNode PaymentTerms = doc.CreateElement("PaymentTerms");

            //GLEntry
            XmlNode PaymentTerm = doc.CreateElement("PaymentTerm");
            ((XmlElement)PaymentTerm).SetAttribute("type", "T");
            ((XmlElement)PaymentTerm).SetAttribute("termType", "W");
            ((XmlElement)PaymentTerm).SetAttribute("status", "J");
            ((XmlElement)PaymentTerm).SetAttribute("entry", "24020023");
            ((XmlElement)PaymentTerm).SetAttribute("ID", "{5E3224EF-910E-43C1-9381-97E22C43363E}");
            ((XmlElement)PaymentTerm).SetAttribute("matchID", "{AFB9FC8A-2CE1-422D-ABB0-D4E15F899020}");
            ((XmlElement)PaymentTerm).SetAttribute("paymentType", "B");
            ((XmlElement)PaymentTerm).SetAttribute("paymentMethod", "T");

            XmlNode Description = doc.CreateElement("Description");
            Description.AppendChild(doc.CreateTextNode("`^&#1108;&#1026;&#1107;&#1031; &#1031;^&#1026;&#1029;&#1028;&#1107;&#1031;&#1107;&#1029;"));
            PaymentTerm.AppendChild(Description);

            XmlNode GLOffset = doc.CreateElement("GLOffset");
            ((XmlElement)GLOffset).SetAttribute("code", "   000002");
            PaymentTerm.AppendChild(GLOffset);

            XmlNode OwnBankAccount = doc.CreateElement("OwnBankAccount");
            ((XmlElement)OwnBankAccount).SetAttribute("code", "202");
            ((XmlElement)OwnBankAccount).SetAttribute("type", "R");

            XmlNode OwnBankAccountDescription = doc.CreateElement("Description");
            OwnBankAccountDescription.AppendChild(doc.CreateTextNode("&#1031;^&#1036;.`^&#1108;&#1026;&#1107; GEL 3406000029"));
            OwnBankAccount.AppendChild(OwnBankAccountDescription);

            XmlNode Currency = doc.CreateElement("Currency");
            ((XmlElement)Currency).SetAttribute("code", "GEL");
            OwnBankAccount.AppendChild(Currency);

            XmlNode Journal = doc.CreateElement("Journal");
            ((XmlElement)Journal).SetAttribute("code", "202");
            ((XmlElement)Journal).SetAttribute("type", "B");

            OwnBankAccount.AppendChild(Journal);

            XmlNode GLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)GLAccount).SetAttribute("code", "   121003");
            ((XmlElement)GLAccount).SetAttribute("side", "D");
            ((XmlElement)GLAccount).SetAttribute("type", "B");
            ((XmlElement)GLAccount).SetAttribute("subtype", "B");

            XmlNode OwnBankAccountGLAccountDescription = doc.CreateElement("Description");
            OwnBankAccountGLAccountDescription.AppendChild(doc.CreateTextNode("&#1031;^&#1036;^&#1110;&#1027;&#164;&#166;&#1106;&#1029;&#1031; `^&#1108;&#1026;&#1107; GEL 3406000029"));
            GLAccount.AppendChild(OwnBankAccountGLAccountDescription);

            OwnBankAccount.AppendChild(GLAccount);


            XmlNode GLPaymentInTransit = doc.CreateElement("GLPaymentInTransit");
            ((XmlElement)GLPaymentInTransit).SetAttribute("code", "   999001");
            OwnBankAccount.AppendChild(GLPaymentInTransit);

            XmlNode Country = doc.CreateElement("Country");
            ((XmlElement)Country).SetAttribute("code", "GE");
            OwnBankAccount.AppendChild(Country);

            XmlNode BankName = doc.CreateElement("BankName");
            BankName.AppendChild(doc.CreateTextNode("Other banks"));
            OwnBankAccount.AppendChild(BankName);

            XmlNode BankCreditor = doc.CreateElement("BankCreditor");
            BankCreditor.AppendChild(doc.CreateTextNode("                 294"));
            OwnBankAccount.AppendChild(BankCreditor);

            PaymentTerm.AppendChild(OwnBankAccount);

            XmlNode Creditor = doc.CreateElement("Creditor");
            ((XmlElement)Creditor).SetAttribute("code", "                   4");
            ((XmlElement)Creditor).SetAttribute("number", "4");

            PaymentTerm.AppendChild(Creditor);

            XmlNode TransactionNumber = doc.CreateElement("TransactionNumber");
            TransactionNumber.AppendChild(doc.CreateTextNode("913"));
            PaymentTerm.AppendChild(TransactionNumber);

            XmlNode Amount = doc.CreateElement("Amount");

            XmlNode AmountCurrency = doc.CreateElement("Currency");
            ((XmlElement)AmountCurrency).SetAttribute("code", "GEL");

            Amount.AppendChild(AmountCurrency);

            XmlNode Debit = doc.CreateElement("Debit");
            Debit.AppendChild(doc.CreateTextNode("0.0"));
            Amount.AppendChild(Debit);

            XmlNode Credit = doc.CreateElement("Credit");
            Credit.AppendChild(doc.CreateTextNode("19.89"));
            Amount.AppendChild(Credit);

            XmlNode VAT = doc.CreateElement("VAT");
            ((XmlElement)VAT).SetAttribute("code", "0");

            Amount.AppendChild(VAT);

            PaymentTerm.AppendChild(Amount);


            XmlNode ForeignAmount = doc.CreateElement("ForeignAmount");

            XmlNode ForeignAmountCurrency = doc.CreateElement("Currency");
            ((XmlElement)ForeignAmountCurrency).SetAttribute("code", "GEL");

            ForeignAmount.AppendChild(ForeignAmountCurrency);

            XmlNode ForeignAmountDebit = doc.CreateElement("Debit");
            ForeignAmountDebit.AppendChild(doc.CreateTextNode("0.0"));
            ForeignAmount.AppendChild(ForeignAmountDebit);

            XmlNode ForeignAmountCredit = doc.CreateElement("Credit");
            ForeignAmountCredit.AppendChild(doc.CreateTextNode("19.89"));
            ForeignAmount.AppendChild(ForeignAmountCredit);

            XmlNode Rate = doc.CreateElement("Rate");
            Rate.AppendChild(doc.CreateTextNode("1"));

            ForeignAmount.AppendChild(Rate);

            PaymentTerm.AppendChild(ForeignAmount);

            XmlNode PaymentCondition = doc.CreateElement("PaymentCondition");
            ((XmlElement)PaymentCondition).SetAttribute("code", "");

            PaymentTerm.AppendChild(PaymentCondition);

            XmlNode DaysToPayment = doc.CreateElement("DaysToPayment");
            DaysToPayment.AppendChild(doc.CreateTextNode("0"));
            PaymentTerm.AppendChild(DaysToPayment);

            XmlNode Percentage = doc.CreateElement("Percentage");
            Percentage.AppendChild(doc.CreateTextNode("1"));
            PaymentTerm.AppendChild(Percentage);

            XmlNode Reference = doc.CreateElement("Reference");
            Reference.AppendChild(doc.CreateTextNode("11207821"));
            PaymentTerm.AppendChild(Reference);

            XmlNode YourRef = doc.CreateElement("YourRef");
            YourRef.AppendChild(doc.CreateTextNode("11207821"));
            PaymentTerm.AppendChild(YourRef);

            XmlNode InvoiceNumber = doc.CreateElement("InvoiceNumber");
            InvoiceNumber.AppendChild(doc.CreateTextNode("11207821"));
            PaymentTerm.AppendChild(InvoiceNumber);

            XmlNode InvoiceDate = doc.CreateElement("InvoiceDate");
            InvoiceDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            PaymentTerm.AppendChild(InvoiceDate);

            XmlNode InvoiceDueDate = doc.CreateElement("InvoiceDueDate");
            InvoiceDueDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            PaymentTerm.AppendChild(InvoiceDueDate);

            XmlNode ProcessingDate = doc.CreateElement("ProcessingDate");
            ProcessingDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            PaymentTerm.AppendChild(ProcessingDate);

            XmlNode ReportingDate = doc.CreateElement("ReportingDate");
            ReportingDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            PaymentTerm.AppendChild(ReportingDate);

            XmlNode Resource = doc.CreateElement("Resource");
            ((XmlElement)Resource).SetAttribute("number", "0");

            PaymentTerm.AppendChild(Resource);

            XmlNode Journalization = doc.CreateElement("Journalization");

            XmlNode JournalizationResource = doc.CreateElement("Resource");
            ((XmlElement)JournalizationResource).SetAttribute("number", "0");
            Journalization.AppendChild(JournalizationResource);

            XmlNode JournalizationDate = doc.CreateElement("Date");
            JournalizationDate.AppendChild(doc.CreateTextNode("2022-02-04"));
            Journalization.AppendChild(JournalizationDate);

            PaymentTerm.AppendChild(Journalization);

            XmlNode IsBlocked = doc.CreateElement("IsBlocked");
            IsBlocked.AppendChild(doc.CreateTextNode("0"));
            PaymentTerm.AppendChild(IsBlocked);

            PaymentTerms.AppendChild(PaymentTerm);
            return PaymentTerms;

        }

        public XmlNode getGLEntryNode(XmlDocument doc,  ExcelWorksheet worksheet)
        {


            //GLEntry
            XmlNode GLEntryNode = doc.CreateElement("GLEntry");
            ((XmlElement)GLEntryNode).SetAttribute("entry", "24020023");
            ((XmlElement)GLEntryNode).SetAttribute("status", "E");

            XmlNode Division = doc.CreateElement("Division");
            ((XmlElement)Division).SetAttribute("code", "150");
            GLEntryNode.AppendChild(Division);

            XmlNode DocumentDate = doc.CreateElement("DocumentDate");
            var d = worksheet.Cells[14, 1].Value.ToString(); //პირველივე თარიღი
            var dateFormated = DateTime.Parse(d).ToString("yyyy-MM-dd");
            DocumentDate.AppendChild(doc.CreateTextNode(dateFormated));
            GLEntryNode.AppendChild(DocumentDate);

            XmlNode Journal = doc.CreateElement("Journal");
            ((XmlElement)Journal).SetAttribute("code", "202");
            ((XmlElement)Journal).SetAttribute("type", "B");


            XmlNode Description = doc.CreateElement("Description");
            Description.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            Journal.AppendChild(Description);

            XmlNode GLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)GLAccount).SetAttribute("code", "  121003");
            ((XmlElement)GLAccount).SetAttribute("type", "B");
            ((XmlElement)GLAccount).SetAttribute("subtype", "B");
            ((XmlElement)GLAccount).SetAttribute("side", "D");

            XmlNode GLDescription = doc.CreateElement("Description");
            GLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            GLAccount.AppendChild(GLDescription);


            XmlNode GLPaymentInTransit = doc.CreateElement("GLPaymentInTransit");
            ((XmlElement)GLPaymentInTransit).SetAttribute("code", "999001");
            ((XmlElement)GLPaymentInTransit).SetAttribute("type", "B");
            ((XmlElement)GLPaymentInTransit).SetAttribute("subtype", "B");
            ((XmlElement)GLPaymentInTransit).SetAttribute("side", "D");

            XmlNode GLPaymentInTransitDescription = doc.CreateElement("Description");
            GLPaymentInTransitDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            GLPaymentInTransit.AppendChild(GLPaymentInTransitDescription);


            Journal.AppendChild(GLAccount);
            Journal.AppendChild(GLPaymentInTransit);

            GLEntryNode.AppendChild(Journal);


            XmlNode Costcenter = doc.CreateElement("Costcenter");
            ((XmlElement)Costcenter).SetAttribute("code", "001CC001");

            XmlNode CostcenterDescription = doc.CreateElement("Description");
            CostcenterDescription.AppendChild(doc.CreateTextNode("Default cost center"));
            Costcenter.AppendChild(CostcenterDescription);

            XmlNode CostcenterGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)CostcenterGLAccount).SetAttribute("code", "719990");
            ((XmlElement)CostcenterGLAccount).SetAttribute("type", "D");
            ((XmlElement)CostcenterGLAccount).SetAttribute("subtype", "W");
            ((XmlElement)CostcenterGLAccount).SetAttribute("side", "K");

            XmlNode CostcenterGLAccountGLDescription = doc.CreateElement("Description");
            CostcenterGLAccountGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            CostcenterGLAccount.AppendChild(CostcenterGLAccountGLDescription);

            Costcenter.AppendChild(CostcenterGLAccount);

            XmlNode GLOffset = doc.CreateElement("GLOffset");
            ((XmlElement)GLOffset).SetAttribute("code", "719990");
            ((XmlElement)GLOffset).SetAttribute("type", "D");
            ((XmlElement)GLOffset).SetAttribute("subtype", "W");
            ((XmlElement)GLOffset).SetAttribute("side", "K");

            XmlNode GLOffsetDescription = doc.CreateElement("Description");
            GLOffsetDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            GLOffset.AppendChild(GLOffsetDescription);

            Costcenter.AppendChild(GLOffset);
            GLEntryNode.AppendChild(Costcenter);

            XmlNode Amount = doc.CreateElement("Amount");

            XmlNode Currency = doc.CreateElement("Currency");
            ((XmlElement)Currency).SetAttribute("code", "  GEL");
            XmlNode Value = doc.CreateElement("Value");
            Value.AppendChild(doc.CreateTextNode("0"));
            Amount.AppendChild(Currency);
            Amount.AppendChild(Value);

            GLEntryNode.AppendChild(Amount);


            XmlNode ForeignAmount = doc.CreateElement("ForeignAmount");
            XmlNode ForeignAmountCurrency = doc.CreateElement("Currency");
            ((XmlElement)ForeignAmountCurrency).SetAttribute("code", "  GEL");
            XmlNode ForeignAmountValue = doc.CreateElement("Value");
            ForeignAmountValue.AppendChild(doc.CreateTextNode("0"));
            ForeignAmount.AppendChild(ForeignAmountCurrency);
            ForeignAmount.AppendChild(ForeignAmountValue);

            GLEntryNode.AppendChild(ForeignAmount);

            return GLEntryNode;
        }

        public string getCreditorCode(string code, string identityNymber)
        {
            // CCO -> კრედიტორი 3
            // COM -> კრედიტორი 4
            // სვა შემთხვევაში select vatnumber,crdnr,debnr from cicmpy where VatNumber = '102189454'(მიმღების საიდენთიფიკაციო კოდი)
            // რომელიც null არაა იმით შეივსება
            if (code == "COM")
            {
                return "4";
            }
            if (code == "CCO")
            {
                return "3";
            }

            var cr = getCreditorFromDB(identityNymber);

            return cr;

        }
        public XmlNode getFinEntryLine(int i, XmlDocument doc, ExcelWorksheet worksheet,string  invoiceNumber, Guid commonId)
        {
            
            //COM დაჯამდება, რეფერენსები იქნება საერთო
            XmlNode FinEntryLine = doc.CreateElement("FinEntryLine");
            ((XmlElement)FinEntryLine).SetAttribute("number", (i-13).ToString());
            ((XmlElement)FinEntryLine).SetAttribute("type", "N");
            ((XmlElement)FinEntryLine).SetAttribute("subtype", "Z");

            XmlNode Date = doc.CreateElement("Date");
            var d = worksheet.Cells[i, 1].Value.ToString();
            var dateFormated = DateTime.Parse(d).ToString("yyyy-MM-dd");
            Date.AppendChild(doc.CreateTextNode(dateFormated));
            FinEntryLine.AppendChild(Date);

            XmlNode FinYear = doc.CreateElement("FinYear");
            ((XmlElement)FinYear).SetAttribute("number", DateTime.Parse(d).Year.ToString());
            FinEntryLine.AppendChild(FinYear);
            
            XmlNode FinPeriod = doc.CreateElement("FinPeriod");
            ((XmlElement)FinPeriod).SetAttribute("number", DateTime.Parse(d).Month.ToString());
            FinEntryLine.AppendChild(FinPeriod);

            var gLAccountCode = getGLAccountInEntryLine(worksheet.Cells[i, 7].Value.ToString());
            
            XmlNode FinEntryLineGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)FinEntryLineGLAccount).SetAttribute("code", gLAccountCode);
            if (gLAccountCode == "741011")
            {
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "W");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "K");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            }
            else
            {
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "B");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            }

            XmlNode FinEntryLineGLDescription = doc.CreateElement("Description");
            FinEntryLineGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            FinEntryLineGLAccount.AppendChild(FinEntryLineGLDescription);
            FinEntryLine.AppendChild(FinEntryLineGLAccount);


            XmlNode FinEntryLineDescription = doc.CreateElement("Description");
            FinEntryLineDescription.AppendChild(doc.CreateTextNode(worksheet.Cells[i, 6].Value.ToString()));
            FinEntryLine.AppendChild(FinEntryLineDescription);


            //-----------------------------
            XmlNode FinEntryLineCostcenter = doc.CreateElement("Costcenter");
            ((XmlElement)FinEntryLineCostcenter).SetAttribute("code", "001CC001");

            XmlNode FinEntryLineCostcenterDescription = doc.CreateElement("Description");
            FinEntryLineCostcenterDescription.AppendChild(doc.CreateTextNode("Default cost center"));
            FinEntryLineCostcenter.AppendChild(FinEntryLineCostcenterDescription);

            XmlNode FinEntryLineCostcenterGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("code", "   719990");
            ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("type", "D");
            ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("subtype", "W");
            ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("side", "K");

            XmlNode FinEntryLineCostcenterGLAccountGLDescription = doc.CreateElement("Description");
            FinEntryLineCostcenterGLAccountGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            FinEntryLineCostcenterGLAccount.AppendChild(FinEntryLineCostcenterGLAccountGLDescription);

            FinEntryLineCostcenter.AppendChild(FinEntryLineCostcenterGLAccount);

            XmlNode FinEntryLineGLOffset = doc.CreateElement("GLOffset");
            ((XmlElement)FinEntryLineGLOffset).SetAttribute("code", " 719990");
            ((XmlElement)FinEntryLineGLOffset).SetAttribute("type", "D");
            ((XmlElement)FinEntryLineGLOffset).SetAttribute("subtype", "W");
            ((XmlElement)FinEntryLineGLOffset).SetAttribute("side", "K");

            XmlNode FinEntryLineGLOffsetDescription = doc.CreateElement("Description");
            FinEntryLineGLOffsetDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            FinEntryLineGLOffset.AppendChild(FinEntryLineGLOffsetDescription);

            FinEntryLineCostcenter.AppendChild(FinEntryLineGLOffset);
            FinEntryLine.AppendChild(FinEntryLineCostcenter);

            //-----------------------------

           
            var creditorCode = getCreditorCode(worksheet.Cells[i, 7].Value.ToString(), worksheet.Cells[i, 16].Value.ToString());
            XmlNode Creditor = doc.CreateElement("Creditor");
            ((XmlElement)Creditor).SetAttribute("code", creditorCode);
            ((XmlElement)Creditor).SetAttribute("number", creditorCode);
            ((XmlElement)Creditor).SetAttribute("type", "S");

            XmlNode CreditorName = doc.CreateElement("Name");
            CreditorName.AppendChild(doc.CreateTextNode("&#1026;&#1110;&#166;~&#1107;&#1111;&#1029;&#1110;&#1107; `^&#1108;&#1026;&#1107;-&#1031;^&#1106;"));
            Creditor.AppendChild(CreditorName);


            FinEntryLine.AppendChild(Creditor);


            XmlNode Resource = doc.CreateElement("Resource");
            ((XmlElement)Resource).SetAttribute("number", "1");
            ((XmlElement)Resource).SetAttribute("code", "BMOSIDZE");

            XmlNode LastName = doc.CreateElement("LastName");
            LastName.AppendChild(doc.CreateTextNode("&#1028;&#1029;&#1031;&#1107;&#1035;&#166;"));
            Resource.AppendChild(LastName);

            XmlNode FirstName = doc.CreateElement("FirstName");
            FirstName.AppendChild(doc.CreateTextNode("`&#166;&#1031;&#1107;&#1026;"));
            Resource.AppendChild(FirstName);

            FinEntryLine.AppendChild(Resource);

            XmlNode Quantity = doc.CreateElement("Quantity");
            Quantity.AppendChild(doc.CreateTextNode("0"));
            FinEntryLine.AppendChild(Quantity);


            XmlNode FinEntryLineAmount = doc.CreateElement("Amount");

            XmlNode FinEntryLineAmountCurrency = doc.CreateElement("Currency");
            //მე-8 ხაზიდან
            ((XmlElement)FinEntryLineAmountCurrency).SetAttribute("code", worksheet.Cells[8, 3].Value.ToString());
            FinEntryLineAmount.AppendChild(FinEntryLineAmountCurrency);

            //დებეტის ველიდან
            
            XmlNode Debit = doc.CreateElement("Debit");
            Debit.AppendChild(doc.CreateTextNode(!String.IsNullOrEmpty(worksheet.Cells[i, 4].Value?.ToString()) ? worksheet.Cells[i, 4].Value?.ToString() : "0"));
            FinEntryLineAmount.AppendChild(Debit);

            //კრედიტის ველიდან
            XmlNode Credit = doc.CreateElement("Credit");
            Credit.AppendChild(doc.CreateTextNode(!String.IsNullOrEmpty(worksheet.Cells[i, 5].Value?.ToString())? worksheet.Cells[i, 5].Value?.ToString() : "0"));
            FinEntryLineAmount.AppendChild(Credit);

            XmlNode VAT = doc.CreateElement("VAT");
            ((XmlElement)VAT).SetAttribute("code", "0");
            ((XmlElement)VAT).SetAttribute("type", "B");
            ((XmlElement)VAT).SetAttribute("vattype", "N");
            ((XmlElement)VAT).SetAttribute("taxtype", "V");

            XmlNode VATDescription = doc.CreateElement("Description");
            VATDescription.AppendChild(doc.CreateTextNode("VAT 0%"));
            VAT.AppendChild(VATDescription);

            XmlNode MultiDescriptions = doc.CreateElement("MultiDescriptions");

            XmlNode MultiDescription1 = doc.CreateElement("MultiDescription");
            ((XmlElement)MultiDescription1).SetAttribute("number", "1");
            MultiDescription1.AppendChild(doc.CreateTextNode("VAT 0%"));
            MultiDescriptions.AppendChild(MultiDescription1);

            XmlNode MultiDescription2 = doc.CreateElement("MultiDescription");
            ((XmlElement)MultiDescription2).SetAttribute("number", "2");
            MultiDescription2.AppendChild(doc.CreateTextNode("VAT 0%"));
            MultiDescriptions.AppendChild(MultiDescription2);

            XmlNode MultiDescription3 = doc.CreateElement("MultiDescription");
            ((XmlElement)MultiDescription3).SetAttribute("number", "3");
            MultiDescription3.AppendChild(doc.CreateTextNode("VAT 0%"));
            MultiDescriptions.AppendChild(MultiDescription3);

            XmlNode MultiDescription4 = doc.CreateElement("MultiDescription");
            ((XmlElement)MultiDescription4).SetAttribute("number", "4");
            MultiDescription4.AppendChild(doc.CreateTextNode("VAT 0%"));
            MultiDescriptions.AppendChild(MultiDescription4);

            VAT.AppendChild(MultiDescriptions);


            XmlNode Percentage = doc.CreateElement("Percentage");
            Percentage.AppendChild(doc.CreateTextNode("0"));
            VAT.AppendChild(Percentage);

            XmlNode Charged = doc.CreateElement("Charged");
            Charged.AppendChild(doc.CreateTextNode("0"));
            VAT.AppendChild(Charged);

            XmlNode VATExemption = doc.CreateElement("VATExemption");
            VATExemption.AppendChild(doc.CreateTextNode("0"));
            VAT.AppendChild(VATExemption);

            XmlNode ExtraDutyPercentage = doc.CreateElement("ExtraDutyPercentage");
            ExtraDutyPercentage.AppendChild(doc.CreateTextNode("0"));
            VAT.AppendChild(ExtraDutyPercentage);



            XmlNode GLToPay = doc.CreateElement("GLToPay");
            ((XmlElement)GLToPay).SetAttribute("code", "   333010");
            ((XmlElement)GLToPay).SetAttribute("side", "C");
            ((XmlElement)GLToPay).SetAttribute("type", "B");
            ((XmlElement)GLToPay).SetAttribute("subtype", "C");

            XmlNode GLToPayDescription = doc.CreateElement("Description");
            GLToPayDescription.AppendChild(doc.CreateTextNode("|^~^&#1031;^&#1118;~&#166;&#1106;&#1107; ~&#1116;|0"));
            GLToPay.AppendChild(GLToPayDescription);

            VAT.AppendChild(GLToPay);

            XmlNode GLToClaim = doc.CreateElement("GLToClaim");
            ((XmlElement)GLToClaim).SetAttribute("code", "   333010");
            ((XmlElement)GLToClaim).SetAttribute("side", "C");
            ((XmlElement)GLToClaim).SetAttribute("type", "B");
            ((XmlElement)GLToClaim).SetAttribute("subtype", "C");

            XmlNode GLToClaimDescription = doc.CreateElement("Description");
            GLToClaimDescription.AppendChild(doc.CreateTextNode("|^~^&#1031;^&#1118;~&#166;&#1106;&#1107; ~&#1116;|0"));
            GLToClaim.AppendChild(GLToClaimDescription);

            VAT.AppendChild(GLToClaim);

            XmlNode VATCreditor = doc.CreateElement("Creditor");
            ((XmlElement)VATCreditor).SetAttribute("code", "        1");
            ((XmlElement)VATCreditor).SetAttribute("number", "        1");
            ((XmlElement)VATCreditor).SetAttribute("type", "S");

            XmlNode VATCreditorName = doc.CreateElement("Name");
            VATCreditorName.AppendChild(doc.CreateTextNode("VAT Creditor"));
            VATCreditor.AppendChild(VATCreditorName);

            VAT.AppendChild(VATCreditor);

            XmlNode PaymentPeriod = doc.CreateElement("PaymentPeriod");
            PaymentPeriod.AppendChild(doc.CreateTextNode("M"));

            VAT.AppendChild(PaymentPeriod);

            FinEntryLineAmount.AppendChild(VAT);
            FinEntryLine.AppendChild(FinEntryLineAmount);


            XmlNode VATTransaction = doc.CreateElement("VATTransaction");
            ((XmlElement)VATTransaction).SetAttribute("code", "0");

            XmlNode VATAmount = doc.CreateElement("VATAmount");
            VATAmount.AppendChild(doc.CreateTextNode("0"));
            VATTransaction.AppendChild(VATAmount);

            XmlNode VATBaseAmount = doc.CreateElement("VATBaseAmount");
            VATBaseAmount.AppendChild(doc.CreateTextNode("0"));
            VATTransaction.AppendChild(VATBaseAmount);

            XmlNode VATBaseAmountFC = doc.CreateElement("VATBaseAmountFC");
            VATBaseAmountFC.AppendChild(doc.CreateTextNode("0"));
            VATTransaction.AppendChild(VATBaseAmountFC);

            FinEntryLine.AppendChild(VATTransaction);



            XmlNode Payment = doc.CreateElement("Payment");

            XmlNode PaymentMethod = doc.CreateElement("PaymentMethod");
            ((XmlElement)PaymentMethod).SetAttribute("code", "B");
            Payment.AppendChild(PaymentMethod);

            XmlNode PaymentCondition = doc.CreateElement("PaymentCondition");
            ((XmlElement)PaymentCondition).SetAttribute("code", "00");

            XmlNode PaymentConditionDescription = doc.CreateElement("Description");
            ((XmlElement)PaymentConditionDescription).SetAttribute("code", "B");
            PaymentCondition.AppendChild(PaymentConditionDescription);


            Payment.AppendChild(PaymentCondition);


            XmlNode CSSDYesNo = doc.CreateElement("CSSDYesNo");
            CSSDYesNo.AppendChild(doc.CreateTextNode("K"));
            Payment.AppendChild(CSSDYesNo);

            XmlNode CSSDAmount1 = doc.CreateElement("CSSDAmount1");
            CSSDAmount1.AppendChild(doc.CreateTextNode("0"));
            Payment.AppendChild(CSSDAmount1);

            XmlNode CSSDAmount2 = doc.CreateElement("CSSDAmount2");
            CSSDAmount2.AppendChild(doc.CreateTextNode("0"));
            Payment.AppendChild(CSSDAmount2);

            XmlNode InvoiceNumber = doc.CreateElement("InvoiceNumber");
            //select faktuurnr from gbkmut where dagbknr='202' დაიწყება 802-ით '80200001'

            InvoiceNumber.AppendChild(doc.CreateTextNode(invoiceNumber));
            Payment.AppendChild(InvoiceNumber);

            XmlNode BankTransactionID = doc.CreateElement("BankTransactionID");
            //new common Guid

            BankTransactionID.AppendChild(doc.CreateTextNode(String.Format("{{{0}}}", commonId.ToString())));
            Payment.AppendChild(BankTransactionID);


            FinEntryLine.AppendChild(Payment);

            XmlNode Delivery = doc.CreateElement("Delivery");

            XmlNode DeliveryDate = doc.CreateElement("Date");
            DeliveryDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            Delivery.AppendChild(DeliveryDate);


            FinEntryLine.AppendChild(Delivery);


            XmlNode FinReferences = doc.CreateElement("FinReferences");
            ((XmlElement)FinReferences).SetAttribute("TransactionOrigin", "P");

            XmlNode UniquePostingNumber = doc.CreateElement("UniquePostingNumber");
            UniquePostingNumber.AppendChild(doc.CreateTextNode("0"));
            FinReferences.AppendChild(UniquePostingNumber);

            XmlNode YourRef = doc.CreateElement("YourRef");
            YourRef.AppendChild(doc.CreateTextNode(""));
            FinReferences.AppendChild(YourRef);

            XmlNode FinReferencesDocumentDate = doc.CreateElement("DocumentDate");
            FinReferencesDocumentDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            FinReferences.AppendChild(FinReferencesDocumentDate);

            FinEntryLine.AppendChild(FinReferences);




            XmlNode FinEntryLineDiscount = doc.CreateElement("Discount");

            XmlNode FinEntryLinePercentage = doc.CreateElement("Percentage");
            FinEntryLinePercentage.AppendChild(doc.CreateTextNode("0"));
            FinEntryLineDiscount.AppendChild(Percentage);

            FinEntryLine.AppendChild(FinEntryLineDiscount);


            XmlNode FreeFields = doc.CreateElement("FreeFields");

            XmlNode FreeTexts = doc.CreateElement("FreeTexts");
            XmlNode FreeText = doc.CreateElement("FreeText");
            FreeText.AppendChild(doc.CreateTextNode("00"));
            ((XmlElement)FreeText).SetAttribute("number", "3");
            FreeTexts.AppendChild(FreeText);
            FreeFields.AppendChild(FreeTexts);

            FinEntryLine.AppendChild(FreeFields);

            return FinEntryLine;
        }
    }
}
