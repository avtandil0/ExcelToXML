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

namespace ExcelToXML.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
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
            foreach (var formFile in files)
            {
                if (formFile.Length > 0)
                {
                    // full path to file in temp location
                    var filePath = Path.GetTempFileName();
                    filePaths.Add(filePath);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await formFile.CopyToAsync(stream);
                    }
                }
            }

            // process uploaded files
            // Don't rely on or trust the FileName property without validation.
            //return Ok(new { count = files.Count, size, filePaths });

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
                doc.AppendChild(employeeDataNode);

                //GLEntries
                XmlNode GLEntriesNode = doc.CreateElement("GLEntries");
                doc.DocumentElement.AppendChild(GLEntriesNode);

                //GLEntry
                XmlNode GLEntryNode = doc.CreateElement("GLEntry");
                ((XmlElement)GLEntryNode).SetAttribute("entry", "24020023");
                ((XmlElement)GLEntryNode).SetAttribute("status", "E");

                XmlNode Division = doc.CreateElement("Division");
                Division.AppendChild(doc.CreateTextNode("150"));
                GLEntryNode.AppendChild(Division);

                XmlNode DocumentDate = doc.CreateElement("DocumentDate");
                DocumentDate.AppendChild(doc.CreateTextNode("2022-02-03"));
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
                ((XmlElement)GLPaymentInTransit).SetAttribute("code", "  121003");
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
                ((XmlElement)CostcenterGLAccount).SetAttribute("code", "  121003");
                ((XmlElement)CostcenterGLAccount).SetAttribute("type", "B");
                ((XmlElement)CostcenterGLAccount).SetAttribute("subtype", "B");
                ((XmlElement)CostcenterGLAccount).SetAttribute("side", "D");

                XmlNode CostcenterGLAccountGLDescription = doc.CreateElement("Description");
                CostcenterGLAccountGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
                CostcenterGLAccount.AppendChild(CostcenterGLAccountGLDescription);

                Costcenter.AppendChild(CostcenterGLAccount);

                XmlNode GLOffset = doc.CreateElement("GLOffset");
                ((XmlElement)GLOffset).SetAttribute("code", "  121003");
                ((XmlElement)GLOffset).SetAttribute("type", "B");
                ((XmlElement)GLOffset).SetAttribute("subtype", "B");
                ((XmlElement)GLOffset).SetAttribute("side", "D");

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


                XmlNode FinEntryLine = doc.CreateElement("FinEntryLine");
                ((XmlElement)FinEntryLine).SetAttribute("number", "1");
                ((XmlElement)FinEntryLine).SetAttribute("type", "N");
                ((XmlElement)FinEntryLine).SetAttribute("subtype", "Z");

                XmlNode Date = doc.CreateElement("Date");
                Date.AppendChild(doc.CreateTextNode("2022-02-03"));
                FinEntryLine.AppendChild(Date);

                XmlNode FinYear = doc.CreateElement("FinYear");
                ((XmlElement)FinYear).SetAttribute("number", "2022");
                FinEntryLine.AppendChild(FinYear);

                XmlNode FinEntryLineGLAccount = doc.CreateElement("GLAccount");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("code", "  121003");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "B");
                ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");

                XmlNode FinEntryLineGLDescription = doc.CreateElement("Description");
                FinEntryLineGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
                FinEntryLineGLAccount.AppendChild(FinEntryLineGLDescription);
                FinEntryLine.AppendChild(FinEntryLineGLAccount);


                XmlNode FinEntryLineDescription = doc.CreateElement("Description");
                FinEntryLineDescription.AppendChild(doc.CreateTextNode("Default cost center"));
                FinEntryLine.AppendChild(FinEntryLineDescription);


                //-----------------------------
                XmlNode FinEntryLineCostcenter = doc.CreateElement("Costcenter");
                ((XmlElement)FinEntryLineCostcenter).SetAttribute("code", "001CC001");

                XmlNode FinEntryLineCostcenterDescription = doc.CreateElement("Description");
                FinEntryLineCostcenterDescription.AppendChild(doc.CreateTextNode("Default cost center"));
                FinEntryLineCostcenter.AppendChild(FinEntryLineCostcenterDescription);

                XmlNode FinEntryLineCostcenterGLAccount = doc.CreateElement("GLAccount");
                ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("code", "  121003");
                ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("type", "B");
                ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("subtype", "B");
                ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("side", "D");

                XmlNode FinEntryLineCostcenterGLAccountGLDescription = doc.CreateElement("Description");
                FinEntryLineCostcenterGLAccountGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
                FinEntryLineCostcenterGLAccount.AppendChild(FinEntryLineCostcenterGLAccountGLDescription);

                FinEntryLineCostcenter.AppendChild(FinEntryLineCostcenterGLAccount);

                XmlNode FinEntryLineGLOffset = doc.CreateElement("GLOffset");
                ((XmlElement)FinEntryLineGLOffset).SetAttribute("code", "  121003");
                ((XmlElement)FinEntryLineGLOffset).SetAttribute("type", "B");
                ((XmlElement)FinEntryLineGLOffset).SetAttribute("subtype", "B");
                ((XmlElement)FinEntryLineGLOffset).SetAttribute("side", "D");

                XmlNode FinEntryLineGLOffsetDescription = doc.CreateElement("Description");
                FinEntryLineGLOffsetDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
                FinEntryLineGLOffset.AppendChild(FinEntryLineGLOffsetDescription);

                FinEntryLineCostcenter.AppendChild(FinEntryLineGLOffset);
                FinEntryLine.AppendChild(FinEntryLineCostcenter);

                //-----------------------------


                XmlNode Creditor = doc.CreateElement("Creditor");
                ((XmlElement)Creditor).SetAttribute("code", "                   3");
                ((XmlElement)Creditor).SetAttribute("number", "     3");
                ((XmlElement)Creditor).SetAttribute("type", "S");

                XmlNode CreditorName = doc.CreateElement("Name");
                CreditorName.AppendChild(doc.CreateTextNode("&#1026;&#1110;&#166;~&#1107;&#1111;&#1029;&#1110;&#1107; `^&#1108;&#1026;&#1107;-&#1031;^&#1106;"));
                Creditor.AppendChild(CreditorName);

                FinEntryLine.AppendChild(Creditor);
                GLEntryNode.AppendChild(FinEntryLine);

                GLEntriesNode.AppendChild(GLEntryNode);

                doc.WriteTo(xw);

            }
            ms.Position = 0;
            return File(ms, "text/xml", "Sample.xml");

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
    }
}
