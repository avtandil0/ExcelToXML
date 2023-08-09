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
using ExcelToXML.Auth;
using Microsoft.AspNetCore.Authorization;



//select dagbknr, dbo.getunicode(dagbk.oms25_0), dagbk.reknr,dbo.getunicode(grtbk.oms25_0)
//from dagbk inner join grtbk on dagbk.reknr = grtbk.reknr
//where type_dgbk = 'B'
//select dagbknr, reknr from dagbk
//getcreditr




namespace ExcelToXML.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment Environment;
        private IConfiguration Configuration;

        private readonly ITokenService _tokenService;
        private string generatedToken = null;
        private readonly IUserRepository _userRepository;



        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment _environment, IConfiguration _configuration, ITokenService tokenService, IUserRepository userRepository)
        {
            _logger = logger;
            Environment = _environment;
            Configuration = _configuration;
            _tokenService = tokenService;
            _userRepository = userRepository;


        }


        [Route("Login")]
        public IActionResult Login()
        {
            return View();
        }

        [Authorize]
        [Route("LogOut")]
        public IActionResult LogOut()
        {
            HttpContext.Session.SetString("Token", "");
            return RedirectToAction("Index");
        }

        [AllowAnonymous]
        [Route("login")]
        [HttpPost]
        public IActionResult Login(UserDTO userModel)
        {
            if (string.IsNullOrEmpty(userModel.UserName) || string.IsNullOrEmpty(userModel.Password))
            {
                ViewBag.error = "მომხმარებელი ან პაროლი არასწორია ! ";

                return View("Login");
            }

            IActionResult response = Unauthorized();
            var validUser = GetUser(userModel);

            if (validUser != null)
            {
                generatedToken = _tokenService.BuildToken(Configuration["Jwt:Key"].ToString(), Configuration["Jwt:Issuer"].ToString(),
                validUser);

                if (generatedToken != null)
                {
                    HttpContext.Session.SetString("Token", generatedToken);
                    return RedirectToAction("Index");
                }
                else
                {
                    ViewBag.error = "მომხმარებელი ან პაროლი არასწორია ! ";

                    return View("Login");
                }
            }
            else
            {
                ViewBag.error = "მომხმარებელი ან პაროლი არასწორია ! ";

                return View("Login");
            }
        }

        private UserDTO GetUser(UserDTO userModel)
        {
            //Write your code here to authenticate the user
            return _userRepository.GetUser(userModel);
        }

        public IActionResult Index()
        {
            string token = HttpContext.Session.GetString("Token");

            if (String.IsNullOrEmpty(token))
            {
                return (RedirectToAction("Login"));
            }


            var jurnals = getJurnals();
            ViewBag.jurnals = jurnals;
            return View();
        }

        public string transformFromUnicode(string str)
        {
            if (String.IsNullOrEmpty(str))
            {
                return "";
            }

            var tmp = "";
            for (var i = 0; i < str.Length; i++)
            {
                switch ((int)str[i])
                {
                    case 4304:
                        tmp += (char)(192);
                        break;
                    case 4305:
                        tmp += (char)(193);
                        break;
                    case 4306:
                        tmp += (char)(194);
                        break;
                    case 4307:
                        tmp += (char)(195);
                        break;
                    case 4308:
                        tmp += (char)(196);
                        break;
                    case 4309:
                        tmp += (char)(197);
                        break;
                    case 4310:
                        tmp += (char)(198);
                        break;
                    case 4311:
                        tmp += (char)(200);
                        break;
                    case 4312:
                        tmp += (char)(201);
                        break;
                    case 4313:
                        tmp += (char)(202);
                        break;
                    case 4314:
                        tmp += (char)(203);
                        break;
                    case 4315:
                        tmp += (char)(204);
                        break;
                    case 4316:
                        tmp += (char)(205);
                        break;
                    case 4317:
                        tmp += (char)(207);
                        break;
                    case 4318:
                        tmp += (char)(208);
                        break;
                    case 4319:
                        tmp += (char)(209);
                        break;
                    case 4320:
                        tmp += (char)(210);
                        break;
                    case 4321:
                        tmp += (char)(211);
                        break;
                    case 4322:
                        tmp += (char)(212);
                        break;
                    case 4323:
                        tmp += (char)(214);
                        break;
                    case 4324:
                        tmp += (char)(215);
                        break;
                    case 4325:
                        tmp += (char)(216);
                        break;
                    case 4326:
                        tmp += (char)(217);
                        break;
                    case 4327:
                        tmp += (char)(218);
                        break;
                    case 4328:
                        tmp += (char)(219);
                        break;
                    case 4329:
                        tmp += (char)(220);
                        break;
                    case 4330:
                        tmp += (char)(221);
                        break;
                    case 4331:
                        tmp += (char)(222);
                        break;
                    case 4332:
                        tmp += (char)(223);
                        break;
                    case 4333:
                        tmp += (char)(224);
                        break;
                    case 4334:
                        tmp += (char)(225);
                        break;
                    case 4335:
                        tmp += (char)(227);
                        break;
                    case 4336:
                        tmp += (char)(228);
                        break;
                    default:
                        tmp += str[i];
                        break;

                }
            }

            return tmp;
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Index(List<IFormFile> files, string jurnal)
        {

            if (files.Count == 0)
            {
                ViewBag.error = "ატვირთეთ ფაილი ! ";

                var jurnals = getJurnals();
                ViewBag.jurnals = jurnals;

                return View("Index");
            }
            if (jurnal == "0")
            {
                ViewBag.error = "აირჩიეთ ჟურნალი ! ";

                var jurnals = getJurnals();
                ViewBag.jurnals = jurnals;

                return View("Index");
            }

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

                List<ExcelData> notExistInDb = new List<ExcelData>() { };

                List<string> TColumnStrings = Configuration.GetSection("TColumnStrings").Get<List<string>>() ?? new List<string> ();

                for (int i = 14; i <= rowLength; i++)
                {
                    if (workSheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                    if (workSheet.Cells[i, 7].Value?.ToString() == "COM" || workSheet.Cells[i, 7].Value?.ToString() == "FEE" 
                        || workSheet.Cells[i, 7].Value?.ToString() == "CCO"

                        || workSheet.Cells[i, 6].Value?.ToString().StartsWith("CCO") == true

                        || String.IsNullOrEmpty( workSheet.Cells[i, 1].Value?.ToString()))
                    {
                        continue;
                    }

                    if (TColumnStrings.Any(r => r.ToLower().Replace(" ", "") == workSheet.Cells[i, 20].Value?.ToString().ToLower().Replace(" ", "")))
                    {
                        continue;
                    }

                    var personalNumberInDesc = workSheet.Cells[i, 6].Value?.ToString().Length < 11? "": workSheet.Cells[i, 6].Value?.ToString().Substring(0, 11);

                    try
                    {
                        Int64.Parse(personalNumberInDesc);
                    }
                    catch(Exception e)
                    {
                        personalNumberInDesc = "";
                    }


                    
                    var IdentNumber = "";

                    if (String.IsNullOrEmpty(personalNumberInDesc))
                    {
                        IdentNumber = workSheet.Cells[i, 16].Value?.ToString();

                        if (workSheet.Cells[i, 16].Value?.ToString() == Configuration["VATNumber"].ToString())
                        {
                            IdentNumber = workSheet.Cells[i, 11].Value.ToString();
                        }
                        if (TColumnStrings.Any(r => r.ToLower().Replace(" ", "") == workSheet.Cells[i, 20].Value?.ToString().ToLower().Replace(" ", "")))
                        {
                            continue;
                        }
                    }
                    else
                    {
                        IdentNumber = personalNumberInDesc;
                    }

                    if (String.IsNullOrEmpty(IdentNumber))
                    {
                        ViewBag.error = "შეავსეთ საიდენტიფიკაციო ნომერი ! ხაზი " + i.ToString();

                        var jurnals = getJurnals();
                        ViewBag.jurnals = jurnals;

                        return View("Index");

                    }
                    notExistInDb.Add(new ExcelData { 
                        ID = Guid.NewGuid(),
                        StatementNumber = workSheet.Cells[i, 2].Value?.ToString(),
                        Debit = workSheet.Cells[i, 4].Value?.ToString(),
                        Credit = workSheet.Cells[i, 5].Value?.ToString(),
                        OperationContent = workSheet.Cells[i, 6].Value?.ToString().Length < 40? workSheet.Cells[i, 6].Value?.ToString()
                                                : workSheet.Cells[i, 6].Value?.ToString()?.Substring(0, 40),
                        OperationType = workSheet.Cells[i, 7].Value?.ToString(),
                        ReceiverName = workSheet.Cells[i, 15].Value?.ToString(),
                        IdentityNumber = IdentNumber,
                        Destination = workSheet.Cells[i, 20].Value?.ToString().Length < 20 ? workSheet.Cells[i, 20].Value?.ToString()
                                                : workSheet.Cells[i, 20].Value?.ToString()?.Substring(0, 20),
                    });
                }



                var nonExist = getNonExistIdentificators(notExistInDb.Select(r => r.IdentityNumber).ToList());

                IList<ExcelData> nonExistsData = new List<ExcelData>();


                if (nonExist.Count != 0)
                {
                    var text = "";
                    foreach (var item in nonExist)
                    {
                        var d = notExistInDb.Where(r => r.IdentityNumber == item).FirstOrDefault();
                        nonExistsData.Add(new ExcelData() { 
                            StatementNumber = d.StatementNumber,
                            Debit = d.Debit,
                            Credit = d.Credit,
                            OperationContent = d.OperationContent,
                            OperationType = d.OperationType,
                            ReceiverName = d.ReceiverName,
                            IdentityNumber = d.IdentityNumber,
                            Destination = d.Destination
                        });
                    }


                    ViewData["nonExists"] = nonExistsData;

                    var jurnals = getJurnals();
                    ViewBag.jurnals = jurnals;

                    return View("Index");
                }
                


                

                xmlDoc = importToXML(workSheet, jurnal);


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

        public List<Jurnal> getJurnals()
        {

            //string connectionString =
            // "Data Source=(local);Initial Catalog=Northwind;"
            // + "Integrated Security=true";

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            var dagbknr = "202";
            List<Jurnal> jurnals = new List<Jurnal>();

            string queryString =
                @"
                select distinct dagbk.ID, dagbknr, dbo.getunicode(dagbk.oms25_0) as dagbkDesc, dagbk.reknr,dbo.getunicode(grtbk.oms25_0) as grtbkDesc
                    from dagbk inner join grtbk on dagbk.reknr = grtbk.reknr
                    where type_dgbk = 'B'";



            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@dagbknr", dagbknr);

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        jurnals.Add( new Jurnal
                        {
                            ID = reader[0].ToString(),
                            Dagbknr = reader[1].ToString(),
                            DagbkDesc = reader[2].ToString(),
                            Reknr = reader[3].ToString(),
                            GrtbkDesc = reader[4].ToString(),
                        });
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }

            
            return jurnals;
        }

        public List<string> getNonExistIdentificators(List<string> names)
        {
            if(names.Count() == 0)
            {
                return names;
            }
            //string connectionString =
            // "Data Source=(local);Initial Catalog=Northwind;"
            // + "Integrated Security=true";

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            var dagbknr = "202";
            List<string> nonExists = new List<string>();
            var list = String.Join(", ", names.ToArray());



            string queryString =
                "SELECT vatnumber FROM cicmpy where vatnumber in ";
            //   + "where dagbknr = @dagbknr ";
            var cc = "(";

          

            string qt = "select iden from (select '"+ names[0]+"' as 'iden'";

            for (int i = 1; i< names.Count; i++)
            {
                qt += " union select '" + names[i]+"'";
            }

            qt += ") a   where not exists(select vatnumber from cicmpy where vatnumber = iden)";

          

            queryString = qt;





            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@list", list);



                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        nonExists.Add(reader[0].ToString());
                    }

                    reader.Close();

                   
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                }
                //Console.ReadLine();
            }

            return nonExists;
        }

        public string getEntryNumber(string dagbknr)
        {

            //string connectionString =
            // "Data Source=(local);Initial Catalog=Northwind;"
            // + "Integrated Security=true";

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            //var dagbknr = "202";
            string entryNumber = "" ;
            string queryString =
                "select max (bkstnr) from gbkmut "
            +"where dagbknr = @dagbknr ";



            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@dagbknr", dagbknr);

                SqlCommand command1 = new SqlCommand(queryString, connection);

              
                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        entryNumber = reader[0].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }
            if (String.IsNullOrEmpty(entryNumber))
            {
                entryNumber = "0";
            }

            int newEntry = Int32.Parse(entryNumber);
            newEntry++;

            string newEntryString = newEntry.ToString();
            if (newEntryString.Length != 8)
            {
                newEntryString = newEntryString.PadLeft(8,'0');
            }

            return newEntryString;
        }

        public JurnalInfo getJurnalInfo(string jurn)
        {

            //string connectionString =
            // "Data Source=(local);Initial Catalog=Northwind;"
            // + "Integrated Security=true";

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            var dagbknr = "202";
            string queryString =
                @"select d.dagbknr, d.reknr, g.bal_vw, g. debcrd, g.omzrek
 from dagbk d inner
 join grtbk g on d.reknr = g.reknr
 where g.omzrek = 'B'and dagbknr = @jurn ";


            var journal = new JurnalInfo { };
            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@jurn", jurn);

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        journal.dagbknr = reader[0].ToString();
                        journal.reknr = reader[1].ToString();
                        journal.bal_vw = reader[2].ToString();
                        journal.debcrd = reader[3].ToString();
                        journal.omzrek = reader[4].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }
           

            return journal;
        }

        public string getInvoiceNumber()
        {
            
            //+1
            
            //745600 W K D


            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");


            var dagbknr = "202";
            string invoiceNumber = "";
            //select faktuurnr from gbkmut where dagbknr='202' დაიწყება 802-ით '80200001'
            string queryString =
                "select faktuurnr from gbkmut where faktuurnr like '114%' order by faktuurnr desc";
            

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
            return "114200001";
        }


        public Cicmpy getCreditorFromDB(string identityNumber)
        {

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");

            string queryString = "select top 1 crdnr,debnr,cmp_name,ClassificationId from cicmpy where vatnumber = @identityNumber ";

            var Cicmpy = new Cicmpy() { };

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
                        Cicmpy.Crdnr= reader[0].ToString();
                        Cicmpy.Debnr = reader[1].ToString();
                        Cicmpy.CmpName = reader[2].ToString();
                        Cicmpy.ClassificationId = reader[3].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }

            
            return Cicmpy;
        }

        public bool checkIfExist(string target)
        {

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");

            string queryString = "select top 1 ordernr from orkrg where ordernr=@target";

            var res = "";

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@target", target);

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        res = reader[0].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }

            if (String.IsNullOrEmpty(res))
            {
                return false;
            }

            return true;
        }

        public string getDivision()
        {

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");

            string queryString = "select bedrnr from bedryf";

            var div = "";
            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);


                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        div = reader[0].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }


            return div;
        }

        public FileStreamResult importToXML(ExcelWorksheet workSheet, string jurnal)
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

                var division = getDivision();

                var GLEntryNode = getGLEntryNode(doc, workSheet, jurnal, division);

                var rowLength = workSheet.Dimension.End.Row;

               

                var invoiceNumber = getInvoiceNumber();

                XmlNode BankStatement = doc.CreateElement("BankStatement");
                ((XmlElement)BankStatement).SetAttribute("number", "24020023");

                XmlNode Date = doc.CreateElement("Date");
                Date.AppendChild(doc.CreateTextNode("2022-02-03"));
                BankStatement.AppendChild(Date);

                XmlNode GLOffset = doc.CreateElement("GLOffset");
                ((XmlElement)GLOffset).SetAttribute("code", "   100001");
                BankStatement.AppendChild(GLOffset);

                List<string> TColumnStrings = Configuration.GetSection("TColumnStrings").Get<List<string>>() ?? new List<string>(); ;


                var comIndex = 0;
                var invNumber = Int32.Parse(invoiceNumber);
                for (int i = 14; i <= rowLength; i++)
                {
                    if (workSheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                    if(workSheet.Cells[i, 7].Value.ToString() == "COM" || workSheet.Cells[i, 7].Value.ToString() == "FEE")
                    {
                        continue;
                    }

                    if (TColumnStrings.Any(r => r.ToLower().Replace(" ", "") == workSheet.Cells[i, 20].Value?.ToString().ToLower().Replace(" ", "")))
                    {
                        continue;
                    }


                    var commonId = Guid.NewGuid();
                    invNumber++;
                    var FinEntryLine = getFinEntryLine(division,i, doc, workSheet, invNumber.ToString(), commonId);
                    GLEntryNode.AppendChild(FinEntryLine);

                    // var BankStatementLine = getBankStatement(i, doc, workSheet, commonId);
                    // BankStatement.AppendChild(BankStatementLine);
                }

                double sumAmount = 0;
                var existCOM = false;
                var COMI = 0;
                for (int i = 14; i <= rowLength; i++)
                {
                    if (workSheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                    if (workSheet.Cells[i, 7].Value.ToString() == "COM" || workSheet.Cells[i, 7].Value.ToString() == "FEE")
                    {
                        existCOM = true;
                        COMI = i;
                        sumAmount += Double.Parse( workSheet.Cells[i, 4].Value.ToString());
                    }



                    // var BankStatementLine = getBankStatement(i, doc, workSheet, commonId);
                    // BankStatement.AppendChild(BankStatementLine);
                }

                if(existCOM == true)
                {
                    invNumber++;
                    var FinEntryLine = getFinEntryLine(division,COMI, doc, workSheet, invNumber.ToString(), Guid.NewGuid(), sumAmount);
                    GLEntryNode.AppendChild(FinEntryLine);
                }

                //TColumnStrings
                double sumAmountForT = 0;
                var existT = false;
                var TIndex = 0;
                for (int i = 14; i <= rowLength; i++)
                {
                    if (workSheet.Cells[i, 1].Value == null)
                    {
                        break;
                    }
                    if (workSheet.Cells[i, 7].Value?.ToString() == "COM" || workSheet.Cells[i, 7].Value?.ToString() == "FEE"
                        || workSheet.Cells[i, 7].Value?.ToString() == "CCO"

                        || workSheet.Cells[i, 6].Value?.ToString().StartsWith("CCO") == true

                        || String.IsNullOrEmpty(workSheet.Cells[i, 1].Value?.ToString()))
                    {
                        continue;
                    }
                    if (TColumnStrings.Any(r => r.ToLower().Replace(" ", "") == workSheet.Cells[i, 20].Value?.ToString().ToLower().Replace(" ", "")))
                    {
                        existT = true;
                        TIndex = i;
                        sumAmountForT += Double.Parse(workSheet.Cells[i, 4].Value.ToString());
                    }



                    // var BankStatementLine = getBankStatement(i, doc, workSheet, commonId);
                    // BankStatement.AppendChild(BankStatementLine);
                }

                if (existT == true)
                {
                    invNumber++;
                    var FinEntryLine = getFinEntryLine(division, TIndex, doc, workSheet, invNumber.ToString(), Guid.NewGuid(), sumAmountForT, true);
                    GLEntryNode.AppendChild(FinEntryLine);
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


        public string getGLAccountInEntryLine(string code, Cicmpy cicmpy, string description)
        {

   //         reknr bal_vw  omzrek debcrd
               //100001   B C   D
               //140001   B D   D
               //311010   B C   C
               //745600   W K   D


            if (code == "COM" || code == "FEE")
            {
                return "745600";
            }
            if (code == "CCO" || description.StartsWith("CCO"))//description
            {
                //costcent = 00
                //costunit = 00
                return "100001";
            }
            if(cicmpy.isDebnr == true)
            {
                return "140001";
            }

            if(cicmpy.ClassificationId == "300")
            {
                return "143010";
            }

            //თუ დებიტორი მაშინ  return '140001'
            return "310001";

            //cicmpy debnr is not null => '140001'
        }

        public string getUnicode(string text)
        {
            var res = transformFromUnicode(text);

            return res;
        }

        public string getDesciption(string text, string code, string recipientName)
        {
            if(code == "CCO" )
            {
                text = "კონვერტაცია";
            }

            if (code == "COM" || code == "FEE")
            {
                text = "ბანკის საკომისიო";
            }

            List<string> strings = Configuration.GetSection("DescriptionStopStrings").Get<List<string>>();

            int indexOfSym = -1;
            indexOfSym = text.IndexOf(recipientName);
            if (indexOfSym >= 0)
            {
                text = text.Substring(0, indexOfSym);
            }
            else
            {
                indexOfSym = -1;
                foreach (var item in strings)
                {
                    indexOfSym = text.IndexOf(item);
                    if (indexOfSym >= 0)
                    {
                        text = text.Substring(0, indexOfSym);
                    }
                }
            }




            var cc = getUnicode(text);
            return cc.Length < 60?  cc :  cc.Substring(0, 60);
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
            Description.AppendChild(doc.CreateTextNode(getDesciption(worksheet.Cells[i, 6].Value.ToString(),
                                                                        worksheet.Cells[i, 7].Value.ToString(),
                                                                        worksheet.Cells[i, 15].Value.ToString())));
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

            var cr = getCreditorCode(worksheet.Cells[i, 7].Value.ToString(), worksheet.Cells[i, 16].Value.ToString(), worksheet.Cells[i, 6].Value.ToString(), false);


            var gLAccountCode = getGLAccountInEntryLine(worksheet.Cells[i, 7].Value.ToString(), cr, worksheet.Cells[i, 6].Value.ToString());
            
            XmlNode GLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)GLAccount).SetAttribute("code", gLAccountCode);


            //100001   B C   D
            //140001   B D   D
            //311010   B C   C
            //745600   W K   D

            if (gLAccountCode == "745600")
            {
                ((XmlElement)GLAccount).SetAttribute("type", "W");
                ((XmlElement)GLAccount).SetAttribute("subtype", "K");
                ((XmlElement)GLAccount).SetAttribute("side", "D");
            }
            if (gLAccountCode == "310001")
            {
                ((XmlElement)GLAccount).SetAttribute("type", "B");
                ((XmlElement)GLAccount).SetAttribute("subtype", "C");
                ((XmlElement)GLAccount).SetAttribute("side", "C");

                //costcentr = 100
                //costunit = 00
            }
            if (gLAccountCode == "140001")//140001
            {
                ((XmlElement)GLAccount).SetAttribute("type", "B");
                ((XmlElement)GLAccount).SetAttribute("subtype", "D");
                ((XmlElement)GLAccount).SetAttribute("side", "D");
            }
            if (gLAccountCode == "100001")
            {
                ((XmlElement)GLAccount).SetAttribute("type", "B");
                ((XmlElement)GLAccount).SetAttribute("subtype", "C");
                ((XmlElement)GLAccount).SetAttribute("side", "D");
            }
            if (gLAccountCode == "143010")
            {
                ((XmlElement)GLAccount).SetAttribute("type", "B");
                ((XmlElement)GLAccount).SetAttribute("subtype", "C");
                ((XmlElement)GLAccount).SetAttribute("side", "C");
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

            ((XmlElement)Creditor).SetAttribute("code", cr.Crdnr);
            ((XmlElement)Creditor).SetAttribute("number", cr.Crdnr);
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

        public XmlNode getGLEntryNode(XmlDocument doc,  ExcelWorksheet worksheet, string jurnal, string division)
        {


            //GLEntry
            XmlNode GLEntryNode = doc.CreateElement("GLEntry");
            //select max(bkstnr) from gbkmut   ->entry
            //where dagbknr = '202' +
            //select cmp_code, cmp_name, VatNumber from cicmpy   creditor Name-shi

            // tu crdnr -> Creditor, debnr -> Debitor,
            // saidentifikacios shemowmebisas ar gaivaliswinos COM CCO +
            var entryNumber = getEntryNumber(jurnal);
            ((XmlElement)GLEntryNode).SetAttribute("entry", entryNumber);
            ((XmlElement)GLEntryNode).SetAttribute("status", "E");

            XmlNode Division = doc.CreateElement("Division");
            ((XmlElement)Division).SetAttribute("code", division);
            GLEntryNode.AppendChild(Division);

            XmlNode DocumentDate = doc.CreateElement("DocumentDate");
            var d = worksheet.Cells[14, 1].Value.ToString(); //პირველივე თარიღი
            var dateFormated = DateTime.Parse(d).ToString("yyyy-MM-dd");
            DocumentDate.AppendChild(doc.CreateTextNode(dateFormated));
            GLEntryNode.AppendChild(DocumentDate);

            XmlNode Journal = doc.CreateElement("Journal");
            ((XmlElement)Journal).SetAttribute("code", jurnal);
            ((XmlElement)Journal).SetAttribute("type", "B");


            XmlNode Description = doc.CreateElement("Description");
            Description.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            Journal.AppendChild(Description);

            //
            var jur = getJurnalInfo(jurnal);
            XmlNode GLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)GLAccount).SetAttribute("code", jur.reknr);
            ((XmlElement)GLAccount).SetAttribute("type", jur.bal_vw);
            ((XmlElement)GLAccount).SetAttribute("subtype", jur.omzrek);
            ((XmlElement)GLAccount).SetAttribute("side", jur.debcrd);

            XmlNode GLDescription = doc.CreateElement("Description");
            GLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            GLAccount.AppendChild(GLDescription);


            XmlNode GLPaymentInTransit = doc.CreateElement("GLPaymentInTransit");
            ((XmlElement)GLPaymentInTransit).SetAttribute("code", "999001");
            ((XmlElement)GLPaymentInTransit).SetAttribute("type", "B");
            ((XmlElement)GLPaymentInTransit).SetAttribute("subtype", "N");
            ((XmlElement)GLPaymentInTransit).SetAttribute("side", "C");

            XmlNode GLPaymentInTransitDescription = doc.CreateElement("Description");
            GLPaymentInTransitDescription.AppendChild(doc.CreateTextNode(transformFromUnicode("გაუნაწილებელი თანხები ბანკში")));
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
            var glAccounCodes = getGlAccountCodes("719990".PadLeft(9, ' '));
            ((XmlElement)CostcenterGLAccount).SetAttribute("code", "719990".PadLeft(9, ' '));
            ((XmlElement)CostcenterGLAccount).SetAttribute("type", glAccounCodes?.Type);
            ((XmlElement)CostcenterGLAccount).SetAttribute("subtype", glAccounCodes?.SubType);
            ((XmlElement)CostcenterGLAccount).SetAttribute("side", glAccounCodes?.Side);

            XmlNode CostcenterGLAccountGLDescription = doc.CreateElement("Description");
            CostcenterGLAccountGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            CostcenterGLAccount.AppendChild(CostcenterGLAccountGLDescription);

            Costcenter.AppendChild(CostcenterGLAccount);

            XmlNode GLOffset = doc.CreateElement("GLOffset");
            ((XmlElement)GLOffset).SetAttribute("code", "719990".PadLeft(9, ' '));
            ((XmlElement)GLOffset).SetAttribute("type", glAccounCodes?.Type);
            ((XmlElement)GLOffset).SetAttribute("subtype", glAccounCodes?.SubType);
            ((XmlElement)GLOffset).SetAttribute("side", glAccounCodes?.Side);

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

        public string getGLAccountCodeFromDB(string crdnr)
        {

            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");

            string queryString = "select top 1 CentralizationAccount from cicmpy  where ltrim(rtrim(crdnr))= ltrim(rtrim(@crdnr))";

            string glAccountCode = "";

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@crdnr", crdnr);

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        glAccountCode = reader[0].ToString();
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }


            return glAccountCode;
        }

        public Cicmpy  getCreditorCode(string code, string identityNymber, string description, bool isSum)
        {
            // CCO -> კრედიტორი 3
            // COM -> კრედიტორი 4
            // სვა შემთხვევაში select vatnumber,crdnr,debnr from cicmpy where VatNumber = '102189454'(მიმღების საიდენთიფიკაციო კოდი)
            // რომელიც null არაა იმით შეივსება

            // if divisio == 150 -> glaccount 745600  if division == 300 -> glaccount 747000 if divission == 350 glaccount 747000 else 747000

            if (isSum)
            {
                string TColumnCreditor = Configuration["TColumnCreditor"].ToString();
                return new Cicmpy()
                {
                    DefaultCode = TColumnCreditor,
                    FromDB = false,
                    isDebnr = false
                };
            }

            if (code == "COM" || code == "FEE")
            {
                return new Cicmpy()
                {
                    DefaultCode = "200002",
                    FromDB = false,
                    isDebnr = false
                };
            }
            if (code == "CCO" || description.StartsWith("CCO")) //description-ში  შემოწმება
            {
                return new Cicmpy()
                {
                    DefaultCode = "200000",
                    FromDB = false,
                    isDebnr = false
                };
            }
            //if 
            var result = getCreditorFromDB(identityNymber);
            //if clasificationid == 300 -> glacount code =143010 
            //

            //select bedrnr from bedryf  -> division code

            //if select CentralizationAccount,* from cicmpy  where ltrim(rtrim(crdnr))= '1041' => glaccount-shi
            return new Cicmpy()
            {
                FromDB = true,
                CmpName = result.CmpName,
                Crdnr = result.Crdnr,
                Debnr = result.Debnr,
                isDebnr = String.IsNullOrEmpty(result.Debnr) ? false : true,
                ClassificationId = result.ClassificationId
            }; 

        }

        public GlAccountCodes getGlAccountCodes(string glAccount)
        {
            string connectionString = this.Configuration.GetConnectionString("DefaultConnection");



            GlAccountCodes glAccountCodes= null ;
            string queryString =
                @"select bal_vw,  omzrek, debcrd from grtbk where reknr = @reknr ";



            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                SqlCommand command = new SqlCommand(queryString, connection);

                command.Parameters.AddWithValue("@reknr", glAccount.PadLeft(9,' '));

                SqlCommand command1 = new SqlCommand(queryString, connection);


                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        glAccountCodes = new GlAccountCodes
                        {
                            Type = reader[0].ToString(),
                            SubType = reader[1].ToString(),
                            Side = reader[2].ToString(),
                        };
                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                //Console.ReadLine();
            }
           

            return glAccountCodes;
        }
        public XmlNode getFinEntryLine(string division,int i, XmlDocument doc, ExcelWorksheet worksheet,string  invoiceNumber,
            Guid commonId, double? sumAmount = 0, bool isSum = false)
        {
            
            //COM დაჯამდება, რეფერენსები იქნება საერთო
            XmlNode FinEntryLine = doc.CreateElement("FinEntryLine");
            ((XmlElement)FinEntryLine).SetAttribute("number", (i-13).ToString());
            ((XmlElement)FinEntryLine).SetAttribute("type", "N");

            var finEntryLineSubType = !String.IsNullOrEmpty(worksheet.Cells[i, 4].Value?.ToString()) ? "Y" : " Z";
            ((XmlElement)FinEntryLine).SetAttribute("subtype", finEntryLineSubType);

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

            var personalNumberInDesc = worksheet.Cells[i, 6].Value?.ToString().Length < 11 ? "" : worksheet.Cells[i, 6].Value?.ToString().Substring(0, 11);

            try
            {
                Int64.Parse(personalNumberInDesc);
            }
            catch (Exception e)
            {
                personalNumberInDesc = "";
            }

            var IdentNumber = "";



            if (String.IsNullOrEmpty(personalNumberInDesc))
            {
                IdentNumber = worksheet.Cells[i, 16].Value?.ToString();

                if (worksheet.Cells[i, 16].Value?.ToString() == Configuration["VATNumber"].ToString())
                {

                    IdentNumber = worksheet.Cells[i, 11].Value.ToString();
                }
            }
            else
            {
                IdentNumber = personalNumberInDesc;
            }
            //var identNumber = worksheet.Cells[i, 16].Value == null? "" : worksheet.Cells[i, 16].Value.ToString();
            //if (worksheet.Cells[i, 16].Value?.ToString() == Configuration["VATNumber"].ToString())
            //{
            //    identNumber = worksheet.Cells[i, 11].Value.ToString();
            //}

            var creditorRes = getCreditorCode(worksheet.Cells[i, 7].Value.ToString(), IdentNumber, worksheet.Cells[i, 6].Value.ToString(), isSum);

            //if divisio == 150 -> glaccount 745600  if division == 300 -> glaccount 747000 if divission == 350 glaccount 747000 else 747000

            var glAccountFromDb = getGLAccountCodeFromDB(creditorRes.Crdnr?? "");

            string gLAccountCode = "";
            if (!String.IsNullOrEmpty(glAccountFromDb))
            {
                gLAccountCode = glAccountFromDb;
            }
            else 
            { 
                gLAccountCode = getGLAccountInEntryLine(worksheet.Cells[i, 7].Value.ToString(), creditorRes, worksheet.Cells[i, 6].Value.ToString());

            }

            if (worksheet.Cells[i, 7].Value.ToString() == "COM" || (worksheet.Cells[i, 7].Value.ToString() == "FEE"))
            {
                if (division == "600")
                {
                    gLAccountCode = "745600";
                }
                else
                {
                    gLAccountCode = "747000";
                }
            }
           

            XmlNode FinEntryLineGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)FinEntryLineGLAccount).SetAttribute("code", gLAccountCode);
            var glAccountCodes = getGlAccountCodes(gLAccountCode);
            ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", glAccountCodes?.Type);
            ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", glAccountCodes?.SubType);
            ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", glAccountCodes?.Side);
            //if (gLAccountCode == "747000")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "W");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "A");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            //}
            //if (gLAccountCode == "745600")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "W");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "K");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            //}
            //if (gLAccountCode == "311010")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "C");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "C");
            //}
            //if (gLAccountCode == "140001")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "D");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            //}
            //if (gLAccountCode == "100001")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "C");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "D");
            //}
            //if (gLAccountCode == "143010")
            //{
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("type", "B");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("subtype", "C");
            //    ((XmlElement)FinEntryLineGLAccount).SetAttribute("side", "C");
            //}

            XmlNode FinEntryLineGLDescription = doc.CreateElement("Description");
            FinEntryLineGLDescription.AppendChild(doc.CreateTextNode("GEL 3406000029"));
            FinEntryLineGLAccount.AppendChild(FinEntryLineGLDescription);
            FinEntryLine.AppendChild(FinEntryLineGLAccount);


            XmlNode FinEntryLineDescription = doc.CreateElement("Description");
            FinEntryLineDescription.AppendChild(doc.CreateTextNode(getDesciption(worksheet.Cells[i, 6].Value.ToString(),
                                                                                 worksheet.Cells[i, 7].Value.ToString(),
                                                                                 worksheet.Cells[i, 15].Value.ToString())));
            FinEntryLine.AppendChild(FinEntryLineDescription);


            //-----------------------------
            XmlNode FinEntryLineCostcenter = doc.CreateElement("Costcenter");
            if (worksheet.Cells[i, 7].Value.ToString() == "COM" || (worksheet.Cells[i, 7].Value.ToString() == "FEE") || gLAccountCode == "745600")
            {
                ((XmlElement)FinEntryLineCostcenter).SetAttribute("code", "80");
            }
            else
            {
                ((XmlElement)FinEntryLineCostcenter).SetAttribute("code", "001CC001");

            }

            //chavamtot CostUnit
            //745600 -> costUnit = 8088.861


            XmlNode FinEntryLineCostcenterDescription = doc.CreateElement("Description");
            FinEntryLineCostcenterDescription.AppendChild(doc.CreateTextNode("Default cost center"));
            FinEntryLineCostcenter.AppendChild(FinEntryLineCostcenterDescription);

            XmlNode FinEntryLineCostcenterGLAccount = doc.CreateElement("GLAccount");
            ((XmlElement)FinEntryLineCostcenterGLAccount).SetAttribute("code", "     9999");
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

            var tagName = String.IsNullOrEmpty(creditorRes.Crdnr) ? "Debtor" : "Creditor";
            var tagValue = tagName == "Creditor" ? creditorRes.Crdnr : creditorRes.Debnr;

            if(creditorRes.FromDB == false)
            {
                tagName = "Creditor";
                tagValue = creditorRes.DefaultCode;
            }

            XmlNode Creditor = doc.CreateElement(tagName);
            ((XmlElement)Creditor).SetAttribute("code", tagValue);
            ((XmlElement)Creditor).SetAttribute("number", tagValue);
            ((XmlElement)Creditor).SetAttribute("type", "S");

            XmlNode CreditorName = doc.CreateElement("Name");
            CreditorName.AppendChild(doc.CreateTextNode(creditorRes.CmpName));
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

            var codeForProject = "";
            var rowDescription = worksheet.Cells[i, 21].Value.ToString();
            try
            {
                int startInd = rowDescription.IndexOf('[');
                int endInd = rowDescription.IndexOf(']') - startInd;
                codeForProject = rowDescription.Substring(startInd + 1, endInd - 1);
                bool exist = checkIfExist(codeForProject);
                if (!exist)
                {
                    codeForProject = "";
                }
            }
            catch(Exception e)
            {
                codeForProject = "";
            }

            if (!String.IsNullOrEmpty(codeForProject))
            {
                XmlNode Project = doc.CreateElement("Project");
                ((XmlElement)Project).SetAttribute("code", codeForProject);
                ((XmlElement)Project).SetAttribute("type", "I");
                ((XmlElement)Project).SetAttribute("status", "P");

                XmlNode projectDescr = doc.CreateElement("Description");
                projectDescr.AppendChild(doc.CreateTextNode(codeForProject));
                Project.AppendChild(projectDescr);

                XmlNode projectSecurityLevel = doc.CreateElement("SecurityLevel");
                projectSecurityLevel.AppendChild(doc.CreateTextNode("10"));
                Project.AppendChild(projectSecurityLevel);

                XmlNode projectDateStart = doc.CreateElement("DateStart");
                projectDateStart.AppendChild(doc.CreateTextNode(DateTime.Now.ToString("yyyy-MM-dd")));
                Project.AppendChild(projectDateStart);

                XmlNode projectDateEnd = doc.CreateElement("DateEnd");
                projectDateEnd.AppendChild(doc.CreateTextNode(DateTime.Now.ToString("yyyy-MM-dd")));
                Project.AppendChild(projectDateEnd);

                XmlNode projectAssortment = doc.CreateElement("Assortment");
                ((XmlElement)projectAssortment).SetAttribute("code", "-1");
                Project.AppendChild(projectAssortment);

                FinEntryLine.AppendChild(Project);

                XmlNode Payment = doc.CreateElement("Payment");

                XmlNode PaymentTransN = doc.CreateElement("TransactionNumberSubAdministration");
                PaymentTransN.AppendChild(doc.CreateTextNode(codeForProject));
                Payment.AppendChild(PaymentTransN);

                FinEntryLine.AppendChild(Payment);
            }




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
            if(sumAmount > 0)
            {
                Debit.AppendChild(doc.CreateTextNode(sumAmount.ToString()));

            }
            else
            {
                Debit.AppendChild(doc.CreateTextNode(!String.IsNullOrEmpty(worksheet.Cells[i, 4].Value?.ToString()) ? worksheet.Cells[i, 4].Value?.ToString() : "0"));

            }
            FinEntryLineAmount.AppendChild(Debit);

            //კრედიტის ველიდან
            XmlNode Credit = doc.CreateElement("Credit");
            Credit.AppendChild(doc.CreateTextNode(!String.IsNullOrEmpty(worksheet.Cells[i, 5].Value?.ToString())? worksheet.Cells[i, 5].Value?.ToString() : "0"));
            FinEntryLineAmount.AppendChild(Credit);

            //XmlNode VAT = doc.CreateElement("VAT");
            //((XmlElement)VAT).SetAttribute("code", "0");
            //((XmlElement)VAT).SetAttribute("type", "B");
            //((XmlElement)VAT).SetAttribute("vattype", "N");
            //((XmlElement)VAT).SetAttribute("taxtype", "V");

            //XmlNode VATDescription = doc.CreateElement("Description");
            //VATDescription.AppendChild(doc.CreateTextNode("VAT 0%"));
            //VAT.AppendChild(VATDescription);

            //XmlNode MultiDescriptions = doc.CreateElement("MultiDescriptions");

            //XmlNode MultiDescription1 = doc.CreateElement("MultiDescription");
            //((XmlElement)MultiDescription1).SetAttribute("number", "1");
            //MultiDescription1.AppendChild(doc.CreateTextNode("VAT 0%"));
            //MultiDescriptions.AppendChild(MultiDescription1);

            //XmlNode MultiDescription2 = doc.CreateElement("MultiDescription");
            //((XmlElement)MultiDescription2).SetAttribute("number", "2");
            //MultiDescription2.AppendChild(doc.CreateTextNode("VAT 0%"));
            //MultiDescriptions.AppendChild(MultiDescription2);

            //XmlNode MultiDescription3 = doc.CreateElement("MultiDescription");
            //((XmlElement)MultiDescription3).SetAttribute("number", "3");
            //MultiDescription3.AppendChild(doc.CreateTextNode("VAT 0%"));
            //MultiDescriptions.AppendChild(MultiDescription3);

            //XmlNode MultiDescription4 = doc.CreateElement("MultiDescription");
            //((XmlElement)MultiDescription4).SetAttribute("number", "4");
            //MultiDescription4.AppendChild(doc.CreateTextNode("VAT 0%"));
            //MultiDescriptions.AppendChild(MultiDescription4);

            //VAT.AppendChild(MultiDescriptions);


            //XmlNode Percentage = doc.CreateElement("Percentage");
            //Percentage.AppendChild(doc.CreateTextNode("0"));
            //VAT.AppendChild(Percentage);

            //XmlNode Charged = doc.CreateElement("Charged");
            //Charged.AppendChild(doc.CreateTextNode("0"));
            //VAT.AppendChild(Charged);

            //XmlNode VATExemption = doc.CreateElement("VATExemption");
            //VATExemption.AppendChild(doc.CreateTextNode("0"));
            //VAT.AppendChild(VATExemption);

            //XmlNode ExtraDutyPercentage = doc.CreateElement("ExtraDutyPercentage");
            //ExtraDutyPercentage.AppendChild(doc.CreateTextNode("0"));
            //VAT.AppendChild(ExtraDutyPercentage);



            //XmlNode GLToPay = doc.CreateElement("GLToPay");
            //((XmlElement)GLToPay).SetAttribute("code", "   333010");
            //((XmlElement)GLToPay).SetAttribute("side", "C");
            //((XmlElement)GLToPay).SetAttribute("type", "B");
            //((XmlElement)GLToPay).SetAttribute("subtype", "C");

            //XmlNode GLToPayDescription = doc.CreateElement("Description");
            //GLToPayDescription.AppendChild(doc.CreateTextNode("|^~^&#1031;^&#1118;~&#166;&#1106;&#1107; ~&#1116;|0"));
            //GLToPay.AppendChild(GLToPayDescription);

            //VAT.AppendChild(GLToPay);

            //XmlNode GLToClaim = doc.CreateElement("GLToClaim");
            //((XmlElement)GLToClaim).SetAttribute("code", "   333010");
            //((XmlElement)GLToClaim).SetAttribute("side", "C");
            //((XmlElement)GLToClaim).SetAttribute("type", "B");
            //((XmlElement)GLToClaim).SetAttribute("subtype", "C");

            //XmlNode GLToClaimDescription = doc.CreateElement("Description");
            //GLToClaimDescription.AppendChild(doc.CreateTextNode("|^~^&#1031;^&#1118;~&#166;&#1106;&#1107; ~&#1116;|0"));
            //GLToClaim.AppendChild(GLToClaimDescription);

            //VAT.AppendChild(GLToClaim);

            //XmlNode VATCreditor = doc.CreateElement("Creditor");
            //((XmlElement)VATCreditor).SetAttribute("code", "        1");
            //((XmlElement)VATCreditor).SetAttribute("number", "        1");
            //((XmlElement)VATCreditor).SetAttribute("type", "S");

            //XmlNode VATCreditorName = doc.CreateElement("Name");
            //VATCreditorName.AppendChild(doc.CreateTextNode("VAT Creditor"));
            //VATCreditor.AppendChild(VATCreditorName);

            //VAT.AppendChild(VATCreditor);

            //XmlNode PaymentPeriod = doc.CreateElement("PaymentPeriod");
            //PaymentPeriod.AppendChild(doc.CreateTextNode("M"));

            //VAT.AppendChild(PaymentPeriod);

            //FinEntryLineAmount.AppendChild(VAT);
            FinEntryLine.AppendChild(FinEntryLineAmount);


            //XmlNode VATTransaction = doc.CreateElement("VATTransaction");
            //((XmlElement)VATTransaction).SetAttribute("code", "0");

            //XmlNode VATAmount = doc.CreateElement("VATAmount");
            //VATAmount.AppendChild(doc.CreateTextNode("0"));
            //VATTransaction.AppendChild(VATAmount);

            //XmlNode VATBaseAmount = doc.CreateElement("VATBaseAmount");
            //VATBaseAmount.AppendChild(doc.CreateTextNode("0"));
            //VATTransaction.AppendChild(VATBaseAmount);

            //XmlNode VATBaseAmountFC = doc.CreateElement("VATBaseAmountFC");
            //VATBaseAmountFC.AppendChild(doc.CreateTextNode("0"));
            //VATTransaction.AppendChild(VATBaseAmountFC);

            //FinEntryLine.AppendChild(VATTransaction);



            //XmlNode Payment = doc.CreateElement("Payment");

            //XmlNode PaymentMethod = doc.CreateElement("PaymentMethod");
            //((XmlElement)PaymentMethod).SetAttribute("code", "B");
            //Payment.AppendChild(PaymentMethod);

            //XmlNode PaymentCondition = doc.CreateElement("PaymentCondition");
            //((XmlElement)PaymentCondition).SetAttribute("code", "00");

            //XmlNode PaymentConditionDescription = doc.CreateElement("Description");
            //((XmlElement)PaymentConditionDescription).SetAttribute("code", "B");
            //PaymentCondition.AppendChild(PaymentConditionDescription);


            //Payment.AppendChild(PaymentCondition);


            //XmlNode CSSDYesNo = doc.CreateElement("CSSDYesNo");
            //CSSDYesNo.AppendChild(doc.CreateTextNode("K"));
            //Payment.AppendChild(CSSDYesNo);

            //XmlNode CSSDAmount1 = doc.CreateElement("CSSDAmount1");
            //CSSDAmount1.AppendChild(doc.CreateTextNode("0"));
            //Payment.AppendChild(CSSDAmount1);

            //XmlNode CSSDAmount2 = doc.CreateElement("CSSDAmount2");
            //CSSDAmount2.AppendChild(doc.CreateTextNode("0"));
            //Payment.AppendChild(CSSDAmount2);

            //XmlNode InvoiceNumber = doc.CreateElement("InvoiceNumber");
            ////select faktuurnr from gbkmut where dagbknr='202' დაიწყება 802-ით '80200001'

            //InvoiceNumber.AppendChild(doc.CreateTextNode(invoiceNumber));
            //Payment.AppendChild(InvoiceNumber);

            //XmlNode BankTransactionID = doc.CreateElement("BankTransactionID");
            ////new common Guid

            //BankTransactionID.AppendChild(doc.CreateTextNode(String.Format("{{{0}}}", commonId.ToString())));
            //Payment.AppendChild(BankTransactionID);


            //FinEntryLine.AppendChild(Payment);

            //XmlNode Delivery = doc.CreateElement("Delivery");

            //XmlNode DeliveryDate = doc.CreateElement("Date");
            //DeliveryDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            //Delivery.AppendChild(DeliveryDate);


            //FinEntryLine.AppendChild(Delivery);


            //XmlNode FinReferences = doc.CreateElement("FinReferences");
            //((XmlElement)FinReferences).SetAttribute("TransactionOrigin", "P");

            //XmlNode UniquePostingNumber = doc.CreateElement("UniquePostingNumber");
            //UniquePostingNumber.AppendChild(doc.CreateTextNode("0"));
            //FinReferences.AppendChild(UniquePostingNumber);

            //XmlNode YourRef = doc.CreateElement("YourRef");
            //YourRef.AppendChild(doc.CreateTextNode(""));
            //FinReferences.AppendChild(YourRef);

            //XmlNode FinReferencesDocumentDate = doc.CreateElement("DocumentDate");
            //FinReferencesDocumentDate.AppendChild(doc.CreateTextNode("2022-02-03"));
            //FinReferences.AppendChild(FinReferencesDocumentDate);

            //FinEntryLine.AppendChild(FinReferences);




            //XmlNode FinEntryLineDiscount = doc.CreateElement("Discount");

            //XmlNode FinEntryLinePercentage = doc.CreateElement("Percentage");
            //FinEntryLinePercentage.AppendChild(doc.CreateTextNode("0"));
            //FinEntryLineDiscount.AppendChild(Percentage);

            //FinEntryLine.AppendChild(FinEntryLineDiscount);


            //XmlNode FreeFields = doc.CreateElement("FreeFields");

            //XmlNode FreeTexts = doc.CreateElement("FreeTexts");
            //XmlNode FreeText = doc.CreateElement("FreeText");
            //FreeText.AppendChild(doc.CreateTextNode("00"));
            //((XmlElement)FreeText).SetAttribute("number", "3");
            //FreeTexts.AppendChild(FreeText);
            //FreeFields.AppendChild(FreeTexts);

            //FinEntryLine.AppendChild(FreeFields);

            return FinEntryLine;
        }

       
    }
}
