using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToXML.Models
{
    public class ExcelData
    {
        public Guid ID { get; set; }
        public string StatementNumber { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }
        public string OperationContent { get; set; }
        public string OperationType { get; set; }
        public string ReceiverName { get; set; }
        public string IdentityNumber { get; set; }
    }
}
