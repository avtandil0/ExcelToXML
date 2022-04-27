using ExcelToXML.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToXML.Auth
{
    public interface IUserRepository
    {
        UserDTO GetUser(UserDTO userMode);
    }
}
