using ExcelToXML.Models;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToXML.Auth
{
    public class UserRepository : IUserRepository
    {
        private readonly List<UserDTO> users = new List<UserDTO>();

        public IConfiguration Configuration { get; }
        public UserRepository(IConfiguration configuration)
        {
            Configuration = configuration;

            var us = configuration.GetSection("Users").Get<List<UserDTO>>();
            users.AddRange(us);
        }


        public UserDTO GetUser(UserDTO userModel)
        {
            return users.Where(x => x.UserName.ToLower() == userModel.UserName.ToLower()
                && x.Password == userModel.Password).FirstOrDefault();
        }
    }

}
