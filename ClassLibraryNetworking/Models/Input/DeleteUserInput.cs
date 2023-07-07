using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteUserInput
    {
        public UserLoginInput userLoginInput { get; set; }
        public DeleteUser deleteUser { get; set; }
    }
}
