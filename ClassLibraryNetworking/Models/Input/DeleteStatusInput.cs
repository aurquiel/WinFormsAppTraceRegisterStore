using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteStatusInput
    {
        public UserLoginInput userLoginInput { get; set; }
        public DeleteStatus deleteStatus { get; set; }
    }
}
