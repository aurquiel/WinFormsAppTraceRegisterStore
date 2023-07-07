using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteStoreInput
    {
        public UserLoginInput userLoginInput { get; set; }
        public DeleteStore deleteStore { get; set; }
    }
}
