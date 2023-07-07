using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class UpdateStatusInput
    {
        public UserLoginInput userLoginInput { get; set; }
        public UpdateStatus updateStatus { get; set; }
    }
}
