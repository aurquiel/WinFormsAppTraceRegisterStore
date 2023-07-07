using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteSupervisorInput
    {
        public UserLoginInput userLoginInput { get; set; }
        public DeleteSupervisor deleteSupervisor { get; set; }
    }
}
