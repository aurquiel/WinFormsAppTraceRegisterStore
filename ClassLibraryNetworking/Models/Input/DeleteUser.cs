using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteUser
    {
        public int usr_id { get; set; }
        public int usr_audit { get; set; }
        public System.DateTime usr_audit_date { get; set; }
        public bool usr_deleted { get; set; }
    }
}
