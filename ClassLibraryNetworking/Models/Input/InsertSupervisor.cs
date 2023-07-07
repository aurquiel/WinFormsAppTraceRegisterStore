using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class InsertSupervisor
    {
        public string sup_description { get; set; }
        public int sup_audit_id { get; set; }
        public System.DateTime sup_audit_date { get; set; }
        public bool sup_audit_delete { get; set; }
    }
}
