using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteStore
    {
        public int sto_id { get; set; }
        public int sto_audit_id { get; set; }
        public System.DateTime sto_audit_date { get; set; }
        public bool sto_audit_delete { get; set; }
    }
}
