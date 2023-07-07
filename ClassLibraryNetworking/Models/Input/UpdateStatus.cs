using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class UpdateStatus
    {
        public int sta_id { get; set; }
        public string sta_description { get; set; }
        public int sta_audit_id { get; set; }
        public System.DateTime sta_audit_date { get; set; }
        public bool sta_audit_delete { get; set; }
    }
}
