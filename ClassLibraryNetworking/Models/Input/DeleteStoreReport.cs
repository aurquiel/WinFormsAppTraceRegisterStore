using System;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteSupervisorReport
    {
        public int suprep_id { get; set; }
        public int suprep_audit_id { get; set; }
        public System.DateTime suprep_audit_date { get; set; }
        public bool suprep_audit_delete { get; set; }
    }
}
