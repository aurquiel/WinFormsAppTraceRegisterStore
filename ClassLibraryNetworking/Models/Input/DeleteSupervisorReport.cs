using System;

namespace ClassLibraryNetworking.Models.Input
{
    public class DeleteStoreReport
    {
        public int storep_id { get; set; }
        public int storep_audit_id { get; set; }
        public DateTime storep_audit_date { get; set; }
        public bool storep_audit_delete { get; set; }
    }
}
