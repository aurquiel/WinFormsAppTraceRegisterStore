using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataSupervisorReport
    {
        public int suprep_id { get; set; }
        public int suprep_sup_id { get; set; }
        public DateTime suprep_date { get; set; }
        public int suprep_sto_id { get; set; }
        public decimal suprep_in_dollar { get; set; }
        public decimal suprep_out_dollar { get; set; }
        public decimal suprep_in_euro { get; set; }
        public decimal suprep_out_euro { get; set; }
        public decimal suprep_in_peso { get; set; }
        public decimal suprep_out_peso { get; set; }
        public string suprep_comments { get; set; }
        public int suprep_starep_id { get; set; }
    }

    public class GetSupervisorReportsOutput
    {
        public string message { get; set; }
        public List<DataSupervisorReport> data { get; set; }
        public bool statusOperation { get; set; }
    }
}
