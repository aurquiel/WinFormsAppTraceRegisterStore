using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataStoreReport
    {
        public int storep_id { get; set; }
        public int sto_code { get; set; }
        public DateTime storep_date { get; set; }
        public decimal storep_total_bs { get; set; }
        public decimal storep_change_bs { get; set; }
        public decimal storep_rate { get; set; }
        public decimal storep_equivalent_dollar { get; set; }
        public decimal storep_payed_euro { get; set; }
        public decimal storep_payed_zelle { get; set; }
        public decimal storep_payed_dollar { get; set; }
        public decimal storep_expended_dollar { get; set; }
        public decimal storep_total_dollar { get; set; }
        public int storep_sup_id { get; set; }
        public int storep_sta_id { get; set; }
        public string? storep_comments { get; set; }
        public string? starep_description { get; set; }
    }

    public class GetStoreReportsOutput
    {
        public string? message { get; set; }
        public List<DataStoreReport>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
