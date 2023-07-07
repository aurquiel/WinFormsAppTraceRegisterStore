using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsAppTrazoRegistroTienda.Excel
{
    internal class ExcelReportStoreDTO
    {
        public int storep_code { get; set; }
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
        public string? storep_sta_description { get; set; }
        public string? storep_supervisor { get; set; }
        public string? storep_comments { get; set; }
        public string? storep_starep_description { get; set; }
    }

    internal class ExcelReportStoreTotalDTO
    {
        public decimal total_storep_total_bs { get; set; }
        public decimal total_storep_change_bs { get; set; }
        public decimal total_storep_equivalent_dollar { get; set; }
        public decimal total_storep_payed_euro { get; set; }
        public decimal total_storep_payed_zelle { get; set; }
        public decimal total_storep_payed_dollar { get; set; }
        public decimal total_storep_expended_dollar { get; set; }
        public decimal total_storep_total_dollar { get; set; }
    }
}
