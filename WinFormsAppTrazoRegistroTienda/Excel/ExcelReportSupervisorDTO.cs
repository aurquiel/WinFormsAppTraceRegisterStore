using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsAppTrazoRegistroTienda.Excel
{
    internal class ExcelReportSupervisorDTO
    {
        public string? suprep_sup_description { get; set; }
        public DateTime suprep_date { get; set; }
        public int suprep_sto_code { get; set; }
        public decimal suprep_in_dollar { get; set; }
        public decimal suprep_out_dollar { get; set; }
        public decimal suprep_in_euro { get; set; }
        public decimal suprep_out_euro { get; set; }
        public decimal suprep_in_peso { get; set; }
        public decimal suprep_out_peso { get; set; }
        public string? suprep_comments { get; set; }
        public string? suprep_starep_description { get; set; }
    }

    internal class ExcelReportSupervisorTotalDTO
    {
        public decimal total_suprep_in_dollar { get; set; }
        public decimal total_suprep_out_dollar { get; set; }
        public decimal total_suprep_in_euro { get; set; }
        public decimal total_suprep_out_euro { get; set; }
        public decimal total_suprep_in_peso { get; set; }
        public decimal total_suprep_out_peso { get; set; }
    }
}
