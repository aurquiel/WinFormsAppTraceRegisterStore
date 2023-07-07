using System.Collections.Generic;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataStatusReport
    {
        public int starep_id { get; set; }
        public string? starep_description { get; set; }
    }

    public class GetStatusReportOutput
    {
        public string? message { get; set; }
        public List<DataStatusReport>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
