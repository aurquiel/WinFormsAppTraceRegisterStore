using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Output
{
    public class StoreReportExitsOutput
    {
        public bool exits { get; set; }
        public bool isProcessed { get; set; }
    }

    public class GetStoreReportExitsOutput
    {
        public string message { get; set; }
        public StoreReportExitsOutput data { get; set; }
        public bool statusOperation { get; set; }
    }
}
