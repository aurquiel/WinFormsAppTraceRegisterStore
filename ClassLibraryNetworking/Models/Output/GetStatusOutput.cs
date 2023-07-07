using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataStatus
    {
        public int sta_id { get; set; }
        public string? sta_description { get; set; }
    }

    public class GetStatusOutput
    {
        [JsonIgnore]
        public EventHandler<Tuple<bool, GetStatusOutput, string>>? RaiseGetStatusOutputChanged;
        public string? message { get; set; }
        public List<DataStatus>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
