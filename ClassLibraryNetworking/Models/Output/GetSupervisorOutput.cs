using Newtonsoft.Json;
using System.Collections.Generic;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataSupervisor
    {
        public int sup_id { get; set; }
        public string? sup_description { get; set; }
    }

    public class GetSupervisorOutput
    {
        [JsonIgnore]
        public EventHandler<Tuple<bool, GetSupervisorOutput, string>>? RaiseGetSupervisorOutputChanged;
        public string? message { get; set; }
        public List<DataSupervisor>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
