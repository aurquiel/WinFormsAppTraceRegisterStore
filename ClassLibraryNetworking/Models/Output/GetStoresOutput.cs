using Newtonsoft.Json;
using System.Collections.Generic;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataStores
    {
        public int sto_id { get; set; }
        public int sto_code { get; set; }
        public int sto_sup_id { get; set; }
    }

    public class GetStoresOutput
    {
        [JsonIgnore]
        public EventHandler<Tuple<bool, GetStoresOutput, string>>? RaiseGetStoresOutputChanged;
        public string? message { get; set; }
        public List<DataStores>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
