using ClassLibraryNetworking.Models.Input;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace ClassLibraryNetworking.Models.Output
{

    public class DataLogin
    {
        public int usr_id { get; set; }
        public int usr_urol_id { get; set; }
        public string? usr_urol_uro_name { get; set; }
        public int usr_manage_id { get; set; }
    }

    public class UserLoginOutput
    {
        [JsonIgnore]
        public EventHandler<Tuple<bool, UserLoginOutput, string>>? RaiseUserLoginOutputChanged;
        public string? message { get; set; }
        public List<DataLogin>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
