using ClassLibraryNetworking.Models.Input;
using Newtonsoft.Json;

namespace ClassLibraryNetworking.Models.Output
{
    public class DataUsers
    {
        public int usr_id { get; set; }
        public string? usr_alias { get; set; }
        public string? usr_password { get; set; }
        public string? usr_firstname { get; set; }
        public string? usr_lastname { get; set; }
        public int usr_manage_id { get; set; }
        public string? usr_email { get; set; }
        public int usr_urol_id { get; set; }
    }

    public class GetUsersOutput
    {
        [JsonIgnore]
        public EventHandler<Tuple<bool, GetUsersOutput, string>>? RaiseGetUsersOutputChanged;
        public string? message { get; set; }
        public List<DataUsers>? data { get; set; }
        public bool statusOperation { get; set; }
    }
}
