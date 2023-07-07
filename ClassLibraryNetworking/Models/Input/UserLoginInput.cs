using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Input
{
    public class UserLoginInput
    {
        [JsonIgnore]
        public EventHandler<UserLoginInput>? RaiseUserLoginInputChanged;
        public string? alias { get; set; }
        public string? password { get; set; }
    }
}
