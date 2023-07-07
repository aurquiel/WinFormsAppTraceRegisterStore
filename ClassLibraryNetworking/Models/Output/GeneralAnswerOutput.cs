using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.Output
{
    public class GeneralAnswerOutput
    { 
        public string message { get; set; }
        public object data { get; set; }
        public bool statusOperation { get; set; }
    }
}
