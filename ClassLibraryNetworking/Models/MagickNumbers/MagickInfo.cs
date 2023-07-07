using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibraryNetworking.Models.MagickNumbers
{
    public static class MagickInfo
    {
        public enum ROL_ID_PERMITTED { STORE_ROL_ID = 10, SUPERVISOR_ROL_ID = 9, ADMIN_ROL_ID = 1 } //Estos son los id en la base de datos
        public enum STATUS_REPORT { PROCESADO = 1, NO_PROCESADO = 2 } //Estos son los id en la base de datos
        public enum STATUS { CIERRE_DE_MES = 5, S_E = 3, AJUSTE_CON_EL_CIERRE = 16} //Estos son los id en la base de datos
        public enum SUPERVISOR { ADMINISTRADOR = 15 } //Estos son los id en la base de datos
        public enum STORE { ZERO = 81 } //Estos son los id en la base de datos
    }
}
