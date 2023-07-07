using LogLibraryClassLibrary;
using System.Configuration;
using System.Diagnostics;
using System.Linq;

namespace MysqlClassLibrary
{
    public class Database
    {
        public int SQL_TIMEOUT_EXECUTION_COMMAND = 10; //By default is 10 seconds

        public Database(int time_out)
        {
            try
            {
                SQL_TIMEOUT_EXECUTION_COMMAND = time_out;
            }
            catch (Exception ex)
            {
                SQL_TIMEOUT_EXECUTION_COMMAND = 10;
                //Logger.WriteToLog("Metodo: " + ex.TargetSite + ", Error: " + ex.Message.ToLower());
            }
        }

        protected async Task<string> GetStringConecctionNoPassword()
        {
            try
            {
                string FILE_PATH = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/TrazoRegistrosTienda/mysql.txt";
                string[] lines = await File.ReadAllLinesAsync(FILE_PATH);
                lines = lines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

                return lines.Length > 0 ?
                    String.Join(";", lines[0].Split(';').Where(x => !x.Contains("password")).ToList()) + ". No se muestra contraseña por seguridad."
                    : 
                    string.Empty;
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Excepcion: " + ex.Message.ToLower());
                return string.Empty;
            }
        }

        protected async Task<string> GetSettingsConenction()
        {
            try
            {
                string FILE_PATH = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/TrazoRegistrosTienda/mysql.txt";
                string[] lines = await File.ReadAllLinesAsync(FILE_PATH);
                lines  = lines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                return lines.Length > 0 ? lines[0] : string.Empty;  
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Error al tomar string conmexion Mysql del archivo de configuracion. Excepcion: " + ex.Message.ToLower());
                return string.Empty;
            }
        }
    }
}