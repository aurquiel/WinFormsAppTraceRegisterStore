using ClassLibraryNetworking.Models.MagickNumbers;
using LogLibraryClassLibrary;
using MySql.Data.MySqlClient;

namespace MysqlClassLibrary
{
    public class CheckValuesMYSQL : Database
    {
        public CheckValuesMYSQL(int time_out) : base(time_out)
        {

        }

        public async Task<Tuple<bool, OutputTotalBsRate, string>> GetRegistersOfDay(DateTime date)
        {
            var salesAndRatesOfDay = await GetSaleOfDayAndRate(date);

            if(!salesAndRatesOfDay.Item1)
            {
                return salesAndRatesOfDay;
            }

            var getExcess = await GetExcess(date);

            if (!getExcess.Item1)
            {
                return new Tuple<bool, OutputTotalBsRate, string>(false, new OutputTotalBsRate(), getExcess.Item3);
            }

            if (getExcess.Item2 != 0)
            {
                var totalBsSum = salesAndRatesOfDay.Item2.ListValues.Sum(x => x.TotalBs);
                var rateAverage = salesAndRatesOfDay.Item2.ListValues.Sum(x => x.Rate) / salesAndRatesOfDay.Item2.ListValues.Count();

                var totalNewRecord = getExcess.Item2 - totalBsSum;

                var newOutput = new OutputTotalBsRate();

                foreach(var item in salesAndRatesOfDay.Item2.ListValues)
                {
                    newOutput.ListValues.Add(item); 
                }

                newOutput.ListValues.Add(new OutputTotalBsRate
                {
                    TotalBs = totalNewRecord,
                    Rate = rateAverage,
                    ProcessedStatus = (int)MagickInfo.STATUS_REPORT.PROCESADO
                }) ;

                return new Tuple<bool, OutputTotalBsRate, string>(true, newOutput, getExcess.Item3);
            }
            else
            {
                return salesAndRatesOfDay;
            }
        }

        public async Task<Tuple<bool, string>> GetCloseDay(DateTime date)
        {
            using (MySqlConnection mysqlConnection = new MySqlConnection(await GetSettingsConenction()))
            {
                try
                {
                    await mysqlConnection.OpenAsync();
                    string query = String.Format("Select count(*) as cantidad from dia_operativo where date(fecha)='{0}-{1}-{2}' and cerrado=1;", date.Year, date.Month.ToString("d2"), date.Day.ToString("d2"));
                    MySqlCommand command = new MySqlCommand(query, mysqlConnection);
                    command.CommandTimeout = SQL_TIMEOUT_EXECUTION_COMMAND;
                    command.CommandType = System.Data.CommandType.Text;

                    MySqlDataReader reader = command.ExecuteReader();

                    int result = 0;
                    while (reader.Read())
                    {
                        if (reader[0] == DBNull.Value)
                        {
                            result = 0;
                        }
                        else if (int.TryParse(reader.GetString(0), out result) == false)
                        {
                            result = 0;
                        }
                    }

                    if (result > 0)
                    {
                        return new Tuple<bool, string>(true, "Operacion exitosa dia cerrado, en MYSQL.");
                    }
                    else
                    {
                        return new Tuple<bool, string>(false, "Error al dia no cerrado, en MYSQL.");
                    }
                }
                catch (Exception ex)
                {
                    LoggerApp.WriteLineToLog("Error al verificar dia cerrado en MYSQL. Excepcion: " + ex.Message.ToLower() + " String Conexion: " + await GetStringConecctionNoPassword());
                    return new Tuple<bool, string>(false, "Error: " + ex.Message.ToLower() + " String Conexion: " + await GetStringConecctionNoPassword());
                }
            }
        }

        private async Task<Tuple<bool, OutputTotalBsRate, string>> GetSaleOfDayAndRate(DateTime date)
        {
            using (MySqlConnection mysqlConnection = new MySqlConnection(await GetSettingsConenction()))
            {
                try
                {
                    await mysqlConnection.OpenAsync();
                    string query = String.Format("SELECT DATE(fecha_creacion) AS Fecha, (SELECT ROUND(SUM(tasa) / COUNT(1), 2) AS tasa FROM tipo_efectivo v1 WHERE v1.codigo_factura = f.codigo_interno AND v1.tipo IN('USD', 'EUR', 'ZEL') AND DATE(f.fecha_creacion) = DATE(f.fecha_creacion)) AS tasa, ROUND(IFNULL(SUM(monto), 0), 2) AS Total FROM factura f, tipo_efectivo v WHERE estado = 'Facturada' AND v.codigo_factura = f.codigo_interno AND tipo IN('USD', 'EUR', 'ZEL') AND DATE(fecha_creacion) = DATE('{0}-{1}-{2}') GROUP BY DATE(fecha_creacion), tasa", date.Year, date.Month.ToString("d2"), date.Day.ToString("d2"));
                    MySqlCommand command = new MySqlCommand(query, mysqlConnection);
                    command.CommandTimeout = SQL_TIMEOUT_EXECUTION_COMMAND;
                    command.CommandType = System.Data.CommandType.Text;

                    MySqlDataReader reader = command.ExecuteReader();

                    OutputTotalBsRate output = new OutputTotalBsRate();

                    while (reader.Read())
                    {
                        output.ListValues.Add(new OutputTotalBsRate { 
                            Rate = decimal.Parse(reader.GetString(1)), 
                            TotalBs = decimal.Parse(reader.GetString(2)),
                            ProcessedStatus = (int)MagickInfo.STATUS_REPORT.NO_PROCESADO
                        });
                    }

                    if (output.ListValues.Count > 0)
                    {
                        return new Tuple<bool, OutputTotalBsRate, string>(true, output, "Operacion exitosa datos venta y tasa de MYSQL obtenidos.");
                    }
                    else
                    {
                        return new Tuple<bool, OutputTotalBsRate, string>(false, output, "Error al obtener total de venta y tasa de MYSQL.");
                    }
                }
                catch (Exception ex)
                {
                    LoggerApp.WriteLineToLog("Error al traer venta del dia. Excepcion: " + ex.Message.ToLower() + " String Conexion: " + await GetStringConecctionNoPassword());
                    return new Tuple<bool, OutputTotalBsRate, string>(false, new OutputTotalBsRate(), "Error: " + ex.Message.ToLower() + " String Conexion: " + await GetStringConecctionNoPassword());
                }
            }
        }

        private async Task<Tuple<bool, decimal, string>> GetExcess(DateTime date)
        {
            using (MySqlConnection mysqlConnection = new MySqlConnection(await GetSettingsConenction()))
            {
                try
                {
                    await mysqlConnection.OpenAsync();
                    string query = String.Format("select sum(monto) from deposito where fecha ='{0}-{1}-{2}' and (numero like 'EDD%' or numero like 'EDE%' or numero like 'ZELLE%');", date.Year, date.Month.ToString("d2"), date.Day.ToString("d2"));
                    MySqlCommand command = new MySqlCommand(query, mysqlConnection);
                    command.CommandTimeout = SQL_TIMEOUT_EXECUTION_COMMAND;
                    command.CommandType = System.Data.CommandType.Text;

                    MySqlDataReader reader = command.ExecuteReader();
                    decimal result = 0;
                    while (reader.Read())
                    {
                        if (reader[0] == DBNull.Value)
                        {
                            result = 0;
                        }
                        else if (!decimal.TryParse(reader.GetString(0), out result))
                        {
                            result = 0;
                        }
                        
                    }

                    return new Tuple<bool, decimal, string>(true, result, "Operacion exitosa datos venta de exedentes de MYSQL obtenidos.");
                    
                }
                catch (Exception ex)
                {
                    LoggerApp.WriteLineToLog("Error al traer venta del dia, tipo excendentes. Excepcion: " + ex.Message.ToLower());
                    return new Tuple<bool, decimal, string>(false, new decimal(), "Error: " + ex.Message.ToLower());
                }
            }
        }
    }
}
