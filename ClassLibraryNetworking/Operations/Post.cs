using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.Output;
using LogLibraryClassLibrary;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Text;

namespace ClassLibraryNetworking.Operations
{
    internal class Post
    {
        #region Login

        internal static async Task<Tuple<bool, UserLoginOutput, string>> LoginUser(string url, int timeOut, UserLoginInput userLogin)
        {
            try
            {
                string json = JsonConvert.SerializeObject(userLogin);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/access/login", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, UserLoginOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<UserLoginOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, UserLoginOutput, string>(
                    false,
                    new UserLoginOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());

            }
        }

        #endregion Login

        #region StatusReport

        internal static async Task<Tuple<bool, GetStatusReportOutput, string>> GetStatusReport(string url, int timeOut, UserLoginInput userLogin)
        {
            try
            {
                string json = JsonConvert.SerializeObject(userLogin);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/statusreport/getstatusreport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetStatusReportOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetStatusReportOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetStatusReportOutput, string>(
                    false,
                    new GetStatusReportOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion StatusReport

        #region StoreReport

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStoreReport(string url, int timeOut, InsertStoreReportInput insertStoreReportInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertStoreReportInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/storereport/insertreport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");

            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GetStoreReportExitsOutput, string>> GetStoreReportExits(string url, int timeOut, GetStoreReportInput getReportStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(getReportStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/storereport/getreportexits", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetStoreReportExitsOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetStoreReportExitsOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetStoreReportExitsOutput, string>(
                    false,
                    new GetStoreReportExitsOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GetStoreReportsOutput, string>> GetStoreReport(string url, int timeOut, GetStoreReportInput getReportStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(getReportStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/storereport/getreports", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetStoreReportsOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetStoreReportsOutput>(result),
                    "Respuesta del servidor obtenida con exito.");

            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetStoreReportsOutput, string>(
                    false,
                    new GetStoreReportsOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStoreReport(string url, int timeOut, UpdateStoreReportInput updateReportStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateReportStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/storereport/updatereport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStoreReport(string url, int timeOut, DeleteStoreReportInput deleteReportStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deleteReportStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/storereport/deletereport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion StoreReport

        #region SupervisorReport

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertSupervisorReport(string url, int timeOut, InsertSupervisorReportInput insertSupervisorReportInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertSupervisorReportInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisorreport/insertreport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GetSupervisorReportsOutput, string>> GetSupervisorReport(string url, int timeOut, GetSupervisorReportInput getReportStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(getReportStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisorreport/getreports", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetSupervisorReportsOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetSupervisorReportsOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetSupervisorReportsOutput, string>(
                    false,
                    new GetSupervisorReportsOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateSupervisorReport(string url, int timeOut, UpdateSupervisorReportInput updateSupervisorReportInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateSupervisorReportInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisorreport/updatereport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }


        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteSupervisorReport(string url, int timeOut, DeleteSupervisorReportInput deletyeSupervisorReportInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deletyeSupervisorReportInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisorreport/deletereport", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion SupervisorReport

        #region Stores

        internal static async Task<Tuple<bool, GetStoresOutput, string>> GetStores(string url, int timeOut, UserLoginInput userLoginInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(userLoginInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/store/getstores", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetStoresOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetStoresOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetStoresOutput, string>(
                    false,
                    new GetStoresOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStore(string url, int timeOut, DeleteStoreInput deleteStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deleteStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/store/deletestore", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStore(string url, int timeOut, UpdateStoreInput updateStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/store/updatestore", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStore(string url, int timeOut, InsertStoreInput insertStoreInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertStoreInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/store/insertstore", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion Stores

        #region Supervisors

        internal static async Task<Tuple<bool, GetSupervisorOutput, string>> GetSupervisors(string url, int timeOut, UserLoginInput userLoginInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(userLoginInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisor/getsupervisors", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetSupervisorOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetSupervisorOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetSupervisorOutput, string>(
                    false,
                    new GetSupervisorOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateSupervisor(string url, int timeOut, UpdateSupervisorInput updateSupervisorInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateSupervisorInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisor/updatesupervisor", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }


        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteSupervisor(string url, int timeOut, DeleteSupervisorInput deleteSupervisorInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deleteSupervisorInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisor/deletesupervisor", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertSupervisor(string url, int timeOut, InsertSupervisorInput insertSupervisorInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertSupervisorInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/supervisor/insertsupervisor", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion Supervisors

        #region Status

        internal static async Task<Tuple<bool, GetStatusOutput, string>> GetStatus(string url, int timeOut, UserLoginInput userLoginInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(userLoginInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/status/getstatus", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetStatusOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetStatusOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetStatusOutput, string>(
                    false,
                    new GetStatusOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStatus(string url, int timeOut, UpdateStatusInput updateStatusInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateStatusInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/status/updatestatus", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");

            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStatus(string url, int timeOut, DeleteStatusInput deleteStatusInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deleteStatusInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/status/deletestatus", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStatus(string url, int timeOut, InsertStatusInput insertStatusInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertStatusInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/status/insertstatus", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion Status

        #region Users

        internal static async Task<Tuple<bool, GetUsersOutput, string>> GetUsers(string url, int timeOut, UserLoginInput updateUserInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateUserInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/user/getusers", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GetUsersOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GetUsersOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GetUsersOutput, string>(
                    false,
                    new GetUsersOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateUser(string url, int timeOut, UpdateUserInput updateUserInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(updateUserInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/user/updateuser", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteUser(string url, int timeOut, DeleteUserInput deleteUserInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(deleteUserInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/user/deleteuser", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>(
                    true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        internal static async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertUser(string url, int timeOut, InsertUserInput insertUserInput)
        {
            try
            {
                string json = JsonConvert.SerializeObject(insertUserInput);

                var data = new StringContent(json, Encoding.UTF8, "application/json");

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(timeOut);

                var response = await client.PostAsync(url + "/api/user/insertuser", data);

                var result = await response.Content.ReadAsStringAsync();

                return new Tuple<bool, GeneralAnswerOutput, string>
                    (true,
                    JsonConvert.DeserializeObject<GeneralAnswerOutput>(result),
                    "Respuesta del servidor obtenida con exito.");
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Exception: " + ex.Message);
                return new Tuple<bool, GeneralAnswerOutput, string>(
                    false,
                    new GeneralAnswerOutput(),
                    "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        #endregion Status
    }
}
