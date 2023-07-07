using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.Output;

namespace ClassLibraryNetworking.Operations
{
    public class NetworkingOperations
    {
        private readonly string _IP_WEB_SERVICE;
        private readonly int _TIMEOUT_WEB_SERVICE;

        public NetworkingOperations(string IP_WEB_SERVICE, int TIMEOUT_WEB_SERVICE)
        {
            _IP_WEB_SERVICE = IP_WEB_SERVICE;
            _TIMEOUT_WEB_SERVICE = TIMEOUT_WEB_SERVICE;
        }

        #region UserLogin

        public async Task<Tuple<bool, UserLoginOutput, string>> LoginUser(UserLoginInput userLoginInput)
        {
            Tuple<bool, UserLoginOutput, string> result = await Post.LoginUser(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, UserLoginOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, UserLoginOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, UserLoginOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion UserLogin

        #region Status
        public async Task<Tuple<bool, GetStatusOutput, string>> GetStatus(UserLoginInput userLoginInput)
        {
            Tuple<bool, GetStatusOutput, string> result = await Post.GetStatus(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetStatusOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetStatusOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetStatusOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStatus(UpdateStatusInput updateStatusInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateStatus(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateStatusInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStatus(DeleteStatusInput deleteStatusInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteStatus(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteStatusInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStatus(InsertStatusInput insertStatusInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertStatus(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertStatusInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion Status

        #region Supervisors
        public async Task<Tuple<bool, GetSupervisorOutput, string>> GetSupervisors(UserLoginInput userLoginInput)
        {
            Tuple<bool, GetSupervisorOutput, string> result = await Post.GetSupervisors(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetSupervisorOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetSupervisorOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetSupervisorOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateSupervisor(UpdateSupervisorInput updateSupervisorInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateSupervisor(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateSupervisorInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteSupervisor(DeleteSupervisorInput deleteSupervisorInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteSupervisor(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteSupervisorInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertSupervisor(InsertSupervisorInput insertSupervisorInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertSupervisor(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertSupervisorInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion Supervisors

        #region Stores
        public async Task<Tuple<bool, GetStoresOutput, string>> GetStores(UserLoginInput userLoginInput)
        {
            Tuple<bool, GetStoresOutput, string> result = await Post.GetStores(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetStoresOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetStoresOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetStoresOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStore(UpdateStoreInput updateStoreInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateStore(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateStoreInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStore(DeleteStoreInput deleteStoreInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteStore(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteStoreInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStore(InsertStoreInput insertStoreInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertStore(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertStoreInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion Stores

        #region Users
        public async Task<Tuple<bool, GetUsersOutput, string>> GetUsers(UserLoginInput userLoginInput)
        {
            Tuple<bool, GetUsersOutput, string> result = await Post.GetUsers(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetUsersOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetUsersOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetUsersOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateUser(UpdateUserInput updateUserInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateUser(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateUserInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteUser(DeleteUserInput deleteUserInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteUser(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteUserInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertUser(InsertUserInput insertUserInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertUser(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertUserInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion Users

        #region StatusReport
        public async Task<Tuple<bool, GetStatusReportOutput, string>> GetStatusReport(UserLoginInput userLoginInput)
        {
            Tuple<bool, GetStatusReportOutput, string> result = await Post.GetStatusReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, userLoginInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetStatusReportOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetStatusReportOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetStatusReportOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion StatusReport

        #region StoreReport

        public async Task<Tuple<bool, GetStoreReportsOutput, string>> GetStoreReport(GetStoreReportInput getStoreReportInput)
        {
            Tuple<bool, GetStoreReportsOutput, string> result = await Post.GetStoreReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, getStoreReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetStoreReportsOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetStoreReportsOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetStoreReportsOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateStoreReport(UpdateStoreReportInput updateStoreReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateStoreReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateStoreReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteStoreReport(DeleteStoreReportInput deleteStoreReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteStoreReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteStoreReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertStoreReport(InsertStoreReportInput insertStoreReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertStoreReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertStoreReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GetStoreReportExitsOutput, string>> GetStoreReportExits(GetStoreReportInput  getStoreReportInput)
        {
            Tuple<bool, GetStoreReportExitsOutput, string> result = await Post.GetStoreReportExits(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, getStoreReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetStoreReportExitsOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetStoreReportExitsOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetStoreReportExitsOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion StoreReport

        #region SupervisorReport

        public async Task<Tuple<bool, GetSupervisorReportsOutput, string>> GetSupervisorReport(GetSupervisorReportInput getSupervisorReportInput)
        {
            Tuple<bool, GetSupervisorReportsOutput, string> result = await Post.GetSupervisorReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, getSupervisorReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GetSupervisorReportsOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GetSupervisorReportsOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GetSupervisorReportsOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> UpdateSupervisorReport(UpdateSupervisorReportInput updateSupervisorReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.UpdateSupervisorReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, updateSupervisorReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> DeleteSupervisorReport(DeleteSupervisorReportInput deleteSupervisorReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.DeleteSupervisorReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, deleteSupervisorReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        public async Task<Tuple<bool, GeneralAnswerOutput, string>> InsertSupervisorReport(InsertSupervisorReportInput insertSupervisorReportInput)
        {
            Tuple<bool, GeneralAnswerOutput, string> result = await Post.InsertSupervisorReport(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE, insertSupervisorReportInput);

            if (result.Item1 && result.Item2.statusOperation)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(true, result.Item2, result.Item2.message);
            }
            else if (result.Item1 && result.Item2.statusOperation == false)
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item2.message);
            }
            else
            {
                return new Tuple<bool, GeneralAnswerOutput, string>(false, result.Item2, result.Item3);
            }
        }

        #endregion SupervisorReport
    }
}
