using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.MagickNumbers;
using ClassLibraryNetworking.Models.Output;
using ClassLibraryNetworking.Operations;
using LogLibraryClassLibrary;
using MySqlX.XDevAPI.Common;
using System.Configuration;
using System.Diagnostics;

namespace WinFormsAppTrazoRegistroTienda
{
    public partial class LoginForm : Form
    {
        private readonly string _IP_WEB_SERVICE;
        private readonly int _TIMEOUT_WEB_SERVICE;
        private readonly int _TIMEOUT_MYSQL;
        private NetworkingOperations _networkingOperations;
        private UserLoginInput _userLoginInput = new UserLoginInput();

        private UserLoginOutput _userLoginOutput = new UserLoginOutput();
        private GetStatusOutput _getStatusOutput = new GetStatusOutput(); 
        private GetSupervisorOutput _getSupervisorsOutput = new GetSupervisorOutput(); 
        private GetUsersOutput _getUsersOutput = new GetUsersOutput(); 
        private GetStoresOutput _getStoresOutput = new GetStoresOutput(); 
        private GetStatusReportOutput _getStatusReportOutput = new GetStatusReportOutput(); 

        

        public LoginForm()
        {
            InitializeComponent();

            LoggerApp.CreateLog();

            try
            {
                _IP_WEB_SERVICE = ConfigurationManager.AppSettings["IP_WEB_SERVICE"].ToString();
            }
            catch(Exception ex)
            {
                LoggerApp.WriteLineToLog("Error al tomar IP webservice del archivo de configuracion, ip por default 200.35.195.130:9095. Excepcion: " + ex.Message.ToLower());
                _IP_WEB_SERVICE = "http://200.35.195.130:9095/";
            }
            try
            {
                _TIMEOUT_WEB_SERVICE = Int32.Parse(ConfigurationManager.AppSettings["TIMEOUT_WEB_SERVICE"].ToString());
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Error al tomar TIMEOUT webservice del archivo de configuracion, ip por default 30s. Excepcion: " + ex.Message.ToLower());
                _TIMEOUT_WEB_SERVICE = 30;
            }
            try
            {
                _TIMEOUT_MYSQL = Int32.Parse(ConfigurationManager.AppSettings["TIMEOUT_MYSQL"].ToString());
            }
            catch (Exception ex)
            {
                LoggerApp.WriteLineToLog("Error al tomar TIMEOUT Mysql del archivo de configuracion, ip por default 10s. Excepcion: " + ex.Message.ToLower());
                _TIMEOUT_MYSQL = 10;
            }

            _networkingOperations = new NetworkingOperations(_IP_WEB_SERVICE, _TIMEOUT_WEB_SERVICE);
        }

        private void UpdateUiFromLoadStore(Tuple<bool, string> result)
        {
            if (result.Item1)
            {
                TextBoxStatusUpdate(result.Item2, Color.Green);
            }
            else
            {
                TextBoxStatusUpdate(result.Item2, Color.Red);
            }
        }

        private void TextBoxStatusUpdate(string text, Color colorBrush)
        {
            this.textBoxStatus.Text = text;
            this.textBoxStatus.ForeColor = colorBrush;
            this.textBoxStatus.BackColor = this.textBoxStatus.BackColor;
        }

        private Tuple<bool, string> ValidateInputUser()
        {
            if (string.IsNullOrWhiteSpace(textBoxUserAlias.Text))
            {
                return new Tuple<bool, string>(false, "Error campo de usuario vacio.");
            }
            else if (string.IsNullOrWhiteSpace(textBoxUserPassword.Text))
            {
                return new Tuple<bool, string>(false, "Error campo de contraseña de usuario vacio.");
            }
            else
            {
                return new Tuple<bool, string>(true, "Entrada de datos de usuarios validados.");
            }
        }

        private async void buttonLogin_Click(object sender, EventArgs e)
        {
            buttonLogin.Enabled = false;
            UpdateUiFromLoadStore(new Tuple<bool, string>(true, "Obteniendo data del servidor..."));

            var resultValidateUser = ValidateInputUser();

            if (resultValidateUser.Item1 == false)
            {
                UpdateUiFromLoadStore(resultValidateUser);
                buttonLogin.Enabled = true;
                return;
            }

            _userLoginInput.alias = textBoxUserAlias.Text;
            _userLoginInput.password = textBoxUserPassword.Text;

            var resultValidateUserDatabase = await _networkingOperations.LoginUser(_userLoginInput);
            _userLoginOutput = resultValidateUserDatabase.Item2;
            if (resultValidateUserDatabase.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultValidateUserDatabase.Item1, resultValidateUserDatabase.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            var resultValidateRol = ValidateRolUser();
            if (resultValidateRol.Item1 == false)
            {
                UpdateUiFromLoadStore(resultValidateRol);
                buttonLogin.Enabled = true;
                return;
            }

            var resultGetStatus= await _networkingOperations.GetStatus(_userLoginInput);
            _getStatusOutput = resultGetStatus.Item2;
            if (resultGetStatus.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultGetStatus.Item1, resultGetStatus.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            var resultGetSupervisors = await _networkingOperations.GetSupervisors(_userLoginInput);
            _getSupervisorsOutput = resultGetSupervisors.Item2;
            if (resultGetSupervisors.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultGetSupervisors.Item1, resultGetSupervisors.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            var resultGetStores = await _networkingOperations.GetStores(_userLoginInput);
            _getStoresOutput = resultGetStores.Item2;
            if (resultGetStores.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultGetStores.Item1, resultGetStores.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            var resultGetUsers = await _networkingOperations.GetUsers(_userLoginInput);
            _getUsersOutput = resultGetUsers.Item2;
            if (resultGetUsers.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultGetUsers.Item1, resultGetUsers.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            var resultGetStatusReport = await _networkingOperations.GetStatusReport(_userLoginInput);
            _getStatusReportOutput = resultGetStatusReport.Item2;
            if (resultGetStatusReport.Item1 == false)
            {
                UpdateUiFromLoadStore(new Tuple<bool, string>(resultGetStatusReport.Item1, resultGetStatusReport.Item3));
                buttonLogin.Enabled = true;
                return;
            }

            LaunchMainWindows();

            buttonLogin.Enabled = true;
            textBoxStatus.Text = string.Empty;
            textBoxUserAlias.Text = string.Empty;
            textBoxUserPassword.Text = string.Empty;
        }

        private Tuple<bool, string> ValidateRolUser()
        {
            if(_userLoginOutput.data.Any(x => x.usr_urol_id == (int)MagickInfo.ROL_ID_PERMITTED.STORE_ROL_ID))
            {
                return new Tuple<bool, string>(true, "Rol de usuario valido.");
            }
            else if (_userLoginOutput.data.Any(x => x.usr_urol_id == (int)MagickInfo.ROL_ID_PERMITTED.SUPERVISOR_ROL_ID))
            {
                return new Tuple<bool, string>(true, "Rol de usuario valido.");
            }
            else
            {
                return new Tuple<bool, string>(false, "Rol de usuario sin permisos necesarios.");
            } 
        }

        private void LaunchMainWindows()
        {
            this.Hide();
            MainForm mainForm = new MainForm(_networkingOperations, _userLoginInput,_userLoginOutput, _getStatusOutput,
                _getSupervisorsOutput, _getUsersOutput, _getStoresOutput, _getStatusReportOutput, _TIMEOUT_MYSQL);
            mainForm.Closed += (s, args) => this.Show();                          
            mainForm.ShowDialog();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
