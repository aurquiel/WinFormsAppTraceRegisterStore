using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.MagickNumbers;
using ClassLibraryNetworking.Models.Output;
using ClassLibraryNetworking.Operations;

namespace WinFormsAppTrazoRegistroTienda
{
    public partial class MainForm : Form
    {
        private UserControl? actualUserControl = null;
        public EventHandler<Tuple<bool, string>>? RaiseRichTextInsertNewMessage;

        private NetworkingOperations _networkingOperations;
        private UserLoginInput _userLoginInput = new UserLoginInput();
        private UserLoginOutput _userLoginOutput = new UserLoginOutput();
        private GetStatusOutput _getStatusOutput = new GetStatusOutput();
        private GetSupervisorOutput _getSupervisorsOutput = new GetSupervisorOutput();
        private GetUsersOutput _getUsersOutput = new GetUsersOutput();
        private GetStoresOutput _getStoresOutput = new GetStoresOutput();
        private GetStatusReportOutput _getStatusReportOutput = new GetStatusReportOutput();
        private readonly int _TIMEOUT_MYSQL;

        public MainForm(NetworkingOperations networkingOperations, UserLoginInput userLoginInput,
            UserLoginOutput userLoginOutput, GetStatusOutput getStatusOutput,
            GetSupervisorOutput getSupervisorsOutput, GetUsersOutput getUsersOutput,
            GetStoresOutput getStoresOutput, GetStatusReportOutput getStatusReportOutput, int TIMEOUT_MYSQL)
        {
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                    ControlStyles.UserPaint |
                    ControlStyles.AllPaintingInWmPaint, true);
            _networkingOperations = networkingOperations;
            _userLoginInput = userLoginInput;
            _userLoginOutput = userLoginOutput;
            _getStatusOutput = getStatusOutput;
            _getSupervisorsOutput = getSupervisorsOutput;
            _getUsersOutput = getUsersOutput;
            _getStoresOutput = getStoresOutput;
            _getStatusReportOutput = getStatusReportOutput;
            _TIMEOUT_MYSQL = TIMEOUT_MYSQL;
            InitializeComponent();
            WireUpEvents();
            SelectChildForm();
        }

        private void WireUpEvents()
        {
            this.RaiseRichTextInsertNewMessage += UpdateOnRichTextInsertNewMessage;
        }

        private string GetTime()
        {
            DateTime time = DateTime.Now;

            return "Tiempo: " + time.ToString(@"hh\:mm\:ss\:fff") + ". ";
        }

        private void UpdateOnRichTextInsertNewMessage(object? sender, Tuple<bool, string> e)
        {
            if (e.Item1)
            {
                AppendText(true, e.Item2, Color.Green, true);
            }
            else
            {
                AppendText(false, e.Item2, Color.Red, true);
            }
        }

        private void AppendText(bool status, string text, Color color, bool addNewLine = false)
        {
            if (status)
            {
                text = text.Insert(0, "EXITOSO: " + GetTime());
            }
            else
            {
                text = text.Insert(0, "ERROR: " + GetTime());
            }
            richTextBoxStatusMessages.SuspendLayout();
            richTextBoxStatusMessages.SelectionColor = color;
            richTextBoxStatusMessages.AppendText(addNewLine
                ? $"{text}{Environment.NewLine}"
                : text);
            richTextBoxStatusMessages.ScrollToCaret();
            richTextBoxStatusMessages.ResumeLayout();
        }

        private void SelectChildForm()
        {
            if(_userLoginOutput.data.Any(x => x.usr_urol_id == (int)MagickInfo.ROL_ID_PERMITTED.STORE_ROL_ID))
            {
                OpenChildForm(new RegisterStoreUserControl(this, _networkingOperations, _userLoginInput, _userLoginOutput, 
                    _getStoresOutput, _getSupervisorsOutput, _getStatusOutput, _getStatusReportOutput, _TIMEOUT_MYSQL));
                UpdateOnRichTextInsertNewMessage(this, new Tuple<bool, string>(true, "Aplicacion modo Tienda."));
            }
            else if(_userLoginOutput.data.Any(x => x.usr_urol_id == (int)MagickInfo.ROL_ID_PERMITTED.SUPERVISOR_ROL_ID))
            {
                OpenChildForm(new RegisterSupervisorUserControl(this, _networkingOperations, _userLoginInput, _userLoginOutput, 
                    _getStoresOutput,  _getSupervisorsOutput, _getStatusReportOutput));
               UpdateOnRichTextInsertNewMessage(this, new Tuple<bool, string>(true, "Aplicacion modo Supervisor."));
            }
            else
            {
                OpenChildForm(new InitUserControl());
                UpdateOnRichTextInsertNewMessage(this, new Tuple<bool, string>(true, "Usuario sin los permisos necesarios."));
            }
        }

        private void OpenChildForm(UserControl childForm)
        {
            if (actualUserControl != null)
            {
                actualUserControl.Dispose();
            }
            actualUserControl = childForm;
            childForm.Dock = DockStyle.Fill;
            panelChildForm.Controls.Clear();
            panelChildForm.Controls.Add(childForm);
            childForm.BringToFront();
            childForm.Show();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}