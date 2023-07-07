using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.MagickNumbers;
using ClassLibraryNetworking.Models.Output;
using ClassLibraryNetworking.Operations;
using System.Reflection;
using WinFormsAppTrazoRegistroTienda.Excel;

namespace WinFormsAppTrazoRegistroTienda
{
    public partial class RegisterSupervisorUserControl : UserControl
    {
        private MainForm _mainForm;
        private NetworkingOperations _networkingOperations;
        private UserLoginInput _userLoginInput;
        private UserLoginOutput _userLoginOutput;
        private GetStoresOutput _getStoresOutput;
        private GetSupervisorOutput _getSupervisorOutput;
        private GetStatusReportOutput _getStatusReportOutput;

        private ToolTip buttonbuttonSupervisorReportAddTooltip = new ToolTip();
        private ToolTip buttonbuttonSupervisorReportAddCleanTooltip = new ToolTip();
        private ToolTip buttonSupervisorReportEditConsultTooltip = new ToolTip();
        private ToolTip buttonSupervisorReportEditTooltip = new ToolTip();
        private ToolTip buttonSupervisorReportDeleteTooltip = new ToolTip();
        private ToolTip buttonSupervisorReportEditCleanTooltip = new ToolTip();
        private ToolTip buttonSupervisorReportEditExcelTooltip = new ToolTip();

        public RegisterSupervisorUserControl(MainForm mainForm, NetworkingOperations networkingOperations, UserLoginInput userLoginInput, UserLoginOutput userLoginOutput,
            GetStoresOutput getStoresOutput, GetSupervisorOutput getSupervisorOutput, GetStatusReportOutput getStatusReportOutput)
        {
            InitializeComponent();
            _mainForm = mainForm;
            _networkingOperations = networkingOperations;
            _userLoginInput = userLoginInput;
            _userLoginOutput = userLoginOutput;
            _getStoresOutput = getStoresOutput;
            _getSupervisorOutput = getSupervisorOutput;
            _getStatusReportOutput = getStatusReportOutput;

            LoadDataGridViewAddSupervisorReport();

            buttonbuttonSupervisorReportAddTooltip.SetToolTip(buttonbuttonSupervisorReportAdd, "Agregar Reporte");
            buttonbuttonSupervisorReportAddCleanTooltip.SetToolTip(buttonbuttonSupervisorReportAddClean, "Limpiar Campos");
            buttonSupervisorReportEditConsultTooltip.SetToolTip(buttonSupervisorReportEditConsult, "Consultar Registros");
            buttonSupervisorReportEditTooltip.SetToolTip(buttonSupervisorReportEdit, "Editar Registros");
            buttonSupervisorReportDeleteTooltip.SetToolTip(buttonSupervisorReportDelete, "Eliminar Registros");
            buttonSupervisorReportEditCleanTooltip.SetToolTip(buttonSupervisorReportEditClean, "Limpiar Registros");
            buttonSupervisorReportEditExcelTooltip.SetToolTip(buttonSupervisorReportEditExcel, "Exportar Registros a Excel");

            dataGridViewSupervisorReportAdd.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridViewSupervisorReportAdd, true, null);
            dataGridViewSupervisorReportEdit.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridViewSupervisorReportEdit, true, null);

        }

        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        var parms = base.CreateParams;
        //        parms.Style &= ~0x02000000;  // Turn off WS_CLIPCHILDREN
        //        return parms;
        //    }
        //}

        private List<ModelAddSupervisorReport> GetdataGridViewSupervisorReportAdd()
        {
            List<ModelAddSupervisorReport> list = new List<ModelAddSupervisorReport>();
            try
            {
                list.Add(new ModelAddSupervisorReport
                {
                    suprep_sup_description = _getSupervisorOutput.data.Where(x => x.sup_id == _userLoginOutput.data[0].usr_manage_id).Select(d => d.sup_description).FirstOrDefault(),
                    suprep_date = DateTime.Now.Date,
                    suprep_sto_id = _getStoresOutput.data.Where(x => x.sto_sup_id == _userLoginOutput.data[0].usr_manage_id).Select(x => x.sto_id).FirstOrDefault(),
                    suprep_in_dollar = 0,
                    suprep_out_dollar = 0,
                    suprep_in_euro = 0,
                    suprep_out_euro = 0,
                    suprep_in_peso = 0,
                    suprep_out_peso = 0,
                    suprep_comments = string.Empty,
                });

                return list;
            }
            catch (Exception ex)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Excepcion: " + ex.Message.ToLower() + " .Funcion: GetdataGridViewSupervisorReportAdd()"));
                return new List<ModelAddSupervisorReport>();
            }
        }

        private void LoadDataGridViewAddSupervisorReport()
        {
            var reports = GetdataGridViewSupervisorReportAdd();

            var comboboxStore = (DataGridViewComboBoxColumn)dataGridViewSupervisorReportAdd.Columns["dgvAddSupervisorReportColumnStore"];
            comboboxStore.DisplayMember = "sto_code";
            comboboxStore.ValueMember = "sto_id";
            comboboxStore.DataSource = _getStoresOutput.data.Where(x => x.sto_sup_id == _userLoginOutput.data[0].usr_manage_id || x.sto_id == (int)MagickInfo.STORE.ZERO).ToList();

            dataGridViewSupervisorReportAdd.DataSource = reports;
        }


        private void dataGridViewSupervisorReportAdd_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewSupervisorReportAdd.IsCurrentCellDirty)
            {
                dataGridViewSupervisorReportAdd.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridViewSupervisorReportAdd_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            ;
        }

        private async void buttonbuttonSupervisorReportAdd_Click(object sender, EventArgs e)
        {

            if (dataGridViewSupervisorReportAdd.DataSource == null || ((List<ModelAddSupervisorReport>)dataGridViewSupervisorReportAdd.DataSource).Count <= 0)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Data vacia"));
                return;
            }

            buttonbuttonSupervisorReportAdd.Enabled = false;

            List<ModelAddSupervisorReport> item = (List<ModelAddSupervisorReport>)dataGridViewSupervisorReportAdd.DataSource;

            var resultInsertReport = await _networkingOperations.InsertSupervisorReport(new InsertSupervisorReportInput
                {
                    userLoginInput = _userLoginInput,
                    insertSupervisorReport = new InsertSupervisorReport
                    {
                        suprep_sup_id = this._userLoginOutput.data[0].usr_manage_id,
                        suprep_date = item[0].suprep_date,
                        suprep_sto_id = item[0].suprep_sto_id,
                        suprep_in_dollar = item[0].suprep_in_dollar,
                        suprep_out_dollar = item[0].suprep_out_dollar,
                        suprep_in_euro = item[0].suprep_in_euro,
                        suprep_out_euro = item[0].suprep_out_euro,
                        suprep_in_peso = item[0].suprep_in_peso,
                        suprep_out_peso = item[0].suprep_out_peso,
                        suprep_comments = item[0].suprep_comments,
                        suprep_starep_id = (int)MagickInfo.STATUS_REPORT.NO_PROCESADO,
                        suprep_audit_id = this._userLoginOutput.data[0].usr_id,
                        suprep_audit_date = DateTime.Now,
                        suprep_audit_delete = false
                    }
                });

            if (resultInsertReport.Item1 && resultInsertReport.Item2.statusOperation == true)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, resultInsertReport.Item3));
                LoadDataGridViewAddSupervisorReport();
            }
            else
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, resultInsertReport.Item3));
            }

            buttonbuttonSupervisorReportAdd.Enabled = true;
        }

        private void buttonbuttonSupervisorReportAddClean_Click(object sender, EventArgs e)
        {
            LoadDataGridViewAddSupervisorReport();
        }

        private async void buttonSupervisorReportEditConsult_Click(object sender, EventArgs e)
        {
            buttonSupervisorReportEditConsult.Enabled = false; 
            buttonSupervisorReportEdit.Enabled = false;
            buttonSupervisorReportDelete.Enabled = false;
            buttonSupervisorReportEditClean.Enabled = false;
            buttonSupervisorReportEditExcel.Enabled = false;
            dateTimePickerSupervisorReportEdit.Enabled = false;

            var result = await _networkingOperations.GetSupervisorReport(new GetSupervisorReportInput
            {
                userLoginInput = _userLoginInput,
                getSupervisorReport = new GetSupervisorReport
                {
                    suprep_sup_id = _userLoginOutput.data[0].usr_manage_id,
                    suprep_date = dateTimePickerSupervisorReportEdit.Value
                }
            });

            if (result.Item1)
            {
                LoadEditSupervisorReport(result.Item2);
                CalculateLastSquareSupervisor();
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item3));
            }
            else
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item3));
            }

            buttonSupervisorReportEditConsult.Enabled = true;
            buttonSupervisorReportEdit.Enabled = true;
            buttonSupervisorReportDelete.Enabled = true;
            buttonSupervisorReportEditClean.Enabled = true;
            buttonSupervisorReportEditExcel.Enabled = true;
            dateTimePickerSupervisorReportEdit.Enabled = true;
        }

        private void LoadEditSupervisorReport(GetSupervisorReportsOutput output)
        {
            var reports = GetEditSupervisorReport(output);

            var comboboxStatus = (DataGridViewComboBoxColumn)dataGridViewSupervisorReportEdit.Columns["dgvEditSupervisorReportColumnStore"];
            comboboxStatus.DisplayMember = "sto_code";
            comboboxStatus.ValueMember = "sto_id";
            comboboxStatus.DataSource = _getStoresOutput.data.Where(x => x.sto_sup_id == _userLoginOutput.data[0].usr_manage_id || x.sto_id == (int)MagickInfo.STORE.ZERO).ToList();

            dataGridViewSupervisorReportEdit.DataSource = reports;
        }

        private List<ModelEditSupervisorReport> GetEditSupervisorReport(GetSupervisorReportsOutput output)
        {
            try
            {
                var list = new List<ModelEditSupervisorReport>();

                foreach (var items in output.data)
                {
                    list.Add(new ModelEditSupervisorReport
                    {
                        suprep_selection = false,
                        suprep_id = items.suprep_id,
                        suprep_date = items.suprep_date,
                        suprep_sup_description = this._getSupervisorOutput.data.Where(s => s.sup_id == items.suprep_sup_id).Select(d => d.sup_description).FirstOrDefault(),
                        suprep_sto_id = items.suprep_sto_id,
                        suprep_in_dollar = items.suprep_in_dollar,
                        suprep_out_dollar = items.suprep_out_dollar,
                        suprep_in_euro = items.suprep_in_euro,
                        suprep_out_euro = items.suprep_out_euro,
                        suprep_in_peso = items.suprep_in_peso,
                        suprep_out_peso = items.suprep_out_peso,
                        suprep_comments = items.suprep_comments,
                        starep_description = _getStatusReportOutput.data.Where(x => x.starep_id == items.suprep_starep_id).Select(d => d.starep_description).FirstOrDefault()
                    });
                }

                return list;
            }
            catch (Exception ex)
            {
                return new List<ModelEditSupervisorReport>();
            }
        }

        private async void buttonSupervisorReportEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewSupervisorReportEdit.DataSource == null || dataGridViewSupervisorReportEdit.Rows.Count <= 0)
            {
                return;
            }

            buttonSupervisorReportEdit.Enabled = false;
            buttonSupervisorReportDelete.Enabled = false;
            buttonSupervisorReportEditClean.Enabled = false;
            buttonSupervisorReportEditExcel.Enabled = false;
            dateTimePickerSupervisorReportEdit.Enabled = false;

            foreach (var item in (List<ModelEditSupervisorReport>)dataGridViewSupervisorReportEdit.DataSource)
            {
                if (item.suprep_selection == true)
                {
                    var resultUpdateReport = await _networkingOperations.UpdateSupervisorReport(new UpdateSupervisorReportInput
                    {
                        userLoginInput = _userLoginInput,
                        updateSupervisorReport = new UpdateSupervisorReport
                        {
                            suprep_id = item.suprep_id,
                            suprep_sup_id = this._getSupervisorOutput.data.Where(s => s.sup_description == item.suprep_sup_description).Select(d => d.sup_id).FirstOrDefault(),
                            suprep_date = item.suprep_date,
                            suprep_sto_id = item.suprep_sto_id,
                            suprep_in_dollar = item.suprep_in_dollar,
                            suprep_out_dollar = item.suprep_out_dollar,
                            suprep_in_euro = item.suprep_in_euro,
                            suprep_out_euro = item.suprep_out_euro,
                            suprep_in_peso = item.suprep_in_peso,
                            suprep_out_peso = item.suprep_out_peso,
                            suprep_comments = item.suprep_comments,
                            suprep_starep_id = (int)MagickInfo.STATUS_REPORT.NO_PROCESADO,
                            suprep_audit_id = this._userLoginOutput.data[0].usr_id,
                            suprep_audit_date = DateTime.Now,
                            suprep_audit_delete = false
                        }
                    });

                    if (resultUpdateReport.Item1)
                    {
                        _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, resultUpdateReport.Item3));
                    }
                    else
                    {
                        _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, resultUpdateReport.Item3));
                    }
                }
            }

            buttonSupervisorReportEditConsult_Click(null, null);

            buttonSupervisorReportEdit.Enabled = true;
            buttonSupervisorReportDelete.Enabled = true;
            buttonSupervisorReportEditClean.Enabled = true;
            buttonSupervisorReportEditExcel.Enabled = true;
            dateTimePickerSupervisorReportEdit.Enabled = true;
        }

        private async void buttonSupervisorReportDelete_Click(object sender, EventArgs e)
        {
            if (dataGridViewSupervisorReportEdit.DataSource == null || dataGridViewSupervisorReportEdit.Rows.Count <= 0)
            {
                return;
            }

            if (MessageBox.Show("Por favor confirme antes de proceder" + "\n" + "¿Esta seguro de eliminar los registros seleccionados?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            buttonSupervisorReportEdit.Enabled = false;
            buttonSupervisorReportDelete.Enabled = false;
            buttonSupervisorReportEditClean.Enabled = false;
            buttonSupervisorReportEditExcel.Enabled = false;
            dateTimePickerSupervisorReportEdit.Enabled = false;

            foreach (var item in (List<ModelEditSupervisorReport>)dataGridViewSupervisorReportEdit.DataSource)
            {
                if (item.suprep_selection == true)
                {
                    var resultDeleteReport = await _networkingOperations.DeleteSupervisorReport(new DeleteSupervisorReportInput
                    {
                        userLoginInput = _userLoginInput,
                        deleteSupervisorReport = new DeleteSupervisorReport
                        {
                            suprep_id = item.suprep_id,
                            suprep_audit_id = _userLoginOutput.data[0].usr_id,
                            suprep_audit_date = DateTime.Now,
                            suprep_audit_delete = true
                        }
                    });

                    if (resultDeleteReport.Item1)
                    {
                        _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, resultDeleteReport.Item3));
                    }
                    else
                    {
                        _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, resultDeleteReport.Item3));
                    }
                }
            }

            buttonSupervisorReportEditConsult_Click(null, null);

            buttonSupervisorReportEdit.Enabled = true;
            buttonSupervisorReportDelete.Enabled = true;
            buttonSupervisorReportEditClean.Enabled = true;
            buttonSupervisorReportEditExcel.Enabled = true;
            dateTimePickerSupervisorReportEdit.Enabled = true;
        }

        private void buttonSupervisorReportEditClean_Click(object sender, EventArgs e)
        {
            if (dataGridViewSupervisorReportEdit.DataSource != null && dataGridViewSupervisorReportEdit.Rows.Count > 0)
            {
                dataGridViewSupervisorReportEdit.DataSource = new List<ModelEditSupervisorReport>();
                dataGridViewSupervisorReportEdit.Refresh();
                CalculateLastSquareSupervisor();
            }
        }

        private void CalculateLastSquareSupervisor()
        {
            if (dataGridViewSupervisorReportEdit.DataSource == null || dataGridViewSupervisorReportEdit.Rows.Count <= 0)
            {
                labelSupervisorInDolares.Text = "0";
                labelSupervisorOutDolares.Text = "0";
                labelSupervisorDifferenceDolar.Text = "0";

                labelSupervisorInEuros.Text = "0";
                labelSupervisorOutEuros.Text = "0";
                labelSupervisorDifferenceEuros.Text = "0";

                labelSupervisorInPesos.Text = "0";
                labelSupervisorOutPesos.Text = "0";
                labelSupervisorDifferencePesos.Text = "0";

                return;
            }

            decimal inDolar = 0;
            decimal outDolar = 0;
            decimal differenceDolar = 0;

            decimal inEuro = 0;
            decimal outEuro = 0;
            decimal differenceEuro = 0;

            decimal inPeso = 0;
            decimal outPeso = 0;
            decimal differencePeso = 0;

            foreach (var item in (List<ModelEditSupervisorReport>)dataGridViewSupervisorReportEdit.DataSource)
            {
                inDolar += item.suprep_in_dollar;
                outDolar += item.suprep_out_dollar;

                inEuro += item.suprep_in_euro;
                outEuro += item.suprep_out_euro;

                inPeso += item.suprep_in_peso;
                outPeso += item.suprep_out_peso;
            }

            differenceDolar = inDolar - outDolar;
            differenceEuro = inEuro - outEuro;
            differencePeso = inPeso - outPeso;

            labelSupervisorInDolares.Text = inDolar.ToString("N2");
            labelSupervisorOutDolares.Text = outDolar.ToString("N2");
            labelSupervisorDifferenceDolar.Text = differenceDolar.ToString("N2");

            labelSupervisorInEuros.Text = inEuro.ToString("N2");
            labelSupervisorOutEuros.Text = outEuro.ToString();
            labelSupervisorDifferenceEuros.Text = differenceEuro.ToString("N2");

            labelSupervisorInPesos.Text = inPeso.ToString("N2");
            labelSupervisorOutPesos.Text = outPeso.ToString("N2");
            labelSupervisorDifferencePesos.Text = differencePeso.ToString("N2");
        }

        private async void buttonSupervisorReportEditExcel_Click(object sender, EventArgs e)
        {
            if (dataGridViewSupervisorReportEdit.DataSource == null || dataGridViewSupervisorReportEdit.Rows.Count <= 0)
            {
                return;
            }

            buttonSupervisorReportEdit.Enabled = false;
            buttonSupervisorReportDelete.Enabled = false;
            buttonSupervisorReportEditClean.Enabled = false;
            buttonSupervisorReportEditExcel.Enabled = false;
            dateTimePickerSupervisorReportEdit.Enabled = false;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFileDialog.Title = "Save XLSX File";
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.DefaultExt = "XLSX";
            saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var listExcelDTO = ((List<ModelEditSupervisorReport>)dataGridViewSupervisorReportEdit.DataSource).Select(x => new ExcelReportSupervisorDTO
                {
                    suprep_sup_description = x.suprep_sup_description,
                    suprep_date = x.suprep_date,
                    suprep_sto_code = _getStoresOutput.data.Where(d => d.sto_id == x.suprep_sto_id).Select(h => h.sto_code).FirstOrDefault(),
                    suprep_in_dollar = x.suprep_in_dollar,
                    suprep_out_dollar = x.suprep_out_dollar,
                    suprep_in_euro = x.suprep_in_euro,
                    suprep_out_euro = x.suprep_out_euro,
                    suprep_in_peso = x.suprep_in_peso,
                    suprep_out_peso = x.suprep_out_peso,
                    suprep_comments = x.suprep_comments,
                    suprep_starep_description = x.starep_description
                }).ToList();

                var result = await ManageExcel.CreateReportSupervisor(listExcelDTO, GetTotalesExcel(listExcelDTO), saveFileDialog.FileName);

                if (result.Item1)
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item2));
                }
                else
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item2));
                }
            }

            buttonSupervisorReportEdit.Enabled = true;
            buttonSupervisorReportDelete.Enabled = true;
            buttonSupervisorReportEditClean.Enabled = true;
            buttonSupervisorReportEditExcel.Enabled = true;
            dateTimePickerSupervisorReportEdit.Enabled = true;
        }

        private ExcelReportSupervisorTotalDTO GetTotalesExcel(List<ExcelReportSupervisorDTO> listExcelDTO)
        {
            ExcelReportSupervisorTotalDTO itemTotal = new ExcelReportSupervisorTotalDTO();
            foreach (var item in listExcelDTO)
            {
                itemTotal.total_suprep_in_dollar += item.suprep_in_dollar;
                itemTotal.total_suprep_out_dollar += item.suprep_out_dollar;
                itemTotal.total_suprep_in_euro+= item.suprep_in_euro;
                itemTotal.total_suprep_out_euro += item.suprep_out_euro;
                itemTotal.total_suprep_in_peso += item.suprep_in_peso;
                itemTotal.total_suprep_out_peso += item.suprep_out_peso;
            }

            return itemTotal;
        }

        private bool flagCommentSize = false;
        private void dataGridViewSupervisorReportEdit_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 11 && flagCommentSize == false)
            {
                dataGridViewSupervisorReportEdit.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                flagCommentSize = true;
            }
            else if (e.ColumnIndex == 11)
            {
                dataGridViewSupervisorReportEdit.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                dataGridViewSupervisorReportEdit.Columns[11].Width = 200;
                flagCommentSize = false;
            }
        }
    }
}
