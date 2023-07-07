using ClassLibraryNetworking.Models.Input;
using ClassLibraryNetworking.Models.MagickNumbers;
using ClassLibraryNetworking.Models.Output;
using ClassLibraryNetworking.Operations;
using MysqlClassLibrary;
using System.Reflection;
using System.Windows.Forms;
using WinFormsAppTrazoRegistroTienda.Excel;

namespace WinFormsAppTrazoRegistroTienda
{
    public partial class RegisterStoreUserControl : UserControl
    {
        private MainForm _mainForm;
        private NetworkingOperations _networkingOperations;
        private UserLoginInput _userLoginInput;
        private UserLoginOutput _userLoginOutput;
        private GetStoresOutput _getStoresOutput;
        private GetSupervisorOutput _getSupervisorOutput;
        private GetStatusOutput _getStatusOutput;
        private GetStatusReportOutput _getStatusReportOutput;
        private CheckValuesMYSQL mysql;

        private ToolTip buttonStoreReportAddonsultTooltip = new ToolTip();
        private ToolTip buttonbuttonStoreReportAddTooltip = new ToolTip();
        private ToolTip buttonbuttonStoreReportAddCleanTooltip = new ToolTip();
        private ToolTip buttonStoreReportEditConsultTooltip = new ToolTip();
        private ToolTip buttonStoreReportEditTooltip = new ToolTip();
        private ToolTip buttonStoreReportEditCleanTooltip = new ToolTip();
        private ToolTip buttonStoreReportEditExcelTooltip = new ToolTip();

        public RegisterStoreUserControl(MainForm mainForm, NetworkingOperations networkingOperations, UserLoginInput userLoginInput, UserLoginOutput userLoginOutput,
            GetStoresOutput getStoresOutput, GetSupervisorOutput getSupervisorOutput, GetStatusOutput getStatusOutput, GetStatusReportOutput getStatusReportOutput,
            int TIMEOUT_MYSQL)
        {
            InitializeComponent();

            _mainForm = mainForm;
            _networkingOperations = networkingOperations;
            _userLoginInput = userLoginInput;
            _userLoginOutput = userLoginOutput;
            _getStoresOutput = getStoresOutput;
            _getSupervisorOutput = getSupervisorOutput;
            _getStatusOutput = getStatusOutput;
            _getStatusReportOutput = getStatusReportOutput;
            mysql = new(TIMEOUT_MYSQL);

            buttonStoreReportAddonsultTooltip.SetToolTip(buttonStoreReportEditConsult, "Consultar Registros");
            buttonbuttonStoreReportAddTooltip.SetToolTip(buttonStoreReportAdd, "Agregar Reporte");
            buttonbuttonStoreReportAddCleanTooltip.SetToolTip(buttonStoreReportAddClean, "Limpiar Campos");
            buttonStoreReportEditConsultTooltip.SetToolTip(buttonStoreReportEditConsult, "Consultar Registros");
            buttonStoreReportEditTooltip.SetToolTip(buttonStoreReportEdit, "Editar Registros");
            buttonStoreReportEditCleanTooltip.SetToolTip(buttonStoreReportEditClean, "Limpiar Registros");
            buttonStoreReportEditExcelTooltip.SetToolTip(buttonStoreReportEditExcel, "Exportar Registros a Excel");

            dataGridViewStoreReportAdd.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridViewStoreReportAdd, true, null);
            dataGridViewStoreReportEdit.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dataGridViewStoreReportEdit, true, null);
           
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

        #region StoreReportAdd

        private async void buttonStoreReportGetSale_Click(object sender, EventArgs e)
        {
            buttonStoreReportGetSale.Enabled = false;
            dateTimePickerStoreReportAdd.Enabled = false;
            buttonStoreReportAdd.Enabled = false;
            buttonStoreReportAddClean.Enabled = false;

            Tuple<bool, GetStoreReportExitsOutput, string> result = await _networkingOperations.GetStoreReportExits(
            new GetStoreReportInput
            {
                getStoreReport = new GetStoreReport
                {
                    storep_sto_id = _userLoginOutput.data[0].usr_manage_id,
                    storep_date = dateTimePickerStoreReportAdd.Value.Date
                },
                userLoginInput = _userLoginInput
            });

            if (result.Item1 && result.Item2.statusOperation == true)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item2.message));

                if (result.Item2.data.exits == false)
                {
                    this.dataGridViewStoreReportAdd.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedFirstInput);
                    LoadDataGridFirstTime();
                    DataGridViewStoreReportAddReadOnly(true);
                }
                else if (result.Item2.data.exits == true && result.Item2.data.isProcessed == false)
                {
                    this.dataGridViewStoreReportAdd.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedNoFirstInput);
                    LoadDataGridNoFirstTime();
                    DataGridViewStoreReportAddReadOnly(false);
                }
                else if (result.Item2.data.isProcessed == true)
                {
                    LoadEmptyDataGridViewStoreReportAdd();
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Error fecha ya procesada no se puede agregar registros."));
                }
            }
            else
            {
                LoadEmptyDataGridViewStoreReportAdd();
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item3));
            }

            dateTimePickerStoreReportAdd.Enabled = true;
            buttonStoreReportAdd.Enabled = true;
            buttonStoreReportAddClean.Enabled = true;
            buttonStoreReportGetSale.Enabled = true;
        }

        private async void LoadDataGridFirstTime()
        {
            Tuple<bool, OutputTotalBsRate, string> resultMysql = await mysql.GetRegistersOfDay(dateTimePickerStoreReportAdd.Value.Date);

            if (resultMysql.Item1)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, resultMysql.Item3));
                LoadFirstDataGridViewStoreReportAdd(resultMysql.Item2);
            }
            else
            {
                LoadEmptyDataGridViewStoreReportAdd();
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, resultMysql.Item3));
            }
        }

        private void LoadFirstDataGridViewStoreReportAdd(OutputTotalBsRate output)
        {
            var reports = GetStoreReportFirstTime(output);

            var comboboxStatus = (DataGridViewComboBoxColumn)dataGridViewStoreReportAdd.Columns["dgvAddStoreReportColumnStatus"];
            comboboxStatus.DisplayMember = "sta_description";
            comboboxStatus.ValueMember = "sta_id";
            comboboxStatus.DataSource = _getStatusOutput.data.ToList();

            dataGridViewStoreReportAdd.DataSource = reports;

            NoWriteAjusmentOMonth();
        }

        void NoWriteAjusmentOMonth()
        {
            if(((List<ModelAddStoreReport>)(dataGridViewStoreReportAdd.DataSource)).Count() > 0)
            {
                int rowIndex = 0;
                foreach(var item in ((List<ModelAddStoreReport>)(dataGridViewStoreReportAdd.DataSource)))
                {
                    if(item.storep_starep_id == (int)MagickInfo.STATUS_REPORT.PROCESADO)
                    {
                        dataGridViewStoreReportAdd.Rows[rowIndex].ReadOnly = true;
                    }
                    rowIndex++;
                }
            }
        }

        private void LoadEmptyDataGridViewStoreReportAdd()
        {
            var reports = new List<ModelAddStoreReport>();

            var comboboxStatus = (DataGridViewComboBoxColumn)dataGridViewStoreReportAdd.Columns["dgvAddStoreReportColumnStatus"];
            comboboxStatus.DisplayMember = "sta_description";
            comboboxStatus.ValueMember = "sta_id";
            comboboxStatus.DataSource = _getStatusOutput.data.ToList();

            dataGridViewStoreReportAdd.DataSource = reports;
        }

        private List<ModelAddStoreReport> GetStoreReportFirstTime(OutputTotalBsRate output)
        {
            try
            {
                var list = new List<ModelAddStoreReport>();

                foreach (var items in output.ListValues)
                {
                    decimal equivalent = decimal.Round((items.TotalBs / items.Rate), 2, MidpointRounding.AwayFromZero);
                    if (items.ProcessedStatus == (int)MagickInfo.STATUS_REPORT.PROCESADO)
                    {
                        list.Add(new ModelAddStoreReport
                        {
                            storep_sto_code = _getStoresOutput.data.Where(s => s.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(s => s.sto_code).FirstOrDefault(),
                            storep_date = dateTimePickerStoreReportAdd.Value.Date,
                            storep_total_bs = items.TotalBs,
                            storep_change_bs = 0,
                            storep_rate = items.Rate,
                            storep_equivalent_dollar = equivalent,
                            storep_payed_euro = 0,
                            storep_payed_zelle = 0,
                            storep_payed_dollar = equivalent,
                            storep_expended_dollar = 0,
                            storep_total_dollar = equivalent,
                            storep_sta_id = _getStatusOutput.data.Where(x => x.sta_id == (int)MagickInfo.STATUS.AJUSTE_CON_EL_CIERRE).Select(x => x.sta_id).FirstOrDefault(),
                            storep_supervisor = _getSupervisorOutput.data.Where(s => s.sup_id == (_getStoresOutput.data.Where(e => e.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(f => f.sto_sup_id).FirstOrDefault())).Select(g => g.sup_description).FirstOrDefault(),
                            storep_comments = "Ajuste con el cierre",
                            storep_starep_id = items.ProcessedStatus
                        });
                    }
                    else
                    {
                        list.Add(new ModelAddStoreReport
                        {
                            storep_sto_code = _getStoresOutput.data.Where(s => s.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(s => s.sto_code).FirstOrDefault(),
                            storep_date = dateTimePickerStoreReportAdd.Value.Date,
                            storep_total_bs = items.TotalBs,
                            storep_change_bs = 0,
                            storep_rate = items.Rate,
                            storep_equivalent_dollar = equivalent,
                            storep_payed_euro = 0,
                            storep_payed_zelle = 0,
                            storep_payed_dollar = equivalent,
                            storep_expended_dollar = 0,
                            storep_total_dollar = equivalent,
                            storep_sta_id = _getStatusOutput.data.Where(x => x.sta_id == (int)MagickInfo.STATUS.S_E).Select(x => x.sta_id).FirstOrDefault(),
                            storep_supervisor = _getSupervisorOutput.data.Where(s => s.sup_id == (_getStoresOutput.data.Where(e => e.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(f => f.sto_sup_id).FirstOrDefault())).Select(g => g.sup_description).FirstOrDefault(),
                            storep_comments = string.Empty,
                            storep_starep_id = items.ProcessedStatus
                        });
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Excepcion: " + ex.Message.ToLower()));
                return new List<ModelAddStoreReport>();
            }
        }

        private void DataGridViewStoreReportAddReadOnly(bool isFirstTime)
        {
            if (isFirstTime == true)
            {
                dataGridViewStoreReportAdd.Columns[0].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[1].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[2].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[3].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[4].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[5].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[6].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[7].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[8].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[9].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[10].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[11].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[12].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[13].ReadOnly = false;
            }
            else
            {
                dataGridViewStoreReportAdd.Columns[0].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[1].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[2].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[3].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[4].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[5].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[6].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[7].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[8].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[9].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[10].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[11].ReadOnly = false;
                dataGridViewStoreReportAdd.Columns[12].ReadOnly = true;
                dataGridViewStoreReportAdd.Columns[13].ReadOnly = false;
            }
        }

        private void LoadDataGridNoFirstTime()
        {
            LoadNoFirstDataGridViewStoreReportAdd();
        }

        private void LoadNoFirstDataGridViewStoreReportAdd()
        {
            var reports = GetStoreReportNoFirstTime();

            var comboboxStatus = (DataGridViewComboBoxColumn)dataGridViewStoreReportAdd.Columns["dgvAddStoreReportColumnStatus"];
            comboboxStatus.DisplayMember = "sta_description";
            comboboxStatus.ValueMember = "sta_id";
            comboboxStatus.DataSource = _getStatusOutput.data.ToList();

            dataGridViewStoreReportAdd.DataSource = reports;
        }

        private List<ModelAddStoreReport> GetStoreReportNoFirstTime()
        {
            try
            {
                var list = new List<ModelAddStoreReport>();

                list.Add(new ModelAddStoreReport
                {
                    storep_sto_code = _getStoresOutput.data.Where(s => s.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(s => s.sto_code).FirstOrDefault(),
                    storep_date = dateTimePickerStoreReportAdd.Value.Date,
                    storep_total_bs = 0,
                    storep_change_bs = 0,
                    storep_rate = 1,
                    storep_equivalent_dollar = 0,
                    storep_payed_euro = 0,
                    storep_payed_zelle = 0,
                    storep_payed_dollar = 0,
                    storep_expended_dollar = 0,
                    storep_total_dollar = 0,
                    storep_sta_id = 1,
                    storep_supervisor = _getSupervisorOutput.data.Where(s => s.sup_id == (_getStoresOutput.data.Where(e => e.sto_id == _userLoginOutput.data[0].usr_manage_id).Select(f => f.sto_sup_id).FirstOrDefault())).Select(g => g.sup_description).FirstOrDefault(),
                    storep_comments = string.Empty,
                    storep_starep_id = (int)MagickInfo.STATUS_REPORT.NO_PROCESADO,
                });

                return list;
            }
            catch (Exception ex)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Excepcion: " + ex.Message.ToLower()));
                return new List<ModelAddStoreReport>();
            }
        }

        private async void buttonbuttonStoreReportAdd_Click(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportAdd.DataSource == null || ((List<ModelAddStoreReport>)dataGridViewStoreReportAdd.DataSource).Count <= 0)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, "Data Invalida"));
                return;
            }

            dateTimePickerStoreReportAdd.Enabled = false;
            buttonStoreReportAdd.Enabled = false;
            buttonStoreReportAddClean.Enabled = false;

            var getCloseDay = await mysql.GetCloseDay(dateTimePickerStoreReportAdd.Value);

            if (getCloseDay.Item1 == false)
            {
                dateTimePickerStoreReportAdd.Enabled = true;
                buttonStoreReportAdd.Enabled = true;
                buttonStoreReportAddClean.Enabled = true;
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, getCloseDay.Item2));
                return; 
            }

            foreach (var item in (List<ModelAddStoreReport>)dataGridViewStoreReportAdd.DataSource)
            {
                var result = await _networkingOperations.InsertStoreReport(
                new InsertStoreReportInput
                {
                    userLoginInput = _userLoginInput,
                    insertStoreReport = new InsertStoreReport
                    {
                        storep_sto_id = _getStoresOutput.data.Where(x => x.sto_code == item.storep_sto_code).Select(f => f.sto_id).FirstOrDefault(),
                        storep_date = item.storep_date,
                        storep_total_bs = item.storep_total_bs,
                        storep_change_bs = item.storep_change_bs,
                        storep_rate = item.storep_rate,
                        storep_equivalent_dollar = item.storep_equivalent_dollar,
                        storep_payed_euro = item.storep_payed_euro,
                        storep_payed_zelle = item.storep_payed_zelle,
                        storep_payed_dollar = item.storep_payed_dollar,
                        storep_expended_dollar = item.storep_expended_dollar,
                        storep_total_dollar = item.storep_total_dollar,
                        storep_sta_id = item.storep_sta_id,
                        storep_sup_id = _getSupervisorOutput.data.Where(s => s.sup_description == item.storep_supervisor).Select(g => g.sup_id).FirstOrDefault(),
                        storep_comments = item.storep_comments,
                        storep_starep_id = item.storep_starep_id,
                        storep_audit_id = _userLoginOutput.data[0].usr_id,
                        storep_audit_date = DateTime.Now,
                        storep_audit_delete = false
                    }
                });

                if (result.Item1)
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item3));
                }
                else
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item3));
                }
            }

            LoadEmptyDataGridViewStoreReportAdd();
            dateTimePickerStoreReportAdd.Enabled = true;
            buttonStoreReportAdd.Enabled = true;
            buttonStoreReportAddClean.Enabled = true;
        }

        private void buttonbuttonStoreReportAddClean_Click(object sender, EventArgs e)
        {
            LoadEmptyDataGridViewStoreReportAdd();
        }

        private void dataGridViewStoreReportAdd_CellValueChangedFirstInput(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 6)//Euros
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); //euros
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[7].Value.ToString()); //euros
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString()); //Equivalente
                decimal dolar = equivalent - euros - zelle; //equivalent dollars - euros - zelle
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[9].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if (e.ColumnIndex == 7)//Zelle
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[6].Value.ToString()); //euros
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); //euros
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString()); //Equivalente
                decimal dolar = equivalent - euros - zelle; //equivalent dollars - euros
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[9].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if (e.ColumnIndex == 9)//Gastos
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[6].Value.ToString());
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[7].Value.ToString());
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString());
                decimal dolar = equivalent - euros - zelle;
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if(e.ColumnIndex == 11)
            {
                int comboboxId = (int)dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if(comboboxId == (int)MagickInfo.STATUS.AJUSTE_CON_EL_CIERRE || comboboxId == (int)MagickInfo.STATUS.CIERRE_DE_MES)
                {
                    this.dataGridViewStoreReportAdd.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedFirstInput);
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (int)MagickInfo.STATUS.S_E;
                    dataGridViewStoreReportAdd.RefreshEdit();
                    this.dataGridViewStoreReportAdd.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedFirstInput);
                }
            }
        }

        private void dataGridViewStoreReportAdd_CellValueChangedNoFirstInput(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3 || e.ColumnIndex == 4)//Cambio Bs || tasa
            {
                decimal cambio = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[3].Value.ToString()); //cambio bs
                decimal tasa = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[4].Value.ToString()); //tasa

                try
                {
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value = decimal.Round((cambio / tasa), 2, MidpointRounding.AwayFromZero); // equivalent dolar
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = decimal.Round((cambio / tasa), 2, MidpointRounding.AwayFromZero); // dolar
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = decimal.Round((cambio / tasa), 2, MidpointRounding.AwayFromZero); // resta Efectivo
                }
                catch
                {
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value = (decimal)0; // equivalent dolar
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = (decimal)0; // dolar
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = (decimal)0; // resta efectivo
                }
            }

            if (e.ColumnIndex == 6)//Euros
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); //euros
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[7].Value.ToString()); //euros
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString()); //Equivalente
                decimal dolar = equivalent - euros - zelle; //equivalent dollars - euros - zelle
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[9].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if (e.ColumnIndex == 7)//Zelle
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[6].Value.ToString()); //euros
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); //euros
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString()); //Equivalente
                decimal dolar = equivalent - euros - zelle; //equivalent dollars - euros
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[9].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if (e.ColumnIndex == 9)//Gastos
            {
                decimal euros = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[6].Value.ToString());
                decimal zelle = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[7].Value.ToString());
                decimal equivalent = Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[5].Value.ToString());
                decimal dolar = equivalent - euros - zelle;
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[8].Value = dolar;

                decimal restaEfectivo = euros + dolar - Decimal.Parse(dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[10].Value = restaEfectivo;
            }
            else if (e.ColumnIndex == 11)
            {
                int comboboxId = (int)dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (comboboxId == (int)MagickInfo.STATUS.AJUSTE_CON_EL_CIERRE || comboboxId == (int)MagickInfo.STATUS.CIERRE_DE_MES)
                {
                    this.dataGridViewStoreReportAdd.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedFirstInput);
                    dataGridViewStoreReportAdd.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (int)MagickInfo.STATUS.S_E;
                    dataGridViewStoreReportAdd.RefreshEdit();
                    this.dataGridViewStoreReportAdd.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportAdd_CellValueChangedFirstInput);
                }
            }
        }

        private void dataGridViewStoreReportAdd_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportAdd.IsCurrentCellDirty)
            {
                dataGridViewStoreReportAdd.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridViewStoreReportAdd_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            ;
        }

        #endregion StoreReportAdd

        #region StoreReportEdit
        private async void buttonStoreReportEditConsult_Click(object sender, EventArgs e)
        {
            buttonStoreReportEditConsult.Enabled = false;
            buttonStoreReportEdit.Enabled = false;
            buttonStoreReportEditClean.Enabled = false;
            buttonStoreReportEditExcel.Enabled = false;
            dateTimePickerStoreReportEdit.Enabled = false;

            var result = await _networkingOperations.GetStoreReport(new GetStoreReportInput
            {
                userLoginInput = _userLoginInput,
                getStoreReport = new GetStoreReport
                {
                    storep_sto_id = _userLoginOutput.data[0].usr_manage_id,
                    storep_date = dateTimePickerStoreReportEdit.Value
                }
            });

            if (result.Item1)
            {
                LoadEditStoreReport(result.Item2);
                DataGridViewEditStoreReportReadOnly();
                CalculateLastSquareStore();
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item3));
            }
            else
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item3));
            }

            buttonStoreReportEditConsult.Enabled = true;
            buttonStoreReportEdit.Enabled = true;
            buttonStoreReportEditClean.Enabled = true;
            buttonStoreReportEditExcel.Enabled = true;
            dateTimePickerStoreReportEdit.Enabled = true;
        }

        private void LoadEditStoreReport(GetStoreReportsOutput output)
        {
            var reports = GetEditStoreReport(output);

            var comboboxStatus = (DataGridViewComboBoxColumn)dataGridViewStoreReportEdit.Columns["dgvEditStoreReportColumnStatus"];
            comboboxStatus.DisplayMember = "sta_description";
            comboboxStatus.ValueMember = "sta_id";
            comboboxStatus.DataSource = _getStatusOutput.data.ToList();

            dataGridViewStoreReportEdit.DataSource = reports;
        }

        private List<ModelEditStoreReport> GetEditStoreReport(GetStoreReportsOutput output)
        {
            try
            {
                var list = new List<ModelEditStoreReport>();

                foreach (var items in output.data)
                {
                    list.Add(new ModelEditStoreReport
                    {
                        storep_selection = false,
                        storep_id = items.storep_id,
                        storep_code = items.sto_code,
                        storep_date = items.storep_date.Date,
                        storep_total_bs = items.storep_total_bs,
                        storep_change_bs = items.storep_change_bs,
                        storep_rate = items.storep_rate,
                        storep_equivalent_dollar = items.storep_equivalent_dollar,
                        storep_payed_euro = items.storep_payed_euro,
                        storep_payed_zelle = items.storep_payed_zelle,
                        storep_payed_dollar = items.storep_payed_dollar,
                        storep_expended_dollar = items.storep_expended_dollar,
                        storep_total_dollar = items.storep_total_dollar,
                        storep_sta_id = items.storep_sta_id,
                        storep_supervisor = this._getSupervisorOutput.data.Where(s => s.sup_id == items.storep_sup_id).Select(g => g.sup_description).FirstOrDefault(),
                        storep_comments = items.storep_comments,
                        storep_starep_description = items.starep_description
                    });
                }

                return list;
            }
            catch (Exception ex)
            {
                _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, "Excepcion: " + ex.Message.ToLower() + " .Funcion: GetEditStoreReport()"));
                return new List<ModelEditStoreReport>();
            }
        }

        private void DataGridViewEditStoreReportReadOnly()
        {
            foreach (DataGridViewRow row in dataGridViewStoreReportEdit.Rows)
            {
                row.Cells[1].ReadOnly = true;
                row.Cells[2].ReadOnly = true;
                row.Cells[3].ReadOnly = true;
                row.Cells[4].ReadOnly = true;

                if (Decimal.Parse(row.Cells[4].Value.ToString()) == 0) //Total Facturado es cero
                {
                    row.Cells[5].ReadOnly = false;
                    row.Cells[6].ReadOnly = false;
                    row.Cells[7].ReadOnly = true;
                    row.Cells[8].ReadOnly = false;
                    row.Cells[9].ReadOnly = false;
                    row.Cells[10].ReadOnly = true;
                    row.Cells[11].ReadOnly = false;
                    row.Cells[12].ReadOnly = true;
                    row.Cells[13].ReadOnly = false;
                    row.Cells[14].ReadOnly = true;
                    row.Cells[15].ReadOnly = false;
                    row.Cells[16].ReadOnly = true;
                }
                else
                {
                    row.Cells[5].ReadOnly = true;
                    row.Cells[6].ReadOnly = true;
                    row.Cells[7].ReadOnly = true;
                    row.Cells[8].ReadOnly = false;
                    row.Cells[9].ReadOnly = false;
                    row.Cells[10].ReadOnly = true;
                    row.Cells[11].ReadOnly = false;
                    row.Cells[12].ReadOnly = true;
                    row.Cells[13].ReadOnly = false;
                    row.Cells[14].ReadOnly = true;
                    row.Cells[15].ReadOnly = false;
                    row.Cells[16].ReadOnly = true;
                }

                if (row.Cells[16].Value.ToString() == "PROCESADO")
                {
                    row.ReadOnly = true;
                }
            }
        }

        private void CalculateLastSquareStore()
        {
            if (dataGridViewStoreReportEdit.DataSource == null || dataGridViewStoreReportEdit.Rows.Count <= 0)
            {
                labelStoreTotalFacturado.Text = "0";
                labelStoreCambioBs.Text = "0";
                labelStoreEquivalente.Text = "0";
                labelStoreEuros.Text = "0";
                labelStoreZelle.Text = "0";
                labelStoreDolares.Text = "0";
                labelStoreGastos.Text = "0";
                labelStoreRestaEfectivo.Text = "0";

                return;
            }

            decimal totalFacturado = 0;
            decimal cambioBs = 0;
            decimal equivalente = 0;
            decimal euros = 0;
            decimal zelle = 0;
            decimal dolares = 0;
            decimal gastos = 0;
            decimal restaEfectivo = 0;

            foreach (var item in (List<ModelEditStoreReport>)dataGridViewStoreReportEdit.DataSource)
            {
                totalFacturado += item.storep_total_bs;
                cambioBs += item.storep_change_bs;
                equivalente += item.storep_equivalent_dollar;
                euros += item.storep_payed_euro;
                zelle += item.storep_payed_zelle;
                dolares += item.storep_payed_dollar;
                gastos += item.storep_expended_dollar;
                restaEfectivo += item.storep_total_dollar;
            }

            labelStoreTotalFacturado.Text = totalFacturado.ToString("N2");
            labelStoreCambioBs.Text = cambioBs.ToString("N2");
            labelStoreEquivalente.Text = equivalente.ToString("N2");
            labelStoreEuros.Text = euros.ToString("N2");
            labelStoreZelle.Text = zelle.ToString("N2");
            labelStoreDolares.Text = dolares.ToString("N2");
            labelStoreGastos.Text = gastos.ToString("N2");
            labelStoreRestaEfectivo.Text = restaEfectivo.ToString("N2");
        }

        private void dataGridViewStoreReportEdit_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 6 || e.ColumnIndex == 8 || e.ColumnIndex == 9 || e.ColumnIndex == 11) && dataGridViewStoreReportEdit.DataSource != null)
            {
                decimal totalFacturado = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[4].Value.ToString());
                decimal cambioBs = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[5].Value.ToString());
                decimal tasa = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[6].Value.ToString());
                decimal euros = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[8].Value.ToString());
                decimal zelle = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[9].Value.ToString());
                decimal gastos = Decimal.Parse(dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[11].Value.ToString());
                decimal equivalente = 0;
                decimal dolares = 0;

                if (totalFacturado != 0 && tasa != 0)
                {
                    equivalente = totalFacturado / tasa;
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[7].Value = decimal.Round(equivalente, 2, MidpointRounding.AwayFromZero); //Equivalente Bs
                }

                if (cambioBs != 0 && tasa != 0)
                {
                    equivalente = cambioBs / tasa;
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[7].Value = decimal.Round(equivalente, 2, MidpointRounding.AwayFromZero); //Equivalente Bs
                }

                try
                {
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[10].Value = dolares = decimal.Round(equivalente - euros - zelle, 2, MidpointRounding.AwayFromZero); //Dolar
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[12].Value = decimal.Round(euros + dolares - gastos, 2, MidpointRounding.AwayFromZero); //Resta Efectivo
                }
                catch
                {
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[10].Value = decimal.Round(0 - euros, 2, MidpointRounding.AwayFromZero); //Dolar
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[12].Value = decimal.Round(0 - gastos, 2, MidpointRounding.AwayFromZero); //Resta Efectivo
                }

                dataGridViewStoreReportEdit.Refresh();
            }
            else if (e.ColumnIndex == 13 && dataGridViewStoreReportEdit.DataSource != null)
            {
                int comboboxId = (int)dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (comboboxId == (int)MagickInfo.STATUS.AJUSTE_CON_EL_CIERRE || comboboxId == (int)MagickInfo.STATUS.CIERRE_DE_MES)
                {
                    this.dataGridViewStoreReportEdit.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportEdit_CellValueChanged);
                    dataGridViewStoreReportEdit.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (int)MagickInfo.STATUS.S_E;
                    dataGridViewStoreReportEdit.RefreshEdit();
                    this.dataGridViewStoreReportEdit.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewStoreReportEdit_CellValueChanged);
                }
            }
            
        }

        private void dataGridViewStoreReportEdit_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            ;
        }

        private void dataGridViewStoreReportEdit_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportEdit.IsCurrentCellDirty)
            {
                dataGridViewStoreReportEdit.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private async void buttonStoreReportEdit_Click(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportEdit.DataSource == null || dataGridViewStoreReportEdit.Rows.Count <= 0)
            {
                return;
            }

            buttonStoreReportEdit.Enabled = false;
            buttonStoreReportEditClean.Enabled = false;
            buttonStoreReportEditExcel.Enabled = false;
            dateTimePickerStoreReportEdit.Enabled = false;

            foreach (var item in (List<ModelEditStoreReport>)dataGridViewStoreReportEdit.DataSource)
            {
                if (item.storep_selection == true)
                {
                    var resultUpdateReport = await _networkingOperations.UpdateStoreReport(new UpdateStoreReportInput
                    {
                        userLoginInput = _userLoginInput,
                        updateStoreReport = new UpdateStoreReport
                        {
                            storep_id = item.storep_id,
                            storep_sto_id = this._getStoresOutput.data.Where(s => s.sto_code == item.storep_code).Select(d => d.sto_id).FirstOrDefault(),
                            storep_date = item.storep_date,
                            storep_total_bs = item.storep_total_bs,
                            storep_change_bs = item.storep_change_bs,
                            storep_rate = item.storep_rate,
                            storep_equivalent_dollar = item.storep_equivalent_dollar,
                            storep_payed_euro = item.storep_payed_euro,
                            storep_payed_zelle = item.storep_payed_zelle,
                            storep_payed_dollar = item.storep_payed_dollar,
                            storep_expended_dollar = item.storep_expended_dollar,
                            storep_total_dollar = item.storep_total_dollar,
                            storep_sta_id = item.storep_sta_id,
                            storep_sup_id = this._getSupervisorOutput.data.Where(s => s.sup_description == item.storep_supervisor).Select(g => g.sup_id).FirstOrDefault(),
                            storep_comments = item.storep_comments,
                            storep_starep_id = _getStatusReportOutput.data.Where(x => x.starep_description == item.storep_starep_description).Select(g => g.starep_id).FirstOrDefault(),
                            storep_audit_id = this._userLoginOutput.data[0].usr_id,
                            storep_audit_date = DateTime.Now,
                            storep_audit_delete = false
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

            buttonStoreReportEditConsult_Click(null, null);

            buttonStoreReportEdit.Enabled = true;
            buttonStoreReportEditClean.Enabled = true;
            buttonStoreReportEditExcel.Enabled = true;
            dateTimePickerStoreReportEdit.Enabled = true;
        }

        private void StoreReportEditClean_Click(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportEdit.DataSource != null && dataGridViewStoreReportEdit.Rows.Count > 0)
            {
                dataGridViewStoreReportEdit.DataSource = new List<ModelEditStoreReport>();
                dataGridViewStoreReportEdit.Refresh();
                CalculateLastSquareStore();
            }
        }

        private async void buttonStoreReportEditExcel_Click(object sender, EventArgs e)
        {
            if (dataGridViewStoreReportEdit.DataSource == null || dataGridViewStoreReportEdit.Rows.Count <= 0)
            {
                return;
            }

            buttonStoreReportEdit.Enabled = false;
            buttonStoreReportEditClean.Enabled = false;
            buttonStoreReportEditExcel.Enabled = false;
            dateTimePickerStoreReportEdit.Enabled = false;

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
                var listExcelDTO = ((List<ModelEditStoreReport>)dataGridViewStoreReportEdit.DataSource).Select(x => new ExcelReportStoreDTO
                {
                    storep_code = x.storep_code,
                    storep_date = x.storep_date,
                    storep_total_bs = x.storep_total_bs,
                    storep_change_bs = x.storep_change_bs,
                    storep_rate = x.storep_rate,
                    storep_equivalent_dollar = x.storep_equivalent_dollar,
                    storep_payed_euro = x.storep_payed_euro,
                    storep_payed_zelle = x.storep_payed_zelle,
                    storep_payed_dollar = x.storep_payed_dollar,
                    storep_expended_dollar = x.storep_expended_dollar,
                    storep_total_dollar = x.storep_total_dollar,
                    storep_sta_description = _getStatusOutput.data.Where(s => s.sta_id == x.storep_sta_id).Select(d => d.sta_description).FirstOrDefault(),
                    storep_supervisor = x.storep_supervisor,
                    storep_comments = x.storep_comments,
                    storep_starep_description = x.storep_starep_description
                }).ToList();

                var result = await ManageExcel.CreateReportStore(listExcelDTO, GetTotalesExcel(listExcelDTO), saveFileDialog.FileName);

                if (result.Item1)
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(true, result.Item2));
                }
                else
                {
                    _mainForm.RaiseRichTextInsertNewMessage?.Invoke(this, new Tuple<bool, string>(false, result.Item2));
                }
            }

            buttonStoreReportEdit.Enabled = true;
            buttonStoreReportEditClean.Enabled = true;
            buttonStoreReportEditExcel.Enabled = true;
            dateTimePickerStoreReportEdit.Enabled = true;
        }

        private ExcelReportStoreTotalDTO GetTotalesExcel(List<ExcelReportStoreDTO> listExcelDTO)
        {
            ExcelReportStoreTotalDTO itemTotal = new ExcelReportStoreTotalDTO();
            foreach (var item in listExcelDTO)
            {
                itemTotal.total_storep_total_bs += item.storep_total_bs;
                itemTotal.total_storep_change_bs += item.storep_change_bs;
                itemTotal.total_storep_equivalent_dollar += item.storep_equivalent_dollar;
                itemTotal.total_storep_payed_euro += item.storep_payed_euro;
                itemTotal.total_storep_payed_zelle += item.storep_payed_zelle;
                itemTotal.total_storep_payed_dollar += item.storep_payed_dollar;
                itemTotal.total_storep_expended_dollar += item.storep_expended_dollar;
                itemTotal.total_storep_total_dollar += item.storep_total_dollar;
            }

            return itemTotal;
        }

        private bool flagCommentSize = false;
        private void dataGridViewStoreReportEdit_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 15 && flagCommentSize == false)
            {
                dataGridViewStoreReportEdit.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                flagCommentSize = true;
            }
            else if (e.ColumnIndex == 15)
            {
                dataGridViewStoreReportEdit.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                dataGridViewStoreReportEdit.Columns[15].Width = 200;
                flagCommentSize = false;
            }
        }

        #endregion StoreReportEdit
    }
}
