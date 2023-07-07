using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsAppTrazoRegistroTienda.Excel
{
    internal class ManageExcel
    {
        public static async Task<Tuple<bool, string>> CreateReportStore(List<ExcelReportStoreDTO> listData, ExcelReportStoreTotalDTO itemTotal, string path)
        {
            return await Task.Run(() => ShellCreateReportStore(listData, itemTotal, path));
        }

        public static async Task<Tuple<bool, string>> CreateReportSupervisor(List<ExcelReportSupervisorDTO> listData, ExcelReportSupervisorTotalDTO itemTotal, string path)
        {
            return await Task.Run(() => ShellCreateReportSupervisor(listData, itemTotal, path));
        }

        private static Tuple<bool, string> ShellCreateReportStore(List<ExcelReportStoreDTO> listData, ExcelReportStoreTotalDTO itemTotal, string path)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            ExcelPackage excel = new ExcelPackage();

            // name of the sheet
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // setting the properties
            // of the work sheet 
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties
            // of the first row
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            try
            {
                workSheet.Cells[1, 1].Value = "Tienda";
                workSheet.Cells[1, 2].Value = "Fecha";
                workSheet.Cells[1, 3].Value = "Total Facturado Bs";
                workSheet.Cells[1, 4].Value = "Cambio Bs";
                workSheet.Cells[1, 5].Value = "Tasa";
                workSheet.Cells[1, 6].Value = "Equivalente $";
                workSheet.Cells[1, 7].Value = "Euros";
                workSheet.Cells[1, 8].Value = "Zelle";
                workSheet.Cells[1, 9].Value = "Dolar $";
                workSheet.Cells[1, 10].Value = "Gastos";
                workSheet.Cells[1, 11].Value = "Resta Efectivo";
                workSheet.Cells[1, 12].Value = "Estatus";
                workSheet.Cells[1, 13].Value = "Supervisor";
                workSheet.Cells[1, 14].Value = "Comentarios";
                workSheet.Cells[1, 15].Value = "Estatus Reporte";

                int col = 1;
                int row = 2;

                foreach (var item in listData)
                {
                    workSheet.Cells[row, col].Value = item.storep_code;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_date.ToString("d/MM/yyyy");
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_total_bs;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_change_bs;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_rate;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_equivalent_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_payed_euro;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_payed_zelle;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_payed_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_expended_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_total_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_sta_description;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_supervisor;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_comments;
                    col++;
                    workSheet.Cells[row, col].Value = item.storep_starep_description;
                    col = 1;
                    row++;
                }

                row++; // Linea en blanco

                workSheet.Cells[row, col].Value = "Total Facturado Bs";
                col++;
                workSheet.Cells[row, col].Value = "Cambio Bs";
                col++;
                workSheet.Cells[row, col].Value = "Equivalente $";
                col++;
                workSheet.Cells[row, col].Value = "Euros";
                col++;
                workSheet.Cells[row, col].Value = "Zelle";
                col++;
                workSheet.Cells[row, col].Value = "Dolar $";
                col++;
                workSheet.Cells[row, col].Value = "Gastos";
                col++;
                workSheet.Cells[row, col].Value = "Resta Efectivo $";
                col = 1;
                row++;

                workSheet.Cells[row, col].Value = itemTotal.total_storep_total_bs;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_change_bs;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_equivalent_dollar;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_payed_euro;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_payed_zelle;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_payed_dollar;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_expended_dollar;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_storep_total_dollar;


                for (int i = 1; i <= 15; i++)
                {
                    workSheet.Column(i).AutoFit();
                }

                if (File.Exists(path))
                    File.Delete(path);

                // Create excel file on physical disk 
                FileStream objFileStrm = File.Create(path);
                objFileStrm.Close();

                // Write content to excel file 
                File.WriteAllBytes(path, excel.GetAsByteArray());
                //Close Excel package
                excel.Dispose();

                return new Tuple<bool, string>(true, "Archivo guardado con exito.");
            }
            catch (Exception ex)
            {
                return new Tuple<bool, string>(false, "Error, Excepcion: " + ex.Message.ToLower());
            }
        }

        private static Tuple<bool, string> ShellCreateReportSupervisor(List<ExcelReportSupervisorDTO> listData, ExcelReportSupervisorTotalDTO itemTotal, string path)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            ExcelPackage excel = new ExcelPackage();

            // name of the sheet
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // setting the properties
            // of the work sheet 
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties
            // of the first row
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            try
            {
                workSheet.Cells[1, 1].Value = "Supervisor";
                workSheet.Cells[1, 2].Value = "Fecha";
                workSheet.Cells[1, 3].Value = "Tienda";
                workSheet.Cells[1, 4].Value = "Entrada Dolares";
                workSheet.Cells[1, 5].Value = "Salida Dolares";
                workSheet.Cells[1, 6].Value = "Entrada Euros";
                workSheet.Cells[1, 7].Value = "Salida Euros";
                workSheet.Cells[1, 8].Value = "Entrada Pesos";
                workSheet.Cells[1, 9].Value = "Salida Pesos";
                workSheet.Cells[1, 10].Value = "Comentarios";
                workSheet.Cells[1, 11].Value = "Estatus Reporte";

                int col = 1;
                int row = 2;

                foreach (var item in listData)
                {
                    workSheet.Cells[row, col].Value = item.suprep_sup_description;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_date.ToString("d/MM/yyyy");
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_sto_code;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_in_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_out_dollar;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_in_euro;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_out_euro;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_in_peso;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_out_peso;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_comments;
                    col++;
                    workSheet.Cells[row, col].Value = item.suprep_starep_description;
                    col = 1;
                    row++;
                }

                row++; // Linea en blanco

                workSheet.Cells[row, col].Value = "Total Mes";
                col++;
                workSheet.Cells[row, col].Value = "Entrada Dolares";
                col++;
                workSheet.Cells[row, col].Value = "Salida Dolares";
                col++;
                workSheet.Cells[row, col].Value = "Entradas Euros";
                col++;
                workSheet.Cells[row, col].Value = "Salida Euros";
                col++;
                workSheet.Cells[row, col].Value = "Entrada Pesos";
                col++;
                workSheet.Cells[row, col].Value = "Salida Pesos";
                col = 1;
                row++;

                workSheet.Cells[row, col].Value = "Totales";
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_dollar;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_out_dollar;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_euro;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_out_euro;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_peso;
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_out_peso;
                col = 1;
                row++;

                workSheet.Cells[row, col].Value = "Diferencias";
                col++;
                workSheet.Cells[row, col].Value = "DOLARES=";
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_dollar - itemTotal.total_suprep_out_dollar;
                col++;
                workSheet.Cells[row, col].Value = "EUROS=";
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_euro - itemTotal.total_suprep_out_euro;
                col++;
                workSheet.Cells[row, col].Value = "PESOS=";
                col++;
                workSheet.Cells[row, col].Value = itemTotal.total_suprep_in_peso - itemTotal.total_suprep_out_peso;

                for (int i = 1; i <= 15; i++)
                {
                    workSheet.Column(i).AutoFit();
                }

                if (File.Exists(path))
                    File.Delete(path);

                // Create excel file on physical disk 
                FileStream objFileStrm = File.Create(path);
                objFileStrm.Close();

                // Write content to excel file 
                File.WriteAllBytes(path, excel.GetAsByteArray());
                //Close Excel package
                excel.Dispose();

                return new Tuple<bool, string>(true, "Archivo guardado con exito.");
            }
            catch (Exception ex)
            {
                return new Tuple<bool, string>(false, "Error, Excepcion: " + ex.Message.ToLower());
            }
        }
    }
}
