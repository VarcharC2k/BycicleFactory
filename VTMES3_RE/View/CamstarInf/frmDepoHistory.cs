using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.Common;
using VTMES3_RE.IFRYDataSetTableAdapters;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.CamstarInf
{
    //frmProductionPlan 으로 재작성 되어 사용안함
    public partial class frmDepoHistory : DevExpress.XtraEditors.XtraForm
    {

        //string folderName = Application.StartupPath + @"\Templates\WorkTemplate";
        //string fileName = "EmployeeWorkTime.xlsx";

        clsWork work = new clsWork();
        DataView collectionView = null;

        public frmDepoHistory()
        {
            InitializeComponent();

            
        }

        private void frmEmployeeWorkTime_Load(object sender, EventArgs e)
        {
            barEditStartDate.EditValue = DateTime.Today.AddDays(1 - DateTime.Today.Day);
            barEditEndDate.EditValue = DateTime.Today.AddDays(1 - DateTime.Today.Day);
            DisplayData();

        }


        private void DisplayData()
        {
            DataTable dataTable = new DataTable("IFRY.dbo.CsiAfterTaskInput");

            collectionView = work.GetDepoBatchNoHistoryDef();

            dataTable = collectionView.Table;

            excelSheetControl.CreateNewDocument();
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(dataTable, 0, 0, externalDSOptions);
        }

        private void DisplayDatas(string start , string end)
        {
            start.Replace("-", "");
            end.Replace("-", "");
            DataTable dataTable = new DataTable("IFRY.dbo.CsiAfterTaskInput");

            collectionView = work.GetDepoBatchNoHistoryDef2(start,end);

            dataTable = collectionView.Table;

            excelSheetControl.CreateNewDocument();
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(dataTable, 0, 0, externalDSOptions);
        }

            // 저장
            private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                if (excelSheetControl.IsCellEditorActive)
                {
                    excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
                }

                CellRange range = excelSheetControl.Document.Worksheets[0].GetDataRange();
                DataTable excelTable = excelSheetControl.Document.Worksheets[0].CreateDataTable(range, true);
                excelTable.TableName = "SaveTable";

                DataTableExporter exporter = excelSheetControl.Document.Worksheets[0].CreateDataTableExporter(range, excelTable, true);
                exporter.Options.ConvertEmptyCells = true;
                exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                foreach (DataRow dr in excelTable.Rows)
                {
                    if ((dr["BATCHNO"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("BATCHNO : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["CsiWeightL"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("CsiWeightL : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["CsiWeightR"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("CsiWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["CsiWeight3"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("CsiWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["CsiWeight4"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("CsiWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["TliWeight"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("TliWeight : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["TliWeight5"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("TliWeight : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["ShutterWeightL"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("ShutterWeightL : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["ShutterWeightR"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("ShutterWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["ShutterWeight3"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("ShutterWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["ShutterWeight4"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("ShutterWeightR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["SampleThickSPL1"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("SampleThickSPL1 : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["SampleThickSPL2"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("SampleThickSPL2 : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                }
                //신규 BatchNo가 있는지 확인후 없으면 Insert하는 로직
                DataRow lastrow = excelTable.Rows[excelTable.Rows.Count - 1];
                
                DataView Bn = work.GetBatchNoDef(lastrow["BatchNo"].ToString());

                if(Bn.Count == 0)
                {
                    work.GetBatchNoInsertDef2(lastrow["BatchNo"].ToString(), lastrow["CsiWeightL"].ToString(), lastrow["CsiWeightR"].ToString(), lastrow["CsiWeight3"].ToString(), lastrow["CsiWeight4"].ToString()
                      , lastrow["TliWeight"].ToString(), lastrow["TliWeight5"].ToString(), lastrow["ShutterWeightL"].ToString()
                      , lastrow["ShutterWeightR"].ToString(), lastrow["ShutterWeight3"].ToString(), lastrow["ShutterWeight4"].ToString(), lastrow["SampleThickSPL1"].ToString(), lastrow["SampleThickSPL2"].ToString());

                    MessageBox.Show("실적이 저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else
                {
                    MessageBox.Show("이미 존재하는 BatchNo입니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                DisplayData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        // 닫기
        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

        private void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            DataTableExporter exporter = sender as DataTableExporter;
            CellValueToColumnTypeConverter defaultToColumnTypeConverter = exporter != null ? exporter.Options.DefaultCellValueToColumnTypeConverter : null;
            if (e.DataColumn.DataType == typeof(Double) && e.CellValue.IsText)
            {
                object newDataTableValue = CellValue.Empty;
                ConversionResult isConverted = defaultToColumnTypeConverter.Convert(e.Cell, e.CellValue, e.DataColumn.DataType, out newDataTableValue);
                e.DataTableValue = newDataTableValue;
                e.Action = isConverted == ConversionResult.Success ? DataTableExporterAction.Continue : DataTableExporterAction.SkipRow;
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string start = barEditStartDate.Text;
            string end = barEditEndDate.Text;
            DisplayDatas(start,end);
        }

        private void cmdRepair_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Validate();

            if (excelSheetControl.IsCellEditorActive)
            {
                excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
            }

            CellRange range = excelSheetControl.Document.Worksheets[0].GetDataRange();
            DataTable excelTable = excelSheetControl.Document.Worksheets[0].CreateDataTable(range, true);
            excelTable.TableName = "SaveTable";

            DataTableExporter exporter = excelSheetControl.Document.Worksheets[0].CreateDataTableExporter(range, excelTable, true);
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

            exporter.CellValueConversionError += exporter_CellValueConversionError;
            exporter.Export();

            foreach (DataRow dr in excelTable.Rows)
            {
                work.GetBatchNoUpdateDef(dr["BatchNo"].ToString(), dr["CsiWeightL"].ToString(), dr["CsiWeightR"].ToString(), dr["CsiWeight3"].ToString(), dr["CsiWeight4"].ToString()
                    , dr["TliWeight"].ToString(), dr["TliWeight5"].ToString(), dr["ShutterWeightL"].ToString(), dr["ShutterWeightR"].ToString(), dr["ShutterWeight3"].ToString(), dr["ShutterWeight4"].ToString()
                    , dr["SampleThickSPL1"].ToString(), dr["SampleThickSPL2"].ToString());

            }
            MessageBox.Show("수정 완료되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);

            DisplayData();
        }

    }
}