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

namespace VTMES3_RE.View.WorkManager.ProductPlan
{
    //frmProductionPlan 으로 재작성 되어 사용안함
    public partial class frmProductionMothlyPlan : DevExpress.XtraEditors.XtraForm
    {

        //string folderName = Application.StartupPath + @"\Templates\WorkTemplate";
        //string fileName = "EmployeeWorkTime.xlsx";

        clsWork work = new clsWork();

        public frmProductionMothlyPlan()
        {
            InitializeComponent();

            ProductionMonthlyPlanBindingSource.AllowNew = false;
        }

        private void frmEmployeeWorkTime_Load(object sender, EventArgs e)
        {
            DisplayData();
        }


        private void DisplayData()
        {
            
            this.productionMonthlyPlanTableAdapter.FillByList(iFRYDataSet.ProductionMonthlyPlan,Convert.ToInt32(txtYear.EditValue),Convert.ToInt32(txtMonth.EditValue));

            excelSheetControl.CreateNewDocument();
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(ProductionMonthlyPlanBindingSource, 0, 0, externalDSOptions);
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
                    if ((dr["PLAN_YEAR"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("PLAN_YEAR : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["PLAN_MONTH"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("PLAN_MONTH : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["PLAN_REVISION"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("PLAN_REVISION : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["SAP_CODE"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("SAP_CODE : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if ((dr["PLAN_QTY"] ?? "").ToString() == "")
                    {
                        MessageBox.Show(string.Format("PLAN_QTY : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                }

                ProductionMonthlyPlanBindingSource.EndEdit();

                productionMonthlyPlanTableAdapter.Update(iFRYDataSet.ProductionMonthlyPlan);

                DataRow[] newRows = excelTable.Select("PLAN_ID IS NULL OR PLAN_ID = 0");

                foreach (DataRow dr in newRows)
                {

                    

                    productionMonthlyPlanTableAdapter.Insert((Convert.ToInt32(dr["PLAN_YEAR"] ?? "")), (Convert.ToInt32(dr["PLAN_MONTH"] ?? "")),
                                                                        (dr["PLAN_REVISION"] ?? "").ToString(), (dr["SAP_CODE"] ?? "").ToString(),
                                                                        (Convert.ToInt32(dr["PLAN_QTY"] ?? "")));
                }

                MessageBox.Show("실적이 저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.productionMonthlyPlanTableAdapter.Fill(this.iFRYDataSet.ProductionMonthlyPlan);
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
            DisplayData();
        }
    }
}