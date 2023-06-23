using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors;
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
using VTMES3_RE.CodeDataSetTableAdapters;
using VTMES3_RE.Common;

namespace VTMES3_RE.View.ProductionManagement
{
    public partial class frmCSI_BatchPlan : DevExpress.XtraEditors.XtraForm
    {
        // 템플릿 파일 설정
        string folderName = Application.StartupPath + @"\Templates\WorkTemplate";
        string fileName = "CsiBatchPlan.xlsx";

        public frmCSI_BatchPlan()
        {
            InitializeComponent();

            barEditStartDate.EditValue = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            barEditEndDate.EditValue = DateTime.Today.AddMonths(1).AddDays(DateTime.Today.Day * -1);

            //barEditStartDate.EditValue = DateTime.Today.AddDays(-1);
            //barEditEndDate.EditValue = DateTime.Today.AddMonths(1).AddDays(DateTime.Today.Day * -1);

            //barEditStartDate.EditValue = Convert.ToDateTime("2023-01-02");
            //barEditEndDate.EditValue = Convert.ToDateTime("2023-01-06");

            csI_Batch_PlanBindingSource.AllowNew = false;
        }

        // 엑셀 시트에 템플릿 파일 로드
        private void frmCSI_BatchPlan_Load(object sender, EventArgs e)
        {
            DisplayData();

            IWorkbook workbook = excelSheetControl.Document;
            workbook.LoadDocument(folderName + "\\" + fileName);
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.DataBindings.BindTableToDataSource(csI_Batch_PlanBindingSource, 0, 0);
        }
        // 검색 클릭 이벤트
        private void btnSearch_Click(object sender, EventArgs e)
        {
            DisplayData();
        }
        // 저장 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                // 현재  Cell Edit Mode 종료
                if (excelSheetControl.IsCellEditorActive)
                {
                    excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
                }
                // 입력, 수정 정보 설정
                foreach (DataRowView drv in csI_Batch_PlanBindingSource.List)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        drv["CreId"] = WrGlobal.LoginID;
                        drv["CreDt"] = DateTime.Now;
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModDt"] = DateTime.Now;
                    }
                    else if (drv.Row.RowState == DataRowState.Modified)
                    {
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModDt"] = DateTime.Now;
                    }
                }

                // 수정 내역 저장
                csI_Batch_PlanBindingSource.EndEdit();
                csI_Batch_PlanTableAdapter.Update(iFRYDataSet.CsI_Batch_Plan);

                CellRange range = excelSheetControl.Document.Worksheets[0].GetDataRange();
                DataTable excelTable = excelSheetControl.Document.Worksheets[0].CreateDataTable(range, true);
                excelTable.TableName = "SaveTable";

                DataTableExporter exporter = excelSheetControl.Document.Worksheets[0].CreateDataTableExporter(range, excelTable, true);
                exporter.Options.ConvertEmptyCells = true;
                exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                // 신규 입력 항목은 Insert 쿼리로 처리
                DataRow[] newRows = excelTable.Select("ID_KEY IS NULL OR ID_KEY = 0");

                foreach (DataRow row in newRows)
                {
                    csI_Batch_PlanTableAdapter.Insert(Convert.ToDateTime(row["증착일"]), (row["주/야"] ?? "").ToString(),
                                                                        (row["증착기"] ?? "").ToString(),
                                                                         (row["모델"] ?? "").ToString(), (row["특이사항"] ?? "").ToString(),
                                                                        WrGlobal.LoginID, DateTime.Now, WrGlobal.LoginID, DateTime.Now);
                }

                MessageBox.Show("Batch 계획이 저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
         
        // 검색 조건에 대한 목록 엑셀 시트에 표시
        private void DisplayData()
        {
            this.csI_Batch_PlanTableAdapter.FillByBatchPlanDate(this.iFRYDataSet.CsI_Batch_Plan, (DateTime)barEditStartDate.EditValue, (DateTime)barEditEndDate.EditValue);
        }

        // 테이블로 변환시 에러 처리
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

    }
}