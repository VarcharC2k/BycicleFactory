using DevExpress.DataAccess.Wizard.Presenters;
using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraBars.Navigation;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Export;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace VTMES3_RE.View.CamstarInf
{
    public partial class frmModelingImport : DevExpress.XtraEditors.XtraForm
    {
        // 템플릿 파일 경로
        string folderName = Application.StartupPath + @"\Templates\Modeling";
        string fileName = "UDCD - Object Data Point.xlsx";
        // 현재 엑셀시트가 제출된 상태인지 확인 변수
        bool IsSubmit = false;
        public frmModelingImport()
        {
            // Test 용 접속 정보
            //WrGlobal.Camstar_Host = "RY-MESAPP-SVR01";
            //WrGlobal.Camstar_Port = 443;
            //WrGlobal.Camstar_UserName = "Administrator";
            //WrGlobal.Camstar_Password = "Dentalimageno.1";

            InitializeComponent();
            // CamstarCommon Object 생성되지 않았으면 생성 
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }
            // 템플릿 폴더내 파일 항목 하위 메뉴에 표시
            DirectoryInfo di = new DirectoryInfo(folderName);
            int FileIdx = 0;

            foreach (System.IO.FileInfo file in di.GetFiles())
            {
                FileIdx++;

                TileNavCategory categoryItem = new TileNavCategory();
                categoryItem.Appearance.Options.UseFont = true;
                categoryItem.AppearanceHovered.Options.UseFont = true;
                categoryItem.AppearanceSelected.Options.UseFont = true;
                categoryItem.Name = "categoryItem" + FileIdx.ToString();
                categoryItem.Caption = file.Name;
                categoryItem.TileText = file.Name;
                categoryItem.ElementClick += new NavElementClickEventHandler(this.categoryItem_ElementClick);
                tileNavPane.Categories.Add(categoryItem);
            }
        }
        // 템플릿 파일 클릭시 로딩
        private void categoryItem_ElementClick(object sender, NavElementEventArgs e)
        {
            fileName = ((DevExpress.XtraBars.Navigation.TileNavElement)e.Element).TileText;
            excelSheetControl.LoadDocument(folderName + "\\" + fileName);
            IsSubmit = false;
        }
        // 시트 -> DataTable 변환시 에러 처리
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

        private void cmdExecute_ElementClick(object sender, NavElementEventArgs e)
        {

            if (fileName == "") return;

            WrGlobal.Camster_Common.IsExecuting = true;
            IsSubmit = true;

            // 첨부파일에 따른 조건 분기
            // fileName = UDCD - Object Data Point.xlsx 일때
            if (fileName.ToString() == "UDCD_Object.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";
                    // 엑셀 시트 내용 -> DataTable 로 변환
                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // 입력 값 검증
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["User Data Name"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("User Data Name 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }
                    // 결과 컬럼 추가
                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    // 시트 테이블에 대한 Api 호출
                    successCnt = WrGlobal.Camster_Common.UserDataCollectDef_ObjectDataPoint(dataTable);
                    // 성공 항목이 없으면 종료
                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }
                    // 결과 메시지 처리 및 색상 처리
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = UDCD - Resource Data Point.xlsx 일때
            else if (fileName.ToString() == "UDCD_Resource.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["User Data Name"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("User Data Name 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.UserDataCollectDef_ResourceDataPoint(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = UDCD - Value Data Point.xlsx 일때
            else if (fileName.ToString() == "UDCD_Value.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["User Data Name"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("User Data Name 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.UserDataCollectDef_ValueDataPoint(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                        lblMemo.Text = "진행 : " + i.ToString() + "건 | ";
                        lblMemo.Invalidate();
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = UDCD_Delete.xlsx 일때
            else if (fileName.ToString() == "UDCD_Delete.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["User Data Name"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("User Data Name 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.UserDataCollectDef_DataPoint_Delete(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (dataTable.Rows[i][j] != DBNull.Value)
                                    //{
                                    //    if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //    {
                                    //        for (int k = 0; k < range.ColumnCount; k++)
                                    //        {
                                    //            range[i + 1, k].FillColor = Color.Red;
                                    //        }
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    break;
                                    //}

                                    break;
                            }
                        }
                        lblMemo.Text = "진행 : " + i.ToString() + "건 | ";
                        lblMemo.Invalidate();
                        //statusStrip11.Refresh();
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = BOP.xlsx 일때
            else if (fileName.ToString() == "BOP.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["BOP Name"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("BOP Name 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.BOP_Modeling_Import(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = Start.xlsx 일때
            else if (fileName.ToString() == "Start.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["Container"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("Container 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.ContainerStart(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // fileName = Close.xlsx 일때
            else if (fileName.ToString() == "Close.xlsx")
            {
                try
                {
                    Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                    CellRange range = worksheet.GetDataRange();

                    DataTable dataTable = worksheet.CreateDataTable(range, true);
                    dataTable.TableName = "ExcelUpload";

                    DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                    exporter.CellValueConversionError += exporter_CellValueConversionError;
                    exporter.Export();

                    if (dataTable.Rows.Count < 1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        if ((dr["Container"] ?? "").ToString() == "")
                        {
                            WrGlobal.Camster_Common.IsExecuting = false;
                            MessageBox.Show("Container 는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                    int successCnt = 0;
                    successCnt = WrGlobal.Camster_Common.Close(dataTable);

                    if (successCnt == -1)
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        return;
                    }

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Container":
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    // 성능 향상을 위해 주석 처리
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    break;
                            }
                        }
                    }

                    lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                    lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            WrGlobal.Camster_Common.IsExecuting = false;
        }

        private void cmdClose_ElementClick(object sender, NavElementEventArgs e)
        {
            this.Close();
        }
    }
}