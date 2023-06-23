using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors;
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

namespace VTMES3_RE.View.PopUp
{
    // 이순희 PM 작성폼
    public partial class frmOrderDispatch_Popup : DevExpress.XtraEditors.XtraForm
    {
        string MfgOrder;
        string folderName = Application.StartupPath + @"\Templates\ProductionManagement";
        string fileName = "MfgOrderStart.xlsx";
        string pre_name = "";
        bool IsSubmit = false;

        public frmOrderDispatch_Popup()
        {
            InitializeComponent();
        }

        public frmOrderDispatch_Popup(String _MfgOrder)
        {
            InitializeComponent();
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }
            MfgOrder = _MfgOrder;
        }

        private void frmOrderDispatch_Popup_Load(object sender, EventArgs e)
        {
            this.v_RY_START_REASONTableAdapter.Fill(this.iFRYDataSet.V_RY_START_REASON);
            this.v_RY_ORDER_DISPATCHTableAdapter.FillByMfgOrderName(this.iFRYDataSet.V_RY_ORDER_DISPATCH, MfgOrder);

            IWorkbook workbook = excelSheetControl.Document;
            workbook.LoadDocument(folderName + "\\" + fileName);
            Worksheet worksheet = workbook.Worksheets[0];

            excel_qty_sum();

            // Start Reason 초기값 설정
            cbo_Start_Reason.ItemIndex = 1;
        }

        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            //if (IsSubmit)
            //{
            //    MessageBox.Show("제출 처리된 양식은 다시 제출할 수 없습니다.\n초기화 후 작업지시를 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}

            if (WrGlobal.Camster_Common.IsExecuting)
            {
                MessageBox.Show("현재 Camstar Interface 기능이 실행 중 입니다.\n잠시 후 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            String msg = "";
            if (txt_1st.Text.Trim() == "")
            {
                msg = "앞첨자가 빈값 입니다.\n앞첨자 입력후 제출을 눌어 주세요.";
                MessageBox.Show(msg);
                txt_1st.Focus();
                return;
            }

            WrGlobal.Camster_Common.IsExecuting = true;
            IsSubmit = true;

            try
            {
                this.Validate();

                Worksheet worksheet = excelSheetControl.Document.Worksheets[0];

                CellRange range = worksheet.GetDataRange();
                DataTable dt = excelSheetControl.Document.Worksheets[0].CreateDataTable(range, true);
                dt.TableName = "SaveTable";

                DataTableExporter exporter = excelSheetControl.Document.Worksheets[0].CreateDataTableExporter(range, dt, true);
                exporter.Options.ConvertEmptyCells = true;
                exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                //DataRow[] newRows = dt.Select("CONTAINER IS NOT NULL AND RESULT IS NOT '성공!' ");
                DataRow[] newRows = dt.Select("CONTAINER <> '' AND RESULT <> '성공!' ");

                if (dt.Rows.Count < 1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출할 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                foreach (DataRow dr in dt.Rows)
                {
                    if ((dr["Container"] ?? "").ToString() == "")
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show("Container (작업지시번호)는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                DataColumn Factory = new DataColumn("Factory", typeof(string));
                DataColumn Level = new DataColumn("Level", typeof(string));
                DataColumn Owner = new DataColumn("Owner", typeof(string));
                DataColumn StartReason = new DataColumn("StartReason", typeof(string));
                DataColumn MfgOrder = new DataColumn("MfgOrder", typeof(string));

                DataColumn Product = new DataColumn("Product", typeof(string));
                DataColumn Product_Rev = new DataColumn("Product Rev", typeof(string));
                DataColumn Workflow = new DataColumn("Workflow", typeof(string));
                DataColumn Workflow_Rev = new DataColumn("Workflow Rev", typeof(string));
                DataColumn PriorityCode = new DataColumn("PriorityCode", typeof(string));

                DataColumn UOM = new DataColumn("UOM", typeof(string));
                DataColumn Comments = new DataColumn("Comments", typeof(string));

                Factory.DefaultValue = "Rayence";
                Level.DefaultValue = "Unit";
                Owner.DefaultValue = "Vatech";
                //StartReason.DefaultValue =  "NORMAL";
                StartReason.DefaultValue = cbo_Start_Reason.Text;
                MfgOrder.DefaultValue = mfgOrderNameTextEdit.Text;

                Product.DefaultValue = productNameTextEdit.Text;
                Product_Rev.DefaultValue = "";
                Workflow.DefaultValue = "";
                Workflow_Rev.DefaultValue = "";
                PriorityCode.DefaultValue = "Medium";

                UOM.DefaultValue = "ea";
                Comments.DefaultValue = txt_container_comment.Text;

                dt.Columns.Add(Factory);
                dt.Columns.Add(Level);
                dt.Columns.Add(Owner);
                dt.Columns.Add(StartReason);
                dt.Columns.Add(MfgOrder);

                dt.Columns.Add(Product);
                dt.Columns.Add(Product_Rev);
                dt.Columns.Add(Workflow);
                dt.Columns.Add(Workflow_Rev);
                dt.Columns.Add(PriorityCode);

                dt.Columns.Add(UOM);
                dt.Columns.Add(Comments);
                dt.Columns.Add("BoolResult", typeof(System.Boolean));

                for (int i = 0; i < dt.Rows.Count; i++)
                {

                }

                lblMemo.Text += "전체 : " + dt.Rows.Count.ToString() + "건 | ";

                int successCnt = 0;

                successCnt = WrGlobal.Camster_Common.ContainerStart(dt);

                if (successCnt == -1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    return;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        switch (dt.Columns[j].ColumnName)
                        {
                            case "Container":
                            case "Result":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dt.Rows[i][j]);
                                break;
                            case "BoolResult":
                                if (!Convert.ToBoolean(dt.Rows[i][j]))
                                {
                                    for (int k = 0; k < range.ColumnCount; k++)
                                    {
                                        range[i + 1, k].FillColor = Color.Red;
                                    }
                                }
                                break;
                        }
                    }
                }

                lblMemo.Text += "성공 : " + successCnt.ToString() + "건 | ";
                lblMemo.Text += "실패 : " + (dt.Rows.Count - successCnt).ToString() + "건 | ";

                //MessageBox.Show("작업지시를 제출 했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            WrGlobal.Camster_Common.IsExecuting = false;
        }

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

        private void excelSheetControl_CellValueChanged(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs e)
        {
            // qty 인경우
            if (e.ColumnIndex == 1)
            {
                excel_qty_sum();
            }

            // Container 인경우
            if (e.ColumnIndex == 0)
            {
                //first_name_delete();
                //first_name_add();
            }
        }


        private void excel_qty_sum()
        {
            int sum = 0;
            try
            {
                // 엑셀의 Row 만큼 Loop 실행
                for (int i = 1; i < excelSheetControl.Document.Worksheets[0].GetDataRange().RowCount; i++)
                {
                    sum += Convert.ToInt32(excelSheetControl.Document.Worksheets[0].GetCellValue(1, i).ToString());
                }
            }
            catch
            {

            }
            txt_job_qty.EditValue = sum.ToString();
        }

        private void first_name_add()
        {
            // 앞첨자 추가 하는 로직
            try
            {
                // 엑셀의 Row 만큼 Loop 실행
                for (int i = 1; i < excelSheetControl.Document.Worksheets[0].GetDataRange().RowCount; i++)
                {
                    if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).IsText == true)
                    {
                        // 앞첨자가 @@ 이면 Skip
                        if (txt_1st.Text.Trim() != "@@")
                        {
                            // 앞첨자 길이가 > 0 이상이면
                            if (txt_1st.Text.Length > 0)
                            {
                                // Container 문자열 길이 >= 앞첨자 문자열 길이 인경우
                                if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue.Length >= txt_1st.Text.Length)
                                {
                                    // Container 에 앞첨자가 없으면
                                    if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue.Substring(0, txt_1st.Text.Length) != txt_1st.Text)
                                    {
                                        // Container 가 Auto 가 아니면
                                        if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue != "Auto")
                                        {
                                            // Container 에 앞첨자 추가
                                            excelSheetControl.Document.Worksheets[0].Cells[i, 0].SetValue(txt_1st.Text + excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue);
                                        }
                                    }
                                }
                                else
                                {
                                    // Container 가 Auto 가 아니면
                                    if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue != "Auto")
                                    {
                                        // Container 에 앞첨자 추가
                                        excelSheetControl.Document.Worksheets[0].Cells[i, 0].SetValue(txt_1st.Text + excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue);
                                    }
                                }
                            }
                        }
                    }
                    else if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).IsNumeric == true)
                    {
                        string Con = Convert.ToString(excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).NumericValue);
                        // 앞첨자가 @@ 이면 Skip
                        if (txt_1st.Text.Trim() != "@@")
                        {
                            // 앞첨자 길이가 > 0 이상이면
                            if (txt_1st.Text.Length > 0)
                            {
                                // Container 문자열 길이 >= 앞첨자 문자열 길이 인경우
                                if (Con.Length >= txt_1st.Text.Length)
                                {
                                    // Container 에 앞첨자가 없으면
                                    if (Con.Substring(0, txt_1st.Text.Length) != txt_1st.Text)
                                    {
                                        // Container 가 Auto 가 아니면
                                        if (Con != "Auto")
                                        {
                                            // Container 에 앞첨자 추가
                                            excelSheetControl.Document.Worksheets[0].Cells[i, 0].SetValue(txt_1st.Text + Con);
                                        }
                                    }
                                }
                                else
                                {
                                    // Container 가 Auto 가 아니면
                                    if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue != "Auto")
                                    {
                                        // Container 에 앞첨자 추가
                                        excelSheetControl.Document.Worksheets[0].Cells[i, 0].SetValue(txt_1st.Text + excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {

            }
            // 이전 앞첨자 기억( 앞첨자 변경 되었을때만 기억 )
            if (pre_name != txt_1st.Text)
            {
                pre_name = txt_1st.Text;
            }
        }

        private void first_name_delete()
        {
            // 앞첨자 삭제 하는 로직
            try
            {
                // 엑셀의 Row 만큼 Loop 실행
                for (int i = 1; i < excelSheetControl.Document.Worksheets[0].GetDataRange().RowCount; i++)
                {
                    if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).IsText == true)
                    {
                        // 이전 앞첨자 길이가 > 0 이상이면
                        if (pre_name.Length > 0)
                        {
                            // Container 가 Auto 가 아니면
                            if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue != "Auto")
                            {
                                // Container 에 이전 앞첨자와 동일한 데이터가 있다면
                                if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue.Substring(0, pre_name.Length) == pre_name.ToString())
                                {
                                    // Container 에 이전 앞첨자 삭제
                                    excelSheetControl.Document.Worksheets[0].Cells[i, 0].
                                            SetValue(excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue.
                                            Substring(pre_name.Length, excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue.Length - pre_name.Length));
                                }

                            }
                        }
                        
                    }
                    else if(excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).IsNumeric == true)
                    {
                        string con = Convert.ToString(excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).NumericValue);
                        // 이전 앞첨자 길이가 > 0 이상이면
                        if (pre_name.Length > 0)
                        {
                            // Container 가 Auto 가 아니면
                            if (con != "Auto")
                            {
                                // Container 에 이전 앞첨자와 동일한 데이터가 있다면
                                if (con.Substring(0, pre_name.Length) == pre_name.ToString())
                                {
                                    // Container 에 이전 앞첨자 삭제
                                    excelSheetControl.Document.Worksheets[0].Cells[i, 0].
                                    SetValue(con.Substring(pre_name.Length, con.Length - pre_name.Length));
                                }

                            }
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void txt_1st_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                first_name_delete();
                first_name_add();
            }

            else if (e.KeyCode == Keys.Enter)
            {
                first_name_delete();
                first_name_add();
            }
        }

        private void txt_1st_Leave(object sender, EventArgs e)
        {
            first_name_delete();
            first_name_add();
        }
    }
}