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
using VTMES3_RE.Common;

namespace VTMES3_RE.View.CamstarInf
{
    public partial class frmCMOS_ContainerAttr : DevExpress.XtraEditors.XtraForm
    {
        // 템플릿 파일 경로 
        string folderName = Application.StartupPath + @"\Templates\ContainerAttr";
        string fileName = "CMOS_ContainerAttr.xlsx";
        // 현재 엑셀시트가 제출된 상태인지 확인 변수
        bool IsSubmit = false;

        public frmCMOS_ContainerAttr()
        {
            InitializeComponent();

            // CamstarCommon Object 생성되지 않았으면 생성 
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }
        }

        private void frmCMOS_ContainerAttr_Load(object sender, EventArgs e)
        {
            // 엑셀 시트컨트롤 로딩 파일 설정
            excelSheetControl.LoadDocument(folderName + "\\" + fileName);
            IsSubmit = false;
        }
        // 제출 클릭
        private void cmdExecute_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (fileName == "") return;

            // 제출된 시트 확인
            if (IsSubmit)
            {
                MessageBox.Show("제출 처리된 양식은 다시 제출할 수 없습니다.\n다시 파일을 선택하여 초기화 후 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            // Camstar Interface 기능이 진행중인지 확인
            if (WrGlobal.Camster_Common.IsExecuting)
            {
                MessageBox.Show("현재 Camstar Interface 기능이 실행 중 입니다.\n잠시 후 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            WrGlobal.Camster_Common.IsExecuting = true;
            IsSubmit = true;

            // 현재 셀 에디트 모드 -> 종료
            if (excelSheetControl.IsCellEditorActive)
            {
                excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
            }

            try
            {
                
                Worksheet worksheet = excelSheetControl.Document.Worksheets[0]; //worksheet 인스턴스를 생성

                CellRange range = worksheet.GetDataRange(); //worksheet 인스턴스에서 데이터를 포함하고 있는 CellRange를 지정.

                DataTable dataTable = worksheet.CreateDataTable(range, true);   //DataTable을 생성
                dataTable.TableName = "ExcelUpload";                                // 테이블명을 ExcelUpload로 지정

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

                // 결과 컬럼 추가
                dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                lblMemo.Text += "전체 : " + dataTable.Rows.Count.ToString() + "건 | ";

                int successCnt = 0;

                List<string> arrList = new List<string>();
                // 컨테이너 속성 시작 인덱스
                int startAttrIdx = dataTable.Columns.IndexOf("Container");
                // 컨테이너 속성 종료 인덱스
                int endAttrIdx = dataTable.Columns.IndexOf("Attribute Result");

                // arrList 에 속성 컬럼 설정
                if (startAttrIdx > -1 || endAttrIdx > -1)
                {
                    startAttrIdx = startAttrIdx + 1;
                    endAttrIdx = endAttrIdx - 1;

                    for (int i = startAttrIdx; i < endAttrIdx + 1; i++)
                    {
                        arrList.Add(dataTable.Columns[i].ColumnName);
                    }
                }

                // 시트 테이블에 대한 컨테이너 속성 변경 Api 호출
                successCnt = WrGlobal.Camster_Common.ContainerAttribute(dataTable);

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
                            case "Attribute Result":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "BoolResult":
                                if (!Convert.ToBoolean(dataTable.Rows[i][j]))
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
                lblMemo.Text += "실패 : " + (dataTable.Rows.Count - successCnt).ToString() + "건 | ";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            WrGlobal.Camster_Common.IsExecuting = false;
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

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

        // 초기화 버튼 클릭 -> 엑셀 시트 초기화
        private void cmdInit_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            excelSheetControl.LoadDocument(folderName + "\\" + fileName);
            lblMemo.Text = "";
            IsSubmit = false;
        }
    }
}