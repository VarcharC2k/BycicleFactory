using DevExpress.DashboardCommon;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraSpreadsheet.Import.OpenXml;
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
using VTMES3_RE.Models;

namespace VTMES3_RE.View.CamstarInf
{
    public partial class frmDynamicComponentIssue : DevExpress.XtraEditors.XtraForm
    {
        // 템플릿 파일 경로 
        string folderName = Application.StartupPath + @"\Templates\ComponentIssue";
        string fileName = "ComponentIssue.xlsx";
        // 현재 엑셀시트가 제출된 상태인지 확인 변수
        bool IsSubmit = false;
        bool isSubmit = false;
        clsWork work = new clsWork();
        DataView collectionView = null;
        DataTable MasterdataTable = new DataTable("DynamicComponentIssue");
        string[] prostr = null;
        string[] instr = null;
        string result = string.Empty;
        public frmDynamicComponentIssue()
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
        private void DisplayData(DataView Dv,int crd)
        {
            MasterdataTable = Dv.Table;
            object[] arrobj = MasterdataTable.Select().Select(x => x["Materials"]).ToArray();
            string[] arrstr = arrobj.Cast<string>().ToArray();

            object[] inobj = MasterdataTable.Select().Select(x => x["InsertQty"]).ToArray();
            instr = arrobj.Cast<string>().ToArray();

            object[] proobj = MasterdataTable.Select().Select(x => x["ProductName"]).ToArray();
            prostr = arrobj.Cast<string>().ToArray();


            //excelSheetControl.CreateNewDocument();
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            int j = 2;
            //excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(dataTable, 0, 0, externalDSOptions);
            
            for(int k =1; k < excelSheetControl.Document.Worksheets[0].GetDataRange().RowCount; k++)
            {
                for (int i = 0; i < arrstr.Length; i++)
                {
                    excelSheetControl.Document.Worksheets[0].Cells[crd, j].SetValue(arrstr[i].ToString());
                    
                    j = j + 3;
                }
                j = 2;
            }

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

                //List<string> arrList = new List<string>();
                //// 컨테이너 속성 시작 인덱스
                //int startAttrIdx = dataTable.Columns.IndexOf("Container");
                //// 컨테이너 속성 종료 인덱스
                //int endAttrIdx = dataTable.Columns.IndexOf("Attribute Result");

                //// arrList 에 속성 컬럼 설정
                //if (startAttrIdx > -1 || endAttrIdx > -1)
                //{
                //    startAttrIdx = startAttrIdx + 1;
                //    endAttrIdx = endAttrIdx - 1;

                //    for (int i = startAttrIdx; i < endAttrIdx + 1; i++)
                //    {
                //        arrList.Add(dataTable.Columns[i].ColumnName);
                //    }
                //}

                // 시트 테이블에 대한 컨테이너 속성 변경 Api 호출

                string InsertContainer = string.Empty;
                successCnt = WrGlobal.Camster_Common.ComponentIssue(dataTable, InsertContainer,MasterdataTable,prostr,instr);

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
                        //result = dataTable.Rows[0][j * 3 + 3].ToString();
                        switch (dataTable.Columns[j].ColumnName)
                        {
                            case "Result":
                            worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "Result2":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "Result3":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "Result4":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "Result5":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
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

        private void cmdSearch_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            
            if (isSubmit)
            {
                MessageBox.Show("재조회시 닫기 버튼을 누른 후 다시 조회 버튼을 눌러 주시길 바랍니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataView co = null;
            for(int i=1; i < excelSheetControl.Document.Worksheets[0].GetDataRange().RowCount; i++)
            {
                co = work.GetComponentIssueSelectDef(excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue);
                if(co.Table.Rows.Count <= 0)
                {
                    excelSheetControl.Document.Worksheets[0].Cells[i, 1].SetValue("해당 " + excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).TextValue + " 컨테이너에 SAP코드가 BOM을 가지고 있지 않습니다 다시한번 확인해주세요.");
                }
                int crd = excelSheetControl.Document.Worksheets[0].Cells[1, i].ColumnIndex;
                DisplayData(co, crd);
            }

            isSubmit = true;


        }
    }
}