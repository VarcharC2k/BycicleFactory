using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors;
using DevExpress.XtraLayout.Customization;
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
using VTMES3_RE.Models;

namespace VTMES3_RE.View.CamstarInf
{
    public partial class frmExecTaskByResource : DevExpress.XtraEditors.XtraForm
    {
        // clsWork 모델 생성
        clsWork work = new clsWork();
      
        DataView collectionView = null;

        string resourceName = "";
        string dataCollectionId = "";
        string dataCollectionName = "";
        bool IsSubmit = false;
        string pre_name = "";
        string JigNo = string.Empty;
        public frmExecTaskByResource()
        {
            InitializeComponent();

            // Resource 그룹 바인딩
            ResourceGroupLookUpEdit.Properties.DataSource = work.GetResourceGroup();
            ResourceGroupLookUpEdit.Properties.DisplayMember = "그룹명";
            ResourceGroupLookUpEdit.Properties.ValueMember = "그룹명";
        }

        private void frmExecTaskByResource_Load(object sender, EventArgs e)
        {

        }

        private void ResourceGroupLookUpEdit_EditValueChanged(object sender, EventArgs e)
        {
            // 리소스 그룹 변경시 설비 리스트 바인딩
            ResourceLookUpEdit.Properties.DataSource = work.GetResourceDef((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
            ResourceLookUpEdit.Properties.DisplayMember = "설비명";
            ResourceLookUpEdit.Properties.ValueMember = "설비명";
            ResourceLookUpEdit.EditValue = null;

            // 리소스 그룹 변경시 UDCD 바인딩
            //DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
            DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
            DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
            DataCollectionLookUpEdit.Properties.ValueMember = "코드";
            DataCollectionLookUpEdit.EditValue = null;
        }

        // 검색 클릭 -> 엑셀 시트에 처리할 항목 표시
        private void btnSearch_Click(object sender, EventArgs e)
        {
            if ((ResourceGroupLookUpEdit.EditValue ?? "").ToString() == "") return;
            if ((ResourceLookUpEdit.EditValue ?? "").ToString() == "") return;
            if ((DataCollectionLookUpEdit.EditValue ?? "").ToString() == "") return;

            string test = "";

            test = DataCollectionLookUpEdit.Text;



            // 선택된 UDCD 에대한 Task 목록 가져오기
            // Datastore DB 에서 Query 결과값 가져옮
            DataView taskView = work.GetAllTaskInfoByCollection((DataCollectionLookUpEdit.EditValue ?? "").ToString());
            if (taskView.Count == 0)
            {
                MessageBox.Show("선택된 DataCollection과 연결된 Task를 찾을수 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Task 바인딩
            TaskLookUpEdit.Properties.DataSource = taskView;
            TaskLookUpEdit.Properties.ValueMember = "TaskValue";
            TaskLookUpEdit.Properties.DisplayMember = "TaskName";

            TaskLookUpEdit.EditValue = taskView[0]["TaskValue"].ToString();

            excelSheetControl.CreateNewDocument();

            // 엑셀 시트에 표시할 테이블 생성
            DataTable dataTable = new DataTable("DataPoint");
            dataTable.Columns.Add("Container", typeof(System.String));
            // UDCD 에서 DataPoint 가져오기 -> collectionView
            collectionView = work.GetDataPointByCollectionResource((DataCollectionLookUpEdit.EditValue ?? "").ToString());

            Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
            CellRange range = worksheet.GetDataRange();
            string test2 = string.Empty;
            int i = 0;
            // 테이블에 DataPoint 항목 컬럼 등록
            foreach (DataRowView drv in collectionView)
            {
               test2 = drv["IsRequired"].ToString();

                dataTable.Columns.Add(drv["DataPointName"].ToString(), typeof(System.String));

                if (test2 =="True" )
                {
                    i++;
                    continue;


                }
                else if(test2 == "False")
                {
                    range[0, i+1].FillColor = Color.Red;
                    i++;
                }
            }

            // 결과값을 표시하기 위한 추가
            //작업시작인지 아닌지 판별로직 추가
            if(test.Contains("작업시작"))                
                {
                    dataTable.Columns.Add("Result", typeof(System.String));
                    dataTable.Columns.Add("Result2", typeof(System.String));
                }
            else
            {
                dataTable.Columns.Add("Comment", typeof(System.String));
                dataTable.Columns.Add("Result", typeof(System.String));
                dataTable.Columns.Add("Result2", typeof(System.String));
            }

            // 엑셀 시트에 표시할 떄 옵션 설정
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(dataTable, 0, 0, externalDSOptions);

            // DataPoint 가 NamedObjectGroupName 일때 콤보 처리
            foreach (DataRowView drv in collectionView)
            {
                if (drv["DataType"].ToString() == "5" && drv["NamedObjectGroupName"].ToString() != "")
                {
                    CellRange comboBoxRange = excelSheetControl.Document.Worksheets[0][string.Format("DataPoint[{0}]", drv["DataPointName"].ToString())];
                    excelSheetControl.Document.Worksheets[0].CustomCellInplaceEditors.Add(comboBoxRange, CustomCellInplaceEditorType.ComboBox, drv["NamedObjectGroupName"].ToString());
                }
            }

            resourceName = (ResourceLookUpEdit.EditValue ?? "").ToString();
            dataCollectionId = (DataCollectionLookUpEdit.EditValue ?? "").ToString();
            dataCollectionName = DataCollectionLookUpEdit.Text;
            IsSubmit = false;
        }

        // 작업시작 클릭
        private void btnTaskStart_Click(object sender, EventArgs e)
        {
            if (dataCollectionId == "") return;

            // CamstarCommon Object 생성되지 않았으면 생성 
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }

            if (WrGlobal.Camster_Common.IsExecuting)
            {
                MessageBox.Show("현재 Camstar Interface 기능이 실행 중 입니다.\n잠시 후 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
                CellRange range = worksheet.GetDataRange();

                DataTable dataTable = worksheet.CreateDataTable(range, true);
                dataTable.TableName = "ExcelUpload";

                DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                if (dataTable.Rows.Count < 1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출할 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 입력값 검증
                foreach (DataRow dr in dataTable.Rows)
                {
                    if (dr["Container"].ToString() == "")
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show(string.Format("Container : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                WrGlobal.Camster_Common.IsExecuting = true;
                //IsSubmit = true;

                int successCnt = 0;

                // 세션 생성
                WrGlobal.Camster_Common.CreateSession();

                string[] arrTask = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' });

                // 작업시작 클릭
                dataTable.Columns.Add("BoolResult", typeof(System.Boolean));
                successCnt = WrGlobal.Camster_Common.WorkStartResult(arrTask[0], resourceName, dataTable, false);                

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
                            case "Result":
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 세션 종료
                WrGlobal.Camster_Common.DestroySession();
            }

            WrGlobal.Camster_Common.IsExecuting = false;
        }

        //제출
        private void btnExcuteTask_Click(object sender, EventArgs e)
        {

            if (dataCollectionId == "") return;

            // CamstarCommon Object 생성되지 않았으면 생성 
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }

            if (IsSubmit)
            {
                MessageBox.Show("제출 처리된 양식은 다시 제출할 수 없습니다.\n다시 검색한 후 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (WrGlobal.Camster_Common.IsExecuting)
            {
                MessageBox.Show("현재 Camstar Interface 기능이 실행 중 입니다.\n잠시 후 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 현재 셀 에디트 모드 -> 종료
            if (excelSheetControl.IsCellEditorActive)
            {
                excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
            }

            try
            {

                Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
                CellRange range = worksheet.GetDataRange();

                DataTable dataTable = worksheet.CreateDataTable(range, true);
                dataTable.TableName = "ExcelUpload";

                DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                if (DataCollectionLookUpEdit.Text == "CSI_증착_작업시작")
                {
                    //Batch No. 로직
                    string re = "0" + ResourceLookUpEdit.Text.Substring(ResourceLookUpEdit.Text.Length - 1);
                    string Gubun = Convert.ToString(dataTable.Rows[0][1]);
                    string Day = DateTime.Now.ToString("yyyy-MM-dd").Replace("-", "");
                    if (Gubun == "주간")
                    {
                        Gubun = "D";
                    }
                    else if (Gubun == "야간")
                    {
                        Gubun = "N";
                    }

                    string BatchNo = re + "-" + Day + "-" + Gubun;

                    foreach (DataRow dr in dataTable.Rows)
                    {
                        dr["배치번호"] = BatchNo;
                    }
                    

                   DataView BN = work.GetBatchNoDef(BatchNo);


                    if (BN.Count == 0)
                    {
                        work.GetBatchNoInsertDef(BatchNo);

                    }

                    else if (BN.Count >= 1)
                    {
                         work.GetBatchNoDelteDef(BatchNo);
                         work.GetBatchNoInsertDef(BatchNo);
                    }

                }
                //지그번호 추가
                if (DataCollectionLookUpEdit.Text == "CSI_증착_지그번호")
                {
                    DataTable dt = new DataTable();
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        DataView Ba = work.GetContainerBatchNo(dr["Container"].ToString());

                        
                        dt = Ba.Table;
                        JigNo = Convert.ToString(dt.Rows[0][0]) + "-" + dr["지그번호(상)"].ToString() + "-" + dr["지그번호(하)"].ToString();

                        dr["지그배치번호"] = JigNo;
                    }
                    
                }

                if (dataTable.Rows.Count < 1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출할 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 입력값 검증
                foreach (DataRow dr in dataTable.Rows)
                {
                    if (dr["Container"].ToString() == "")
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show(string.Format("Container : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                // 입력값 검증
                foreach (DataRowView drv in collectionView)
                {
                    if (Convert.ToBoolean(drv["IsRequired"]))
                    {
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            if (dr[drv["DataPointName"].ToString()].ToString() == "")
                            {
                                WrGlobal.Camster_Common.IsExecuting = false;
                                MessageBox.Show(string.Format("{0} - '{1}' : 필수 입력 항목입니다.", dr["Container"].ToString(), drv["DataPointName"].ToString()), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }

                    if (drv["DataType"].ToString() == "5" && drv["NamedObjectGroupName"].ToString() != "")
                    {
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            if (dr[drv["DataPointName"].ToString()].ToString() != "")
                            {
                                if (drv["NamedObjectGroupName"].ToString().IndexOf(dr[drv["DataPointName"].ToString()].ToString()) < 0)
                                {
                                    WrGlobal.Camster_Common.IsExecuting = false;
                                    MessageBox.Show(string.Format("{0} - '{1}' : 정의되지 않은 입력값입니다.", dr["Container"].ToString(), drv["DataPointName"].ToString()), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }
                }

                WrGlobal.Camster_Common.IsExecuting = true;
                IsSubmit = true;

                int successCnt = 0;

                // 세션 생성
                WrGlobal.Camster_Common.CreateSession();

                string[] arrTask = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' });

                dataTable.Columns.Add("BoolResult", typeof(System.Boolean));
                successCnt = WrGlobal.Camster_Common.ExecuteTaskByUDCResult(arrTask[1], arrTask[2], dataCollectionName, dataTable, collectionView, false);

                //if (arrTask[1] == "Data1")
                //{   // 전공정 TaskName 이 Data1 으로 고정됨 -> Data1 아닐시 명칭 수정 필요
                // 작업시작 체크시 Api 호출
                //if (StartCheckEdit.Checked)
                //{
                //    dataTable.Columns.Add("StartResult");
                //    dataTable.Columns.Add("StartBoolResult", typeof(System.Boolean));
                //    successCnt = WrGlobal.Camster_Common.WorkStart(arrTask[0], resourceName, dataTable, false);
                //}

                // Task 실행 API 호출
                //dataTable.Columns.Add("TaskResult");
                    //dataTable.Columns.Add("BoolResult", typeof(System.Boolean));
                    //successCnt = WrGlobal.Camster_Common.ExecuteTaskByUDC(arrTask[1], arrTask[2], dataCollectionName, dataTable, collectionView, false);
                //}
                //else if (arrTask[1] == "Data2")
                //{   // 후공정 TaskName 이 Data2 으로 고정됨 -> Data2 아닐시 명칭 수정 필요
                //    // Task 실행 API 호출
                //    dataTable.Columns.Add("TaskResult");
                //    dataTable.Columns.Add("TaskBoolResult", typeof(System.Boolean));
                //    successCnt = WrGlobal.Camster_Common.ExecuteTaskByUDC(arrTask[0], arrTask[1], dataCollectionName, dataTable, collectionView, false);

                //    ////작업종료 체크시 Api 호출
                //    //if (EndCheckEdit.Checked)
                //    //{
                //    //    dataTable.Columns.Add("EndResult");
                //    //    dataTable.Columns.Add("EndBoolResult", typeof(System.Boolean));
                //    //    successCnt = WrGlobal.Camster_Common.WorkFinishByUDC(arrTask[0], resourceName, dataTable, false);
                //    //}

                //    //// 다음공정으로 체크시 Api 호출
                //    //if (EndCheckEdit.Checked)
                //    //{
                //    //    dataTable.Columns.Add("MoveStdResult");
                //    //    dataTable.Columns.Add("MoveStdBoolResult", typeof(System.Boolean));
                //    //    successCnt = WrGlobal.Camster_Common.NextWorkByUDC(arrTask[0], resourceName, dataTable, false);
                //    //}
                //}
                //else
                //{
                //    WrGlobal.Camster_Common.IsExecuting = false;
                //    MessageBox.Show("제출 가능한 업무가 아닙니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}

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
                            case "Result":
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 세션 종료
                WrGlobal.Camster_Common.DestroySession();
            }

            WrGlobal.Camster_Common.IsExecuting = false;
        }

        // 작업종료 + 다음공정 이동
        private void btnTaskEnd_Click(object sender, EventArgs e)
        {
            if (dataCollectionId == "") return;

            // CamstarCommon Object 생성되지 않았으면 생성 
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }

            if (WrGlobal.Camster_Common.IsExecuting)
            {
                MessageBox.Show("현재 Camstar Interface 기능이 실행 중 입니다.\n잠시 후 다시 제출하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
                CellRange range = worksheet.GetDataRange();

                DataTable dataTable = worksheet.CreateDataTable(range, true);
                dataTable.TableName = "ExcelUpload";

                DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;
                exporter.Export();

                if (dataTable.Rows.Count < 1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출할 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 입력값 검증
                foreach (DataRow dr in dataTable.Rows)
                {
                    if (dr["Container"].ToString() == "")
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show(string.Format("Container : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                WrGlobal.Camster_Common.IsExecuting = true;
                //IsSubmit = true;

                int successCnt = 0;
                int successCnt2 = 0;

                // 세션 생성
                WrGlobal.Camster_Common.CreateSession();

                string[] arrTask = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' });

                // 작업종료 Bool
                dataTable.Columns.Add("BoolResult", typeof(System.Boolean));
                successCnt = WrGlobal.Camster_Common.WorkFinishByUDCResult(navTitle.Caption.ToString(), arrTask[0], resourceName, dataTable, false);

                // 다음공정이동 Bool
                dataTable.Columns.Add("BoolResult2", typeof(System.Boolean));
                successCnt2 = WrGlobal.Camster_Common.NextWorkByUDCHeatResult(arrTask[0], resourceName, dataTable, chk_next_heat.Checked, false);

                if (successCnt == -1 || successCnt2 == -1)
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
                            case "Result":
                                worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                break;
                            case "Result2":
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
                            case "BoolResult2":
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 세션 종료
                WrGlobal.Camster_Common.DestroySession();
            }

            WrGlobal.Camster_Common.IsExecuting = false;
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
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
                    else if (excelSheetControl.Document.Worksheets[0].GetCellValue(0, i).IsNumeric == true)
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

        private void first_name_add()
        {
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

        // Task 변경시 Data1, Data2 검증
        private void TaskLookUpEdit_EditValueChanged(object sender, EventArgs e)
        {
            string taskListName = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' })[0];
            string taskItemName = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' })[1];

            //if (taskItemName != "Data1" && taskItemName != "Data2")
            //{
            //    MessageBox.Show(string.Format("TaskItemName : {0} 은 처리할 수 없습니다.", taskItemName), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}

            if (taskItemName == "Data1")
            {
                //startControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                //endControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //nextControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if (taskItemName == "Data2")
            {
                //startControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //endControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                //nextControlItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            
        }

        // 제출 클릭 ( 사용 안함 )
        private void cmdExecute_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {

        }

        private void txt_1st_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if(e.KeyCode == Keys.Tab)
            {
                first_name_delete();
                first_name_add();
            }

            else if(e.KeyCode == Keys.Enter)
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

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}