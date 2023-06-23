using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.Xpo.DB.Helpers;
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
    public partial class frmExecTaskByCmosLot : DevExpress.XtraEditors.XtraForm
    {
        clsWork work = new clsWork();
        Database2 ifrydb = new Database2();
        string query = "";

        DataView collectionView = null;
        DataView collectionView2 = null;

        string resourceName = "";
        string dataCollectionId = "";
        string dataCollectionName = "";
        bool IsSubmit = false;

        public frmExecTaskByCmosLot()
        {
            InitializeComponent();

            ResourceGroupLookUpEdit.Properties.DataSource = work.GetResourceGroup2();
            ResourceGroupLookUpEdit.Properties.DisplayMember = "그룹명";
            ResourceGroupLookUpEdit.Properties.ValueMember = "그룹명";

            //DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection2("CMOS_");
            //DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
            //DataCollectionLookUpEdit.Properties.ValueMember = "코드";
            //DataCollectionLookUpEdit.EditValue = null;

            StartCheckEdit.Visible = true;
            EndCheckEdit.Visible = true;
            NextCheckEdit.Visible = true;
        }

        private void frmExecTaskByCmosLot_Load(object sender, EventArgs e)
        {
            //btnPqcOpen.Visible = false;
            BtnSchOQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        //설비 그룹 비활성화
        //private void ResourceGroupLookUpEdit_EditValueChanged(object sender, EventArgs e)
        //{
        //    ResourceLookUpEdit.Properties.DataSource = work.GetResourceDef((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
        //    ResourceLookUpEdit.Properties.DisplayMember = "설비명";
        //    ResourceLookUpEdit.Properties.ValueMember = "설비명";
        //    ResourceLookUpEdit.EditValue = null;

        //    //DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
        //    DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection2("CMOS_");
        //    DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
        //    DataCollectionLookUpEdit.Properties.ValueMember = "코드";
        //    DataCollectionLookUpEdit.EditValue = null;
        //}

        private void btnSearch_Click(object sender, EventArgs e)
        {
            //설비 그룹 비활성화
            //if ((ResourceGroupLookUpEdit.EditValue ?? "").ToString() == "") return;

            //if ((ResourceLookUpEdit.EditValue ?? "").ToString() == "") return; ==> 설비로직
            if ((DataCollectionLookUpEdit.EditValue ?? "").ToString() == "") return;

            string test = "";

            test = DataCollectionLookUpEdit.Text;

            DataView taskView = work.GetTaskInfoByCollection_Origin((DataCollectionLookUpEdit.EditValue ?? "").ToString());
            if (taskView.Count == 0)
            {
                MessageBox.Show("선택된 DataCollection과 연결된 Task를 찾을수 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((DataCollectionLookUpEdit.Text ?? "").ToString().Contains("QC"))
            {
                BtnSchOQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                BtnSchOQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }

            TaskLookUpEdit.Properties.DataSource = taskView;
            TaskLookUpEdit.Properties.ValueMember = "TaskValue";
            TaskLookUpEdit.Properties.DisplayMember = "TaskName";

            TaskLookUpEdit.EditValue = taskView[0]["TaskValue"].ToString();     //Task DropDownList에 자동으로 삽입되는 로직

            excelSheetControl.CreateNewDocument();

            DataTable dataTable = new DataTable("DataPoint");
            dataTable.Columns.Add("Container", typeof(System.String));          //Container 컬럼을 추가하는 로직

            collectionView = work.GetDataPointByCollection((DataCollectionLookUpEdit.EditValue ?? "").ToString());

            foreach (DataRowView drv in collectionView)
            {
                dataTable.Columns.Add(drv["DataPointName"].ToString(), typeof(System.String));

                //collectionView2 = work.GetDataPointByCollection(drv["DataPointName"].ToString());

            }

            // 결과값을 표시하기 위한 추가
            if (test.Contains("작업시작"))
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
            var externalDSOptions = new ExternalDataSourceOptions();
            externalDSOptions.ImportHeaders = true;
            excelSheetControl.Document.Worksheets[0].DataBindings.BindTableToDataSource(dataTable, 0, 0, externalDSOptions);

            // DataPoint 가 NamedObjectGroupName 일때 콤보 처리
            int cnt = 1;
            Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
            CellRange range = worksheet.GetDataRange();

            string dfv = string.Empty;

            foreach (DataRowView drv in collectionView)
            {
                if (drv["DataType"].ToString() == "5" && drv["NamedObjectGroupName"].ToString() != "")
                {
                    CellRange comboBoxRange = excelSheetControl.Document.Worksheets[0][string.Format("DataPoint[{0}]", drv["DataPointName"].ToString())];
                    excelSheetControl.Document.Worksheets[0].CustomCellInplaceEditors.Add(comboBoxRange, CustomCellInplaceEditorType.ComboBox, drv["NamedObjectGroupName"].ToString());
                }

                dfv = drv["dfv"].ToString();

                if (dfv == string.Empty)
                {
                    cnt++;
                    continue;
                } 
                else
                {
                    range[0, cnt].FillColor = Color.Red;
                    cnt++;
                }
                


                //foreach (DataRow dr in dataTable.Rows)
                //{
                //    if (dr[drv["DataPointName"].ToString()].ToString() != "")
                //    {
                //        if (drv["NamedObjectGroupName"].ToString().IndexOf(dr[drv["DataPointName"].ToString()].ToString()) < 0)
                //        {
                //            dr[drv["DataPointName"].ToString()] = Color.Red;
                //            return;
                //        }
                //    }
                //}

            }

            resourceName = (ResourceLookUpEdit.EditValue ?? "").ToString();
            dataCollectionId = (DataCollectionLookUpEdit.EditValue ?? "").ToString();
            dataCollectionName = DataCollectionLookUpEdit.Text;
            IsSubmit = false;

            //string collectionName = (gvCheckMaster.GetFocusedRowCellValue("DataPointName") ?? "").ToString();
            //string revName = (gvCheckMaster.GetFocusedRowCellValue("DataCollectionDefRevision") ?? "").ToString();

            //if (collectionName == "") return;
            //if (revName == "") return;

            ////DataRowView drv = (DataRowView)checkDetailBindingSource.Current;

            //foreach (DataRowView drv in collectionView)
            //{
            //    drv["DataCollectionDefName"] = collectionName;
            //    //drv["DataCollectionDefRevision"] = revName;

            //    if (drv["DataCollectionDefName"].ToString() != "")
            //    {
            //        foreach (DataRow dr in dataTable.Rows)
            //        {
            //            if (dr[drv["DataPointName"].ToString()].ToString() != "")
            //            {
            //                if (drv["NamedObjectGroupName"].ToString().IndexOf(dr[drv["DataPointName"].ToString()].ToString()) < 0)
            //                {
            //                    //drv["DataPointName"].ToString(). = Color.Red;
            //                    //return;
            //                }
            //            }
            //        }
            //    }
            //}


        }

        // 제출
        private void cmdExecute_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (dataCollectionId == "") return;

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

            if (excelSheetControl.IsCellEditorActive)
            {
                excelSheetControl.CloseCellEditor(CellEditorEnterValueMode.ActiveCell);
            }

            try
            {
                Worksheet worksheet = excelSheetControl.Document.Worksheets[0];
                CellRange range = worksheet.GetDataRange();  //데이터가 있는 셀의 범위를 가져옴.

                DataTable dataTable = worksheet.CreateDataTable(range, true); //DataTable을 생성
                dataTable.TableName = "ExcelUpload";

                DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
                exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;

                exporter.CellValueConversionError += exporter_CellValueConversionError;  //에러발생시 호출하는 함수
                exporter.Export();

                if (dataTable.Rows.Count < 1)
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출할 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                foreach (DataRow dr in dataTable.Rows)
                {
                    if (dr["Container"].ToString() == "")
                    {
                        WrGlobal.Camster_Common.IsExecuting = false;
                        MessageBox.Show(string.Format("Container : 필수 입력 항목입니다."), "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

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
                //WrGlobal.Camstar_UserName = "Administrator";
                //WrGlobal.Camstar_Password = "Rkddkwl2014!@";

                WrGlobal.Camster_Common.CreateSession();

                string[] arrTask = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' });

                if (arrTask[1] == "Data1")
                {
                    // 작업시작
                    if (StartCheckEdit.Checked)
                    {
                        dataTable.Columns.Add("StartResult");
                        dataTable.Columns.Add("StartBoolResult", typeof(System.Boolean));
                        successCnt = WrGlobal.Camster_Common.WorkStart(arrTask[0], resourceName, dataTable, false);
                    }

                    dataTable.Columns.Add("TaskResult");
                    dataTable.Columns.Add("TaskBoolResult", typeof(System.Boolean));
                    successCnt = WrGlobal.Camster_Common.ExecuteTaskByUDC(arrTask[0], arrTask[1], dataCollectionName, dataTable, collectionView, false);
                }
                else if (arrTask[1] == "Data2")
                {
                    dataTable.Columns.Add("TaskResult");
                    dataTable.Columns.Add("TaskBoolResult", typeof(System.Boolean));
                    successCnt = WrGlobal.Camster_Common.ExecuteTaskByUDC(arrTask[0], arrTask[1], dataCollectionName, dataTable, collectionView, false);

                    //작업종료
                    if (EndCheckEdit.Checked)
                    {
                        dataTable.Columns.Add("EndResult");
                        dataTable.Columns.Add("EndBoolResult", typeof(System.Boolean));
                        successCnt = WrGlobal.Camster_Common.WorkFinishByUDC(arrTask[0], resourceName, dataTable, false);
                    }

                    // 다음공정으로
                    if (EndCheckEdit.Checked)
                    {
                        dataTable.Columns.Add("MoveStdResult");
                        dataTable.Columns.Add("MoveStdBoolResult", typeof(System.Boolean));
                        successCnt = WrGlobal.Camster_Common.NextWorkByUDC(arrTask[0], resourceName, dataTable, false);
                    }
                }
                else
                {
                    WrGlobal.Camster_Common.IsExecuting = false;
                    MessageBox.Show("제출 가능한 업무가 아닙니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

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

                WrGlobal.Camster_Common.DestroySession();
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

        private void txtContainerId_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection3((txtContainerId.EditValue ?? "").ToString());
                DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
                DataCollectionLookUpEdit.Properties.ValueMember = "코드";
            }
        }

        private void txtContainerId_EditValueChanged(object sender, EventArgs e)
        {
            //DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection3((txtContainerId.EditValue ?? "").ToString());
            //DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
            //DataCollectionLookUpEdit.Properties.ValueMember = "코드";

            //DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection((ResourceGroupLookUpEdit.EditValue ?? "").ToString());
                     
          

            DataCollectionLookUpEdit.Properties.DataSource = work.GetDataCollection4("CMOS_", txtContainerId.Text.ToString());
            DataCollectionLookUpEdit.Properties.DisplayMember = "명칭";
            DataCollectionLookUpEdit.Properties.ValueMember = "코드";
            DataCollectionLookUpEdit.EditValue = null;
        }

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
                                else
                                {
                                    for (int k = 0; k < range.ColumnCount; k++)
                                    {
                                        range[i + 1, k].FillColor = Color.Empty;
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
                                else
                                {
                                    for (int k = 0; k < range.ColumnCount; k++)
                                    {
                                        range[i + 1, k].FillColor = Color.Empty;
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
                successCnt = WrGlobal.Camster_Common.WorkFinishByUDCResult(navTitle.Caption.ToString(),arrTask[0], resourceName, dataTable, false);

                // 다음공정이동 Bool
                dataTable.Columns.Add("BoolResult2", typeof(System.Boolean));
                successCnt2 = WrGlobal.Camster_Common.NextWorkByUDCResult(arrTask[0], resourceName, dataTable, false);

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
                                else
                                {
                                    for (int k = 0; k < range.ColumnCount; k++)
                                    {
                                        range[i + 1, k].FillColor = Color.Empty;
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
                                else
                                {
                                    for (int k = 0; k < range.ColumnCount; k++)
                                    {
                                        range[i + 1, k].FillColor = Color.Empty;
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

        private void btnPqcOpen_Click(object sender, EventArgs e)
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
                    MessageBox.Show("불러올 컨테이너 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                //Camstar와의 커넥션 실행 중인지 확인 
                WrGlobal.Camster_Common.IsExecuting = true;
                //IsSubmit = true;

                int successCnt = 0;

                // 세션 생성
                WrGlobal.Camster_Common.CreateSession();

                string[] arrTask = TaskLookUpEdit.EditValue.ToString().Split(new char[] { '|' });

                //PQC Data 불러오기 클릭
                if (arrTask[1].Contains("QC"))
                {
                    //BtnSchOQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    //BtnSchOQC.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInRuntime;
                    dataTable.Columns.Add("BoolResult", typeof(System.Boolean));

                    //arrTask[0] : 작시 작종
                    //arrTask[1] : UDCD
                    //arrTask[2] : UDCD의 taskName
                    //dataTable  : 화면에 뿌려진 엑셀
                    //successCnt = WrGlobal.Camster_Common.PQCDataOpen(arrTask[0], resourceName, dataTable, false);
                    //DataCollectionLookUpEdit.EditValue : 작업 종류의 id 값
                    successCnt = WrGlobal.Camster_Common.CMOSPQCDataOpenNew(DataCollectionLookUpEdit.EditValue.ToString(), dataTable, false);

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
                            worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);

                            switch (dataTable.Columns[j].ColumnName)
                            {
                                case "Result":
                                    worksheet.Cells[range.TopRowIndex + 1 + i, range.LeftColumnIndex + j].SetValue(dataTable.Rows[i][j]);
                                    break;
                                case "BoolResult":
                                    //if (!Convert.ToBoolean(dataTable.Rows[i][j]))
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Red;
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    for (int k = 0; k < range.ColumnCount; k++)
                                    //    {
                                    //        range[i + 1, k].FillColor = Color.Empty;
                                    //    }
                                    //}
                                    break;
                            }
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
    }
}