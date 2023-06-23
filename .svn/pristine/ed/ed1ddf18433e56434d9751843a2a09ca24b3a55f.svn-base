using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraWaitForm;
using Microsoft.Win32;
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

namespace VTMES3_RE.View.Approval
{
    public partial class frmESigManager_CMOS_Runsheet : DevExpress.XtraEditors.XtraForm
    {
        clsWork camWork = new clsWork();
        string query = "";

        public frmESigManager_CMOS_Runsheet()
        {
            InitializeComponent();

            barEditStartDate.EditValue = DateTime.Today.AddDays(-1);
            barEditEndDate.EditValue = DateTime.Today;
        }

        private void frmESigManager_Load(object sender, EventArgs e)
        {
            // Camstar Esig Role 확인 -> 검토자, 승인자 세팅
            if (WrGlobal.Camstar_RoleName == "검토자")
            {
                cmdCheck.Visible = true;
                searchApprovaYnComboBoxEdit.EditValue = "검토대기";
            }
            else if (WrGlobal.Camstar_RoleName == "승인자")
            {
                cmdApprove.Visible = true;
                searchApprovaYnComboBoxEdit.EditValue = "검토완료";
            }
            // 검색 리스트 표시
            DisplayEsigHistory();
        }
        /// <summary>
        /// 조건에 대한 결재 항목 리스트
        /// </summary>
        private void DisplayEsigHistory()
        {
            query = string.Format("exec CAMDBsh.RY_VM_Proc_CMOS_Runsheet_GetESigHistory '{0:yyyy-MM-dd}', '{1:yyyy-MM-dd}', N'{2}'",
                (DateTime)barEditStartDate.EditValue, (DateTime)barEditEndDate.EditValue, searchApprovaYnComboBoxEdit.Text);

            gcESigHistory.DataSource = camWork.GetDataViewByQuery("EsigHistory", query);
        }

        //엑셀출력
        private void cmdExcel_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Worksheets|*.Xls";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                XlsExportOptionsEx o = new XlsExportOptionsEx { ExportType = DevExpress.Export.ExportType.WYSIWYG };

                gvESigHistory.ExportToXls(sfd.FileName, o);
                System.Diagnostics.Process.Start(sfd.FileName);
            }//end fnction
        }
        //닫기
        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
        //검토 처리 버튼클릭
        private void cmdCheck_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (gvESigHistory.SelectedRowsCount == 0)
            {
                MessageBox.Show("선택된 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int[] idxs = gvESigHistory.GetSelectedRows();
            int cnt = 0;
            List<string> queryList = new List<string>();

            foreach (int idx in idxs)
            {
                if (!gvESigHistory.IsDataRow(idx)) continue;

                if ((gvESigHistory.GetRowCellValue(idx, "CheckUser") ?? "").ToString() == "")
                {   // 미검토 일때
                    if ((gvESigHistory.GetRowCellValue(idx, "ApporveUser") ?? "").ToString() == "")
                    {  //미승인 이면 검토 처리 Insert
                        query = string.Format("INSERT INTO IFRY.dbo.VM_ESigHistory(StepEntryTxnId, FactoryId, ContainerId, Gubun, WorkflowStepId, CheckUser, CheckDate, CheckComment, CheckUserName) "
                                            + "VALUES('{0}', '{1}', '{2}', 'Runsheet', '{3}', '{4}', GETDATE(), '{5}', '{6}')"
                                            , gvESigHistory.GetRowCellValue(idx, "StepEntryTxnId").ToString()
                                            , gvESigHistory.GetRowCellValue(idx, "FactoryId").ToString()
                                            , gvESigHistory.GetRowCellValue(idx, "ContainerId").ToString()
                                            , gvESigHistory.GetRowCellValue(idx, "WorkflowStepId").ToString()
                                            , WrGlobal.EmployeeId
                                            , gvESigHistory.GetRowCellValue(idx, "CheckComment").ToString()
                                            , WrGlobal.LoginID);

                        queryList.Add(query);

                        cnt++;
                    }
                    else
                    {  // 승인 이면 Update x -> 에러
                        MessageBox.Show("승인완료된 항목이 선택되었습니다.\n처리에 실패했습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("검토완료된 항목이 선택되었습니다.\n처리에 실패했습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (cnt > 0)
            {
                // 체크된 항목 쿼리문 실행
                int success = camWork.ExecuteQryList(queryList);

                if (cnt == success)
                {
                    MessageBox.Show(string.Format("{0}건 검토 처리되었습니다.", success), "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(string.Format("{0}건 검토 처리되었습니다. ({1}건 처리 실패)", success, cnt - success), "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (success > 0)
                {
                    DisplayEsigHistory();
                }
            }

        }
        //승인 처리 버튼클릭
        private void cmdApprove_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (gvESigHistory.SelectedRowsCount == 0)
            {
                MessageBox.Show("선택된 항목이 없습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 체크된 항목 가져오기
            int[] idxs = gvESigHistory.GetSelectedRows();
            int cnt = 0;
            List<string> queryList = new List<string>();

            // 체크 건별 승인 처리
            foreach (int idx in idxs)
            {
                if (!gvESigHistory.IsDataRow(idx)) continue;

                if ((gvESigHistory.GetRowCellValue(idx, "ApproveUser") ?? "").ToString() == "")
                {   // 미승인 일때
                    if ((gvESigHistory.GetRowCellValue(idx, "CheckUser") ?? "").ToString() == "")
                    {   //미검토건은 Insert x -> 에러
                        MessageBox.Show("검토완료 처리되지 않은 항목이 선택되었습니다.\n처리에 실패했습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {   // 검토된 건은 승인 처리 Update
                        query = string.Format("Update IFRY.dbo.VM_ESigHistory "
                                                + "Set ApproveUser = '{3}', "
                                                + "ApproveDate = GETDATE(), "
                                                + "ApproveComment = '{4}', "
                                                + "ApproveUserName = '{5}' "
                                        + "Where StepEntryTxnId = '{0}' and FactoryId = '{1}' and ContainerId = '{2}'"
                                        , gvESigHistory.GetRowCellValue(idx, "StepEntryTxnId").ToString()
                                        , gvESigHistory.GetRowCellValue(idx, "FactoryId").ToString()
                                        , gvESigHistory.GetRowCellValue(idx, "ContainerId").ToString()
                                        , WrGlobal.EmployeeId
                                        , gvESigHistory.GetRowCellValue(idx, "ApproveComment").ToString()
                                        , WrGlobal.LoginID);

                        queryList.Add(query);

                        cnt++;
                    }
                }
                else
                {
                    MessageBox.Show("승인완료된 항목이 선택되었습니다.\n처리에 실패했습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (cnt > 0)
            {
                // 체크된 항목 쿼리문 실행
                int success = camWork.ExecuteQryList(queryList);

                if (cnt == success)
                {
                    MessageBox.Show(string.Format("{0}건 승인 처리되었습니다.", success), "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(string.Format("{0}건 승인 처리되었습니다. ({1}건 처리 실패)", success, cnt - success), "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (success > 0)
                {
                    DisplayEsigHistory();
                }
            }
        }
        /// <summary>
        /// 검토 완료: 노란색, 승인 완료 : 오렌지색 처리
        /// </summary>
        private void gvESigHistory_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView gridView = sender as GridView;;

            if (e.RowHandle < 0) return;

            if ((gridView.GetRowCellValue(e.RowHandle, "CheckUser") ?? "").ToString() != "" && (gridView.GetRowCellValue(e.RowHandle, "ApproveUser") ?? "").ToString() != "")
            {
                e.Appearance.BackColor = Color.Orange;
                e.HighPriority = true;
            }
            else if ((gridView.GetRowCellValue(e.RowHandle, "CheckUser") ?? "").ToString() != "" || (gridView.GetRowCellValue(e.RowHandle, "ApproveUser") ?? "").ToString() != "")
            {
                e.Appearance.BackColor = Color.Yellow;
            }
        }

        // 조건 검색 클릭
        private void btnSearch_Click(object sender, EventArgs e)
        {
            DisplayEsigHistory();
        }
        // 레포트 보기 클릭 -> 해당 웹 레포트 URL 표시
        private void btnReport_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataRowView drv = (DataRowView)gvESigHistory.GetRow(gvESigHistory.FocusedRowHandle);

            string reportUrl = String.Format(@"{0}/View/QC/CMOS/Report_QC_CMOS_REPORT.aspx?cmd=esig&reportid=78&ContainerName={1}&StepEntryTxnId={2}",
                            WrGlobal.reportRootUrl, drv["ContainerName"].ToString(), drv["StepEntryTxnId"].ToString());
            layoutControlGroup3.Text = string.Format("런시트 : {0} ", drv["ContainerName"].ToString());
            webView1.Source = new Uri(reportUrl);
        }
    }
}