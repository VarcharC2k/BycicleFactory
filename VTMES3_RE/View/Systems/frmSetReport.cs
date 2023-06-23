using DevExpress.DashboardWin.Design;
using DevExpress.XtraEditors;
using DevExpress.XtraReports.Extensions;
using DevExpress.XtraReports.UI;
using DevExpress.XtraReports.UserDesigner;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.CodeDataSetTableAdapters;
using VTMES3_RE.Common;
using static DevExpress.XtraEditors.Mask.MaskSettings;

namespace VTMES3_RE.View.Systems
{
    // 양식 설정 폼
    public partial class frmSetReport : DevExpress.XtraEditors.XtraForm
    {
        public frmSetReport()
        {
            InitializeComponent();
        }

        // 로드시 양식 리스트 바인딩
        private void frmSetReport_Load(object sender, EventArgs e)
        {
            this.reportItemTableAdapter.FillByList(storageDataSet.ReportItem, WrGlobal.CorpID);
        }
        // 레포트 등록 버튼 이벤트
        private void cmdReportInsert_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            // 레포트 등록 폼 오픈
            XRDesignForm form = new XRDesignForm();
            form.ShowDialog(this);
            // 레포트 등록 폼이 닫히면 재조회
            this.reportItemTableAdapter.FillByList(storageDataSet.ReportItem, WrGlobal.CorpID);
        }
        // 레포트 수정 버튼 이벤트
        private void cmdReportEdit_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (!gvReportItem.IsDataRow(gvReportItem.FocusedRowHandle)) return;
            if ((gvReportItem.GetRowCellValue(gvReportItem.FocusedRowHandle, "ReportName") ?? "").ToString() == "") return;

            // 레포트 등록 폼 오픈
            XRDesignForm form = new XRDesignForm();
            string url = (gvReportItem.GetRowCellValue(gvReportItem.FocusedRowHandle, "ReportName") ?? "").ToString();
            form.OpenReport(url);
            form.ShowDialog(this);
            // 레포트 등록 폼이 닫히면 재조회
            this.reportItemTableAdapter.FillByList(storageDataSet.ReportItem, WrGlobal.CorpID);
        }
        // 레포트 보기 버튼 이벤트
        private void cmdReportView_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (!gvReportItem.IsDataRow(gvReportItem.FocusedRowHandle)) return;
            if ((gvReportItem.GetRowCellValue(gvReportItem.FocusedRowHandle, "ReportName") ?? "").ToString() == "") return;
            // 선택된 레포트 뷰어 오픈
            XtraReport report = GetSelectedReport((gvReportItem.GetRowCellValue(gvReportItem.FocusedRowHandle, "ReportName") ?? "").ToString());
            if (report != null)
                report.ShowPreviewDialog();
        }
        /// <summary>
        /// 레포트를 ReportStorage 에서 읽어 XtraReport 리턴
        /// </summary>
        /// <param name="url">선택된 ReportName</param>
        /// <returns></returns>
        XtraReport GetSelectedReport(string url)
        {
            // Return a report by a URL selected in the ListBox.
            if (url == "")
                return null;
            using (MemoryStream stream = new MemoryStream(Program.ReportStorage.GetData(url)))
            {
                return XtraReport.FromStream(stream, true);
            }
        }

        // 저장 버튼 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                // 양식 명 수정시 LayoutData 에도 적용해야 함
                foreach (DataRowView drv in reportItemBindingSource.List)
                {
                    if (drv.Row.RowState == DataRowState.Modified)
                    {    
                        drv["LayoutData"] = drv["LayoutData"].ToString().Replace("DisplayName=\"" + drv["OldName"].ToString() + "\"", "DisplayName=\"" + drv["ReportName"].ToString() + "\"");
                        drv["OldName"] = drv["ReportName"];
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModIP"] = WrGlobal.ClientHostName;
                        drv["ModDt"] = DateTime.Now;
                    }
                }


                // reportItem 테이블 저장
                reportItemBindingSource.EndEdit();
                reportItemTableAdapter.Update(storageDataSet.ReportItem);

                // reportStorage 재로드
                Program.reportStorage = new DataSetReportStorage();
                ReportStorageExtension.RegisterExtensionGlobal(Program.reportStorage);

                MessageBox.Show("저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
        // 삭제 버튼 클릭 이벤트
        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (reportItemBindingSource.Current == null) return;

            if (MessageBox.Show("선택한 자료를 삭제 하시겠습니까?", "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            try
            {
                //reportItem 테이블에서 삭제
                reportItemBindingSource.RemoveCurrent();
                reportItemTableAdapter.Update(storageDataSet.ReportItem);

                MessageBox.Show("삭제되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // 조회 버튼 클릭 이벤트 -> 전체 양식 조회
        private void cmdDisplay_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.reportItemTableAdapter.FillByList(storageDataSet.ReportItem, WrGlobal.CorpID);
        }
    }
}