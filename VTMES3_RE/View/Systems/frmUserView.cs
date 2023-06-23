using DevExpress.XtraGrid.Views.Grid;
using VTMES3_RE.Common;
using VTMES3_RE.Models;
using System;
using System.Data;
using System.Windows.Forms;

namespace VTMES3_RE.View.Systems
{
    public partial class frmUserView : DevExpress.XtraEditors.XtraForm
    {
        // clsCode 모델 생성
        clsCode user = new clsCode();

        public frmUserView()
        {
            InitializeComponent();

            authorCheckedComboBoxEdit.DataSource = user.GetGroupAuthorList();
        }
        private void frmUserView_Load(object sender, EventArgs e)
        {
            // ReportDB의 CodeUser 테이블 조회
            cmdDisplay_ElementClick(null, null);
        }
        // ReportDB의 CodeUser 테이블 조회
        private void DisplayData()
        {
            codeUserTableAdapter.FillByList(codeDataSet.CodeUser, WrGlobal.CorpID);
        }

        // 사용자 정보 저장 버튼 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                // Password2에 값이 잇으면 -> 암호화 하여 기존 비밀번호 변경
                foreach (DataRowView drv in codeUserBindingSource.List)
                {
                    if (drv.Row.RowState == DataRowState.Added || drv.Row.RowState == DataRowState.Modified)
                    {
                        if (drv["Password2"].ToString() != "")
                        {
                            drv["Password"] = clsCommon.SHA256Hash(drv["Password2"].ToString());
                            drv["Password2"] = "";
                        }
                        //if (drv["TeamName"].ToString().ToUpper() != "TFT" && drv["TeamName"].ToString().ToUpper() != "CMOS" && drv["TeamName"].ToString().ToUpper() != "CSI")
                        //{
                        //    MessageBox.Show("팀명이 잘못되었습니다 팀명을 확인해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //    return;
                        //}

                    }
                }
                // 사용자 정보 저장
                codeUserBindingSource.EndEdit();
                codeUserTableAdapter.Update(codeDataSet.CodeUser);
                MessageBox.Show("사용자 정보를 저장했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 사용자 삭제 버튼 클릭 이벤트
        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            string Id = gvCodeUser.GetFocusedRowCellDisplayText("EmployeeName") == null ? "" : gvCodeUser.GetFocusedRowCellDisplayText("EmployeeName");

            if (MessageBox.Show(string.Format("선택한 '{0}' 사용자를 삭제하시겠습니까?", Id), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }//end if
            // 선택된 사용자 삭제
            codeUserBindingSource.RemoveCurrent();
            codeUserTableAdapter.Update(codeDataSet.CodeUser);
            MessageBox.Show("자료가 삭제 되었습니다.");
        }
        // 엑셀출력
        private void cmdExcel_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Worksheets|*.Xls";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                gvCodeUser.ExportToXls(sfd.FileName);
                System.Diagnostics.Process.Start(sfd.FileName);
            }//end fnction
        }
        // 닫기
        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
        // 조회
        private void cmdDisplay_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            DisplayData();
        }
        // 선택된 로우 색상 표시
        private void gvUser_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (e.RowHandle == view.FocusedRowHandle)
            {
                //Apply the appearance of the SelectedRow
                e.Appearance.Assign(view.PaintAppearance.SelectedRow);
                e.Appearance.Options.UseForeColor = true;
            }//end if
        }

    }
}