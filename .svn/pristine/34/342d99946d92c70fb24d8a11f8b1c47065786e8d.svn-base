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
using VTMES3_RE.IFRYDataSetTableAdapters;

namespace VTMES3_RE.View.ProductInfo
{
    // AS IS 템플릿 항목 관리와 유사
    // MES3 애서 UDCD와 Datapoint 관리 폼
    public partial class frmCheckDefinition : DevExpress.XtraEditors.XtraForm
    {
        public frmCheckDefinition()
        {
            InitializeComponent();
        }

        // 로드시 전체 검사 그룹 바인딩
        private void frmCheckDefinition_Load(object sender, EventArgs e)
        {
            this.checkMasterTableAdapter.Fill(this.iFRYDataSet.CheckMaster);
        }

        // 검사 그룹 변경시 검사 항목 바인딩
        private void gvCheckMaster_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string collectionName = (gvCheckMaster.GetFocusedRowCellValue("DataCollectionDefName") ?? "").ToString();
            string revName = (gvCheckMaster.GetFocusedRowCellValue("DataCollectionDefRevision") ?? "").ToString();

            this.checkDetailTableAdapter.FillBySelect(this.iFRYDataSet.CheckDetail, collectionName, revName);
        }
        // 신규 검사항목 입력 시 검사 그룹 항목값 기본 설정해줌
        private void gvCheckDetail_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            string collectionName = (gvCheckMaster.GetFocusedRowCellValue("DataCollectionDefName") ?? "").ToString();
            string revName = (gvCheckMaster.GetFocusedRowCellValue("DataCollectionDefRevision") ?? "").ToString();

            if (collectionName == "") return;
            if (revName == "") return;

            DataRowView drv = (DataRowView)checkDetailBindingSource.Current;

            drv["DataCollectionDefName"] = collectionName;
            drv["DataCollectionDefRevision"] = revName;
        }
        // 저장 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();
                // 입력자, 입력일 설정
                foreach (DataRowView drv in checkMasterBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        drv["CreId"] = WrGlobal.LoginID;
                        drv["CreDt"] = DateTime.Now;
                    }
                    else if (drv.Row.RowState == DataRowState.Modified)
                    {
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModDt"] = DateTime.Now;
                    }
                }
                // 수정자, 수정일 설정
                foreach (DataRowView drv in checkDetailBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        drv["CreId"] = WrGlobal.LoginID;
                        drv["CreDt"] = DateTime.Now;
                    }
                    else if (drv.Row.RowState == DataRowState.Modified)
                    {
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModDt"] = DateTime.Now;
                    }
                }
                // 검사그룹 저장
                checkMasterBindingSource.EndEdit();
                checkMasterTableAdapter.Update(this.iFRYDataSet.CheckMaster);
                // 검사항목 저장
                checkDetailBindingSource.EndEdit();
                checkDetailTableAdapter.Update(this.iFRYDataSet.CheckDetail);

                MessageBox.Show("작업한 내역이 저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 검사 그룹 삭제
        private void btnDelete1_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == DevExpress.XtraEditors.Controls.ButtonPredefines.Delete)
            {
                if (MessageBox.Show(string.Format("선택한 검사 그룹을 삭제하시겠습니까?"), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }//end if

                checkMasterBindingSource.RemoveCurrent();
                checkMasterTableAdapter.Update(this.iFRYDataSet.CheckMaster);

                MessageBox.Show("자료가 삭제되었습니다.", "삭제", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        // 검사 항목 삭제
        private void btnDelete2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == DevExpress.XtraEditors.Controls.ButtonPredefines.Delete)
            {
                if (MessageBox.Show(string.Format("선택한 검사 항목을 삭제하시겠습니까?"), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
                {
                    return;
                }//end if

                checkDetailBindingSource.RemoveCurrent();
                checkDetailTableAdapter.Update(this.iFRYDataSet.CheckDetail);

                MessageBox.Show("자료가 삭제되었습니다.", "삭제", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}