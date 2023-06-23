using DevExpress.XtraEditors;
using DevExpress.XtraExport.Helpers;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraSpreadsheet.Layout;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.CodeDataSetTableAdapters;
using VTMES3_RE.Common;
//using static DevExpress.Data.Mask.Internal.MaskSettings<T>;

namespace VTMES3_RE.View.ProductionManagement
{
    // 이순희 PM 작성 폼
    public partial class frmBatchPlan : DevExpress.XtraEditors.XtraForm
    {
        public frmBatchPlan()
        {
            InitializeComponent();
            // 이번달 1일 ~ 현재일
            dt_start.EditValue = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            //dt_start.EditValue = DateTime.Today;
            dt_end.EditValue = DateTime.Today;
        }

        private void frmBatchPlan_Load(object sender, EventArgs e)
        {
            // TODO: 이 코드는 데이터를 'iFRYDataSet.V_RY_CSI_RESOURCE_DEPO' 테이블에 로드합니다. 필요 시 이 코드를 이동하거나 제거할 수 있습니다.
            this.v_RY_CSI_RESOURCE_DEPOTableAdapter.Fill(this.iFRYDataSet.V_RY_CSI_RESOURCE_DEPO);
            // TODO: 이 코드는 데이터를 'iFRYDataSet1.V_RY_CSI_RESOURCE_DEPO' 테이블에 로드합니다. 필요 시 이 코드를 이동하거나 제거할 수 있습니다.
            //this.v_RY_CSI_RESOURCE_DEPOTableAdapter.Fill(this.iFRYDataSet1.V_RY_CSI_RESOURCE_DEPO);


            // TODO: 이 코드는 데이터를 'iFRYDataSet.CsI_Batch_Plan' 테이블에 로드합니다. 필요 시 이 코드를 이동하거나 제거할 수 있습니다.
            this.csI_Batch_PlanTableAdapter.FillByBatchPlanDate(this.iFRYDataSet.CsI_Batch_Plan, (DateTime)dt_start.EditValue, (DateTime)dt_end.EditValue);
            // TODO: 이 코드는 데이터를 'iFRYDataSet.csI_Batch_Plan' 테이블에 로드합니다. 필요 시 이 코드를 이동하거나 제거할 수 있습니다.
            //this.csI_Batch_PlanTableAdapter.Fill(this.iFRYDataSet.csI_Batch_Plan);

        }
        private void cmdAdd_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            csI_Batch_PlanBindingSource.AddNew();
            DataRowView drv = (DataRowView)csI_Batch_PlanBindingSource.Current;

        }
        private void gvCodeUser_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            //DataRowView drv = (DataRowView)csI_Batch_PlanBindingSource.Current;
            //if (drv.Row.RowState == DataRowState.Detached || drv.Row.RowState == DataRowState.Added)
            //{
            //    colMODEL_CODE.OptionsColumn.AllowEdit = true;
            //}
            //else
            //{
            //    colMODEL_CODE.OptionsColumn.AllowEdit = false;

            //}
        }

        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                //foreach (DataRowView drv in csI_Batch_PlanBindingSource)
                //{
                //    if (drv.Row.RowState == DataRowState.Added)
                //    {
                //        if ((drv["UseYn"] ?? "").ToString() == "")
                //        {
                //            drv["UseYn"] = "Y";
                //        }
                //    }
                //}

                csI_Batch_PlanBindingSource.EndEdit();
                csI_Batch_PlanTableAdapter.Update(iFRYDataSet.CsI_Batch_Plan);
                MessageBox.Show("Batch 계획을 저장했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            string Id = gvModelCode.GetFocusedRowCellDisplayText("MODEL_CODE") == null ? "" : gvModelCode.GetFocusedRowCellDisplayText("MODEL_CODE");

            if (MessageBox.Show(string.Format("선택한 '{0}' 모델을 삭제하시겠습니까?", Id), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }//end if

            csI_Batch_PlanBindingSource.RemoveCurrent();
            csI_Batch_PlanTableAdapter.Update(iFRYDataSet.CsI_Batch_Plan);
            MessageBox.Show("자료가 삭제 되었습니다.");
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            this.csI_Batch_PlanTableAdapter.FillByBatchPlanDate(this.iFRYDataSet.CsI_Batch_Plan, (DateTime)dt_start.EditValue, (DateTime)dt_end.EditValue);
        }

        private void gvModelCode_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            //csI_Batch_PlanBindingSource

            DataRowView drv = (DataRowView)csI_Batch_PlanBindingSource.Current;

            drv["DAY_NIGHT"] = "주간";
            drv["BATCH_DATE"] = DateTime.Today;
            drv["QTY"] = 1;

            // KMC 22.12.14
            // CSI 증착기 view 의 첫번째 이름을 가져오기
            //foreach (DataRowView item in v_RY_CSI_RESOURCE_DEPOBindingSource)
            //{
            //    drv["DEPO"] = item[0];
            //    break;
            //}

            // LSH 22.12.14
            // CSI 증착기 view 의 첫번째 이름을 가져오기
            DataRowView item = (DataRowView)v_RY_CSI_RESOURCE_DEPOBindingSource.Current;
            drv["DEPO"] = item.Row["NickName"].ToString();

        }
    }
}