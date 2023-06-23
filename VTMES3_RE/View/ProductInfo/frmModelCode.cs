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

namespace VTMES3_RE.View.ProductInfo
{
    // 김성재 주임 작성
    // 김문철 부장 -> 성적서 관련 필요 컬럼 VIsble 처리

    public partial class frmModelCode : DevExpress.XtraEditors.XtraForm
    {
        public frmModelCode()
        {
            InitializeComponent();
        }

        private void frmModelCode_Load(object sender, EventArgs e)
        {
            this.mES2_MODEL_MASTERTableAdapter.Fill(this.iFRYDataSet.MES2_MODEL_MASTER);

        }
        private void cmdAdd_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            mES2_MODEL_MASTERBindingSource.AddNew();
            DataRowView drv = (DataRowView)mES2_MODEL_MASTERBindingSource.Current;

        }
        private void gvCodeUser_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            DataRowView drv = (DataRowView)mES2_MODEL_MASTERBindingSource.Current;
            if (drv.Row.RowState == DataRowState.Detached || drv.Row.RowState == DataRowState.Added)
            {
                colMODEL_CODE.OptionsColumn.AllowEdit = true;
            }
            else
            {
                colMODEL_CODE.OptionsColumn.AllowEdit = false;

            }
        }

        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                foreach (DataRowView drv in mES2_MODEL_MASTERBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        if ((drv["UseYn"] ?? "").ToString() == "")
                        {
                            drv["UseYn"] = "Y";
                        }
                    }
                }

                mES2_MODEL_MASTERBindingSource.EndEdit();
                mES2_MODEL_MASTERTableAdapter.Update(iFRYDataSet.MES2_MODEL_MASTER);
                MessageBox.Show("모델정보를 저장했습니다.");
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

            mES2_MODEL_MASTERBindingSource.RemoveCurrent();
            mES2_MODEL_MASTERTableAdapter.Update(iFRYDataSet.MES2_MODEL_MASTER);
            MessageBox.Show("자료가 삭제 되었습니다.");
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

    }
}