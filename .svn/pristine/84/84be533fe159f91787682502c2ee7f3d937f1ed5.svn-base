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
using VTMES3_RE.IFRYDataSetTableAdapters;

namespace VTMES3_RE.View.ProductInfo
{
    // 김성재 주임 작성

    public partial class frmProductST : DevExpress.XtraEditors.XtraForm
    {
        public frmProductST()
        {
            InitializeComponent();
        }

        private void frmProductST_Load(object sender, EventArgs e)
        {
            this.mES2_ST_MASTERTableAdapter.FillByList(this.iFRYDataSet.MES2_ST_MASTER);
            gvProductST.BestFitColumns();
        }

        private void cmdAdd_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            mES2_ST_MASTERBindingSource.AddNew();
            DataRowView drv = (DataRowView)mES2_ST_MASTERBindingSource.Current;
        }


        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();
                this.mES2_ST_MASTERBindingSource.EndEdit();
                this.mES2_ST_MASTERTableAdapter.Update(this.iFRYDataSet.MES2_ST_MASTER);
                MessageBox.Show("모델정보를 저장했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            string Id = gvProductST.GetFocusedRowCellDisplayText("ITEM_CODE") == null ? "" : gvProductST.GetFocusedRowCellDisplayText("ITEM_CODE");

            if (MessageBox.Show(string.Format("선택한 '{0}' 모델을 삭제하시겠습니까?", Id), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }//end if

            mES2_ST_MASTERBindingSource.RemoveCurrent();
            mES2_ST_MASTERTableAdapter.Update(iFRYDataSet.MES2_ST_MASTER);
            MessageBox.Show("모델정보가 삭제 되었습니다.");
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

        private void gvProductST_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            DataRowView drv = (DataRowView)mES2_ST_MASTERBindingSource.Current;
            if (drv.Row.RowState == DataRowState.Detached || drv.Row.RowState == DataRowState.Added)
            {
                colITEM_CODE2.OptionsColumn.AllowEdit = true;
                colITEM_NAME.OptionsColumn.AllowEdit = true;
                colFA_ID2.OptionsColumn.AllowEdit = true;
                colOP_ID2.OptionsColumn.AllowEdit = true;
            }
            else
            {
                colITEM_CODE2.OptionsColumn.AllowEdit = false;
                colITEM_NAME.OptionsColumn.AllowEdit = false;
                colFA_ID2.OptionsColumn.AllowEdit = false;
                colOP_ID2.OptionsColumn.AllowEdit = false;
            }
        }
    }
}