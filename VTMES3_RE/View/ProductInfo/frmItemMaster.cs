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
using VTMES3_RE.Models;

namespace VTMES3_RE.View.ProductInfo
{
    // 김성재 주임 작성
    // 김문철 부장 -> 런시트공정, 성적서 공정 컬럼 추가
    public partial class frmItemMaster : DevExpress.XtraEditors.XtraForm
    {
        clsWork work = new clsWork();

        public frmItemMaster()
        {
            InitializeComponent();
        }

        private void frmItemMaster_Load(object sender, EventArgs e)
        {
            work.MES2_ITEM_MASTER_insert();

            DisplayData();
        }

        private void cmdSearch_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            DisplayData();
        }

        private void DisplayData()
        {
            this.mES2_ITEM_MASTERTableAdapter.Fill(this.iFRYDataSet.MES2_ITEM_MASTER);

            gvITEM_MASTER.BestFitColumns();
        }
    }
}