using DevExpress.DashboardCommon.Native;
using DevExpress.DashboardWin;
using DevExpress.Data.Svg;
using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.XtraEditors;
using DevExpress.XtraRichEdit.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.Common;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.Dashboards.CMOS
{ 
    // 테스트 폼 삭제 가능
    public partial class frmCMOS_PerfomanceAgainstPlan : DevExpress.XtraEditors.XtraForm
    {
        clsCode code = new clsCode();
        DataRowView master = null;

        string menuId = "";

        public frmCMOS_PerfomanceAgainstPlan()
        {
            menuId = WrGlobal.OpeningMenuId;

            InitializeComponent();

            startDateEdit.EditValue = Convert.ToDateTime(DateTime.Today.Year.ToString() + "-01-01");
            endDateEdit.EditValue = DateTime.Today;

            master = code.IsExistDashboardItem(menuId);


            if (master["XML"].ToString() == "")
            {
                MessageBox.Show("선택된 메뉴는 대시보드가 작성되지 않았습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            master["XML"] = master["XML"].ToString().Replace("?<?", "<?");
            MemoryStream ms = new MemoryStream();
            byte[] m_Buffer;

            m_Buffer = System.Text.Encoding.UTF8.GetBytes(master["XML"].ToString());
            ms.Write(m_Buffer, 0, m_Buffer.Length);
            ms.Seek(0, SeekOrigin.Begin);
            dashboardViewer.LoadDashboard(ms);
            ms.Flush();
            ms.Close();
        }

        private void frmCMOS_PerfomanceAgainstPlan_Load(object sender, EventArgs e)
        {
            ResetParameter();
        }

        private void dashboardViewer_ConfigureDataConnection(object sender, DevExpress.DashboardCommon.DashboardConfigureDataConnectionEventArgs e)
        {
            if (e.ConnectionParameters.GetType().Name != "MsSqlConnectionParameters") return;
            MsSqlConnectionParameters parameters = (MsSqlConnectionParameters)e.ConnectionParameters;

            parameters.UserName = WrGlobal.DBUserName;
            parameters.Password = WrGlobal.DBUserPassword;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            ResetParameter();
        }

        private void ResetParameter()
        {
            DashboardParameters parameters1 = dashboardViewer.Parameters;

            dashboardViewer.BeginUpdateParameters();
            parameters1[0].SelectedValue = Convert.ToDateTime(startDateEdit.EditValue);
            parameters1[1].SelectedValue = Convert.ToDateTime(endDateEdit.EditValue);
            dashboardViewer.EndUpdateParameters();
        }

    }
}