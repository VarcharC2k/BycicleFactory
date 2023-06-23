using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.XtraEditors;
using DevExpress.XtraTreeList.Nodes;
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
using VTMES3_RE.Common;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.Dashboards.Tools
{
    // 사용자별로 관리하는 대시보드 폼
    public partial class frmUserDashboardView : DevExpress.XtraEditors.XtraForm
    {
        clsCode code = new clsCode();
        DataRowView master = null;

        string menuId = "";

        public frmUserDashboardView()
        {
            InitializeComponent();

            menuId = WrGlobal.OpeningMenuId;

            // menuid로 대시보드 정보 가져오기
            master = code.IsExistDashboardItem(menuId);

            if (master["XML"].ToString() == "")
            {
                MessageBox.Show("선택된 메뉴는 대시보드가 작성되지 않았습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 대시보드 XML로 로드
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

        private void frmUserDashboardView_Load(object sender, EventArgs e)
        {

        }

        private void dashboardViewer_ConfigureDataConnection(object sender, DevExpress.DashboardCommon.DashboardConfigureDataConnectionEventArgs e)
        {
            if (e.ConnectionParameters.GetType().Name != "MsSqlConnectionParameters") return;
            MsSqlConnectionParameters parameters = (MsSqlConnectionParameters)e.ConnectionParameters;

            parameters.UserName = WrGlobal.DBUserName;
            parameters.Password = WrGlobal.DBUserPassword;
        }
        // 시작일, 종료일 파라메타 값 자동 설정
        private void dashboardViewer_DashboardLoaded(object sender, DevExpress.DashboardWin.DashboardLoadedEventArgs e)
        {
            int cnt = e.Dashboard.Parameters.Count;

            for (int i = 0; i < (cnt); i++)
            {
                if (e.Dashboard.Parameters[i].Type.Name == "DateTime")
                {
                    if (e.Dashboard.Parameters[i].Name.ToLower().IndexOf("start") >= 0 || e.Dashboard.Parameters[i].Name.ToLower().IndexOf("from") >= 0)
                    {
                        e.Dashboard.Parameters[i].Value = DateTime.Today.AddDays(-1);
                    }
                    else if (e.Dashboard.Parameters[i].Name.ToLower().IndexOf("end") >= 0 || e.Dashboard.Parameters[i].Name.ToLower().IndexOf("to") >= 0)
                    {
                        e.Dashboard.Parameters[i].Value = DateTime.Today;
                    }
                    else
                    {
                        e.Dashboard.Parameters[i].Value = DateTime.Today;
                    }
                }

            }
        }

        private void frmUserDashboardView_Shown(object sender, EventArgs e)
        {
            dashboardViewer.ShowDashboardParametersForm();
        }
        // 대시보드 작성 폼 오픈
        private void cmdDashboardDesign_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            DataRowView drv = code.IsExistDashboardItem(menuId);

            if (drv == null)
            {
                MessageBox.Show("신규 등록된 메뉴는 저장 후 작성하세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            frmDashBoardDesign form = new frmDashBoardDesign(drv);
            form.OnItemSaveChanged = new frmDashBoardDesign.ItemSaveChanged(ReloadDashboard);
            form.ShowDialog();


        }
        // 대시보드 리로드
        private void ReloadDashboard(string xml)
        {
            MemoryStream ms = new MemoryStream();
            byte[] m_Buffer;

            m_Buffer = System.Text.Encoding.UTF8.GetBytes(xml);
            ms.Write(m_Buffer, 0, m_Buffer.Length);
            ms.Seek(0, SeekOrigin.Begin);
            dashboardViewer.LoadDashboard(ms);
            ms.Flush();
            ms.Close();
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
    }
}