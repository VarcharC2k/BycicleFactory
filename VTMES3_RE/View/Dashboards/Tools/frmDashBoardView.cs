using DevExpress.DashboardCommon;
using DevExpress.DashboardWin;
using DevExpress.DataAccess.ConnectionParameters;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using VTMES3_RE.Common;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.Dashboards.Tools
{
    public partial class frmDashBoardView : DevExpress.XtraEditors.XtraForm
    {
        clsCode code = new clsCode();
        DataRowView master = null;

        string menuId = "";

        public frmDashBoardView()
        {
            InitializeComponent();

            menuId = WrGlobal.OpeningMenuId;

            // menuid로 대시보드 정보 가져오기
            master = code.IsExistDashboardItem(menuId);

            if (master["XML"].ToString() == "")
            {
                MessageBox.Show("선택된 메뉴는 대시보드가 작성되지 않았습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
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
        
        // DashboardItem 의 로우를 전달받아 대시보드 로드
        public frmDashBoardView(DataRowView _drv)
        {
            InitializeComponent();

            menuId = _drv["MenuId"].ToString();

            master = _drv;

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

        private void frmDashBoardView_Load(object sender, EventArgs e)
        {

        }
        // 대시보드 Datasource 연결 정보 설정 
        private void dashboardViewer_ConfigureDataConnection(object sender, DevExpress.DashboardCommon.DashboardConfigureDataConnectionEventArgs e)
        {
            if (e.ConnectionParameters.GetType().Name != "MsSqlConnectionParameters") return;
            MsSqlConnectionParameters parameters = (MsSqlConnectionParameters)e.ConnectionParameters;

            parameters.UserName = WrGlobal.DBUserName;
            parameters.Password = WrGlobal.DBUserPassword;
        }

        // 시작일, 종료일 파라메타 값 자동 설정
        private void dashboardViewer_DashboardLoaded(object sender, DashboardLoadedEventArgs e)
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

        private void frmDashBoardView_Shown(object sender, EventArgs e)
        {
            dashboardViewer.ShowDashboardParametersForm();
        }
    }
}