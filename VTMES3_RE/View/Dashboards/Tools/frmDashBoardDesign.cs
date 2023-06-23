using DevExpress.XtraBars;
using VTMES3_RE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.DashboardWin;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using System.IO;
using DevExpress.DashboardCommon;
using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.DataAccess.Sql;
using DevExpress.DashboardCommon.Native;
using DevExpress.DataAccess.Native;

namespace VTMES3_RE.View.Dashboards.Tools
{
    // 대시보드 디자인 폼
    public partial class frmDashBoardDesign : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        
        Database db = new Database();
        string query = "";
        DataRowView master = null;

        public delegate void ItemSaveChanged(string xml);
        public ItemSaveChanged OnItemSaveChanged;

        DateTime startDate = Convert.ToDateTime(DateTime.Now.Year.ToString() + "-01-01");
        DateTime endDate = Convert.ToDateTime(DateTime.Now.Year.ToString() + "-12-31");

        public frmDashBoardDesign()
        {
            InitializeComponent();
        }

        public frmDashBoardDesign(DataRowView _item)
        {
            InitializeComponent();

            master = _item;

            this.Text = master["ParentMenuName"].ToString() + " - " + master["MenuName"].ToString();

            // 대시보드 디자이너 리본 설정 -> 홈 그룹 메뉴 비활성화
            dashboardDesigner.CreateRibbon();

            RibbonControl ribbon = dashboardDesigner.Ribbon;
            RibbonPage homeRibbonPage = ribbon.GetDashboardRibbonPage(DashboardBarItemCategory.None, DashboardRibbonPage.Home);
            RibbonPageGroup fileRibbonPageGroup = homeRibbonPage.Groups[0];
            fileRibbonPageGroup.Enabled = false;
            fileRibbonPageGroup.Visible = false;
            ribbon.Toolbar.ItemLinks.RemoveAt(0);
            Control backstageViewControl = ribbon.ApplicationButtonDropDownControl as Control;
            if (backstageViewControl != null)
                backstageViewControl.Enabled = false;

            // 저장된 XML 이 없으면 신규 세팅 -> 파라메타(시작일, 종료일), CAMDB 데이터 소스 생성
            if (master["XML"].ToString() == "")
            {
                DashboardParameter parameter1 = new DashboardParameter("startDate", typeof(DateTime), startDate);
                dashboardDesigner.Dashboard.Parameters.Add(parameter1);
                DashboardParameter parameter2 = new DashboardParameter("endDate", typeof(DateTime), endDate);
                dashboardDesigner.Dashboard.Parameters.Add(parameter2);

                MsSqlConnectionParameters sqlParams = new MsSqlConnectionParameters();
                sqlParams.AuthorizationType = MsSqlAuthorizationType.SqlServer;
                sqlParams.ServerName = WrGlobal.DBServer;
                sqlParams.DatabaseName = "RYCAMDB";

                DashboardSqlDataSource sqlDataSource = new DashboardSqlDataSource("CAMDB Data Source", sqlParams);
                sqlDataSource.ConnectionOptions.DbCommandTimeout = 300;
                dashboardDesigner.Dashboard.DataSources.Add(sqlDataSource);

                dashboardDesigner.Dashboard.Title.Text = master["MenuName"].ToString();
            }
            else
            {   // 기존 작성된 대시보드 -> XML 로드
                master["XML"] = master["XML"].ToString().Replace("?<?", "<?");
                MemoryStream ms = new MemoryStream();
                byte[] m_Buffer;

                m_Buffer = System.Text.Encoding.UTF8.GetBytes(master["XML"].ToString());
                ms.Write(m_Buffer, 0, m_Buffer.Length);
                ms.Seek(0, SeekOrigin.Begin);

                dashboardDesigner.LoadDashboard(ms);
                ms.Flush();
                ms.Close();

                //dashboardDesigner.Dashboard.Parameters[0].Value = startDate;
                //dashboardDesigner.Dashboard.Parameters[1].Value = endDate;
            }


            // Creates a new dashboard parameter.
            //StaticListLookUpSettings staticSettings = new StaticListLookUpSettings();
            //staticSettings.Values = new string[] { "1994", "1995", "1996" };
            //DashboardParameter yearParameter = new DashboardParameter("yearParameter",
            //    typeof(string), "1995", "Select year:", true, staticSettings);
            //dashboardDesigner.Parameters.Add(yearParameter);
             
            //DashboardSqlDataSource dataSource = (DashboardSqlDataSource)dashboard.DataSources[0];
            //CustomSqlQuery salesPersonQuery = (CustomSqlQuery)dataSource.Queries[0];
            //salesPersonQuery.Parameters.Add(new QueryParameter("startDate", typeof(Expression),
            //    new Expression("[Parameters.yearParameter] + '/01/01'")));
            //salesPersonQuery.Parameters.Add(new QueryParameter("endDate", typeof(Expression),
            //    new Expression("[Parameters.yearParameter] + '/12/31'")));
            //salesPersonQuery.Sql =
            //    "select * from SalesPerson where OrderDate between @startDate and @endDate";

            //dashboardViewer1.Dashboard = dashboard;
        }
        private void frmDashBoardDesign_Load(object sender, EventArgs e)
        {

        }
        // 디자이너 닫을때 변경된 내용 XML로 DB 저장
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            if (dashboardDesigner.IsDashboardModified)
            {
                DialogResult result = XtraMessageBox.Show(LookAndFeel, this, "변경 내용을 저장하겠습니까 ?", "대시보드 디자이너",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result == DialogResult.Cancel)
                    e.Cancel = true;
                else if (result == DialogResult.Yes)
                {
                    MemoryStream ms = new MemoryStream();
                    byte[] m_Buffer;
                    string m_XML = "";

                    dashboardDesigner.Dashboard.SaveToXml(ms);
                    m_Buffer = ms.ToArray();
                    ms.Flush();
                    ms.Close();
                    m_XML = System.Text.Encoding.UTF8.GetString(m_Buffer);
                    m_XML = m_XML.Replace("?<?", "<?");
                    m_XML = m_XML.Replace("'", "''");
                    master["XML"] = m_XML;

                    query = String.Format("Update {0}_ReportDB.dbo.DashBoardItem Set XML = N'{2}', ModId = '{3}', ModIP = '{4}', ModDt = getdate() Where CorpId = '{0}' and MenuId = '{1}'",
                                WrGlobal.CorpID, master["MenuId"], master["XML"], WrGlobal.LoginID, WrGlobal.ClientHostName);
                    db.ExecuteQuery(query);
                    //e.Cancel = true;

                    if (OnItemSaveChanged != null)
                    {
                        OnItemSaveChanged(master["XML"].ToString());
                    }  
                }
            }
        }

        private void dashboardDesigner_DashboardSaving(object sender, DashboardSavingEventArgs e)
        {
            e.Handled = true;
            e.Saved = false;
        }

        private void dashboardDesigner_DashboardClosing(object sender, DashboardClosingEventArgs e)
        {
            e.IsDashboardModified = false;
        }

        // Data source 연결에 대한 DB ID/PW 설정
        private void dashboardDesigner_ConfigureDataConnection(object sender, DashboardConfigureDataConnectionEventArgs e)
        {
            if (e.ConnectionParameters.GetType().Name != "MsSqlConnectionParameters") return;

            MsSqlConnectionParameters parameters = (MsSqlConnectionParameters)e.ConnectionParameters;
           
            parameters.UserName = WrGlobal.DBUserName;
            parameters.Password = WrGlobal.DBUserPassword;
        }

    }
}