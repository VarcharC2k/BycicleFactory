using DevExpress.DashboardCommon.DataProcessing;
using DevExpress.XtraBars;
using DevExpress.XtraReports;
using DevExpress.XtraReports.Design;
using DevExpress.XtraReports.UI;
using DevExpress.XtraReports.UserDesigner;
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

namespace VTMES3_RE.View.Reports.Tools
{
    // 레포트 작성 디자이너 폼
    public partial class frmReportDesign : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Database db = new Database();
        public string query = "";
        DataRowView master = null; 

        public frmReportDesign()
        {
            InitializeComponent();
        }

        // _item : ReportItem 테이블의 Row 
        public frmReportDesign(DataRowView _item)
        {
            InitializeComponent();

            master = _item;

            XtraReport report = new XtraReport();
            // LayoutData (XML) 로 레포트 양식 로드 
            if (master["LayoutData"].ToString() != "")
            {
                using (StreamWriter sw = new StreamWriter(new MemoryStream()))
                {
                    sw.Write(master["LayoutData"].ToString().Replace("?<?", "<?"));
                    sw.Flush();
                    report = XtraReport.FromStream(sw.BaseStream, true);
                }
            }
            reportDesigner.OpenReport(report);
        }

        private void frmReportDesign_Load(object sender, EventArgs e)
        {
       
        }
        // 레포트 작성 및 수정후 저장 이벤트 핸들러
        public class SaveCommandHandler : DevExpress.XtraReports.UserDesigner.ICommandHandler
        {
            XRDesignPanel panel;
            string menuId = "";

            public SaveCommandHandler(XRDesignPanel panel, string menuId)
            {
                this.panel = panel;
                this.menuId = menuId;
            }

            public void HandleCommand(DevExpress.XtraReports.UserDesigner.ReportCommand command,
            object[] args)
            {
                Save();
            }

            public bool CanHandleCommand(DevExpress.XtraReports.UserDesigner.ReportCommand command,
            ref bool useNextHandler)
            {
                 useNextHandler = !(command == ReportCommand.SaveFile ||
                    command == ReportCommand.SaveFileAs);
                return !useNextHandler;
            }

            // ReportItem 테이블 Update
            void Save()
            {
                MemoryStream stream = new MemoryStream();
                panel.Report.SaveLayout(stream);
                panel.ReportState = ReportState.Saved;

                stream.Position = 0;

                using (StreamReader sr = new StreamReader(stream))
                {
                    string s = sr.ReadToEnd();

                    Database repDb = new Database();
                    string query = String.Format("Update {0}_ReportDB.dbo.ReprotItem Set Report = N'{2}', ModId = '{3}', ModIP = '{4}', ModDt = getdate() Where CorpId = '{0}' and ReportId = '{1}'",
                    WrGlobal.CorpID, this.menuId, s, WrGlobal.LoginID, WrGlobal.ClientHostName);
                    repDb.ExecuteQuery(query);
                }       
            }
        }

        private void reportDesigner_DesignPanelLoaded(object sender, DesignerLoadedEventArgs e)
        {
            XRDesignPanel panel = (XRDesignPanel)sender;
            panel.AddCommandHandler(new SaveCommandHandler(panel, master["MenuId"].ToString()));
        }
    }
}