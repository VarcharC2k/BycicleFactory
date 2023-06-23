using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using VTMES3_RE.Common;
using VTMES3_RE.View.Reports.Tools;
using DevExpress.XtraReports.Extensions;

namespace VTMES3_RE
{
    internal static class Program
    {
        public static ReportStorageExtension reportStorage;
        public static ReportStorageExtension ReportStorage
        {
            get
            {
                return reportStorage;
            }
        }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]  //COM객체는 STAThread 기반이어서 이를 명시해줘야 한다.
        static void Main()
        {
            DevExpress.DashboardCommon.Localization.DashboardLocalizer.Active = new clsDashboardCommonLocalizer();
            DevExpress.DashboardWin.Localization.DashboardWinLocalizer.Active = new clsDashboardWinLocalizer();

            InstalledFontCollection installedFontCollection = new InstalledFontCollection();
            bool isFontInstall = false;
            foreach (FontFamily fontFamily in installedFontCollection.Families)
            {
                if (fontFamily.Name.Equals("Pretendard SemiBold"))
                {
                    isFontInstall = true;
                }
            }

            // 폰트 설치 확인 및 설치
            if (!isFontInstall)
            {
                Shell32.Shell shell = new Shell32.Shell();
                Shell32.Folder fontFolder = shell.NameSpace(0x14);
                fontFolder.CopyHere(Application.StartupPath + @"\Fonts\Pretendard-SemiBold.ttf");
            }
 
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 사업장 코드 설정
            WrGlobal.CorpID = "RY";

            // Camstar Tran DB 정보 
            WrGlobal.Camstar_SQL_SERVER = "RY-MESDB-SVR01";
            WrGlobal.Camstar_SQL_Database = "RYCAMDB";
            WrGlobal.Camstar_SQL_Id = "sa";
            WrGlobal.Camstar_SQL_Password = "Dentalimageno.1";

            // Camstar API 접속 정보
            WrGlobal.Camstar_Host = "RY-MESAPP-SVR01";
            WrGlobal.Camstar_Port = 443;
            WrGlobal.Camstar_UserName = "Administrator";
            WrGlobal.Camstar_Password = "Dentalimageno.1";

            // 웹 레포트 연결 기본 URL
            WrGlobal.reportRootUrl = @"http://RY-MESAPP-SVR01/RYReport";

            // 레포트 스토리지 설정
            reportStorage = new DataSetReportStorage();
            ReportStorageExtension.RegisterExtensionGlobal(reportStorage);

            if (new frmLogin().ShowDialog() == DialogResult.OK)
            {
                Application.Run(new frmMain());
            }

            //Application.Run(new View.Systems.frmUserView());
            //Application.Run(new View.WorkManager.frmEmployeeWorkTime());
            //Application.Run(new View.CamstarInf.frmDepoHistory());
            //Application.Run(new View.CamstarInf.frmExecTaskByResource());
            //Application.Run(new View.WorkManager.frm_Wafer_His_CSI_PopUp());
            //Application.Run(new View.CamstarInf.frmExecTaskByCmosLot());
            //Application.Run(new View.CamstarInf.frmMapingSnWn());
            //Application.Run(new View.WorkManager.ProductPlan.frmProductionPlan());
            //Application.Run(new View.WorkManager.ProductPlan.frmProductionMothlyPlan());
        }
    }
}
