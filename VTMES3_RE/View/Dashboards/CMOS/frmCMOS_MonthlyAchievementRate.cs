using DevExpress.DashboardCommon;
using DevExpress.DashboardCommon.Native;
using DevExpress.DashboardWin;
using DevExpress.DataAccess.ConnectionParameters;
using DevExpress.XtraEditors;
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
using System.Xml.Linq;
using VTMES3_RE.Common;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.Dashboards.CMOS
{
    // 테스트 폼 삭제 가능
    public partial class frmCMOS_MonthlyAchievementRate : DevExpress.XtraEditors.XtraForm
    {
        clsCode code = new clsCode();
        DataRowView master = null;
        string menuId = "";

        public frmCMOS_MonthlyAchievementRate()
        {
            menuId = WrGlobal.OpeningMenuId;

            InitializeComponent();

            startDateEdit.EditValue = DateTime.Today.AddDays(-1);
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

        private void frmCMOS_MonthlyAchievementRate_Load(object sender, EventArgs e)
        {
            //ResetParameter();
        }

        private void ResetParameter()
        {
            if (dashboardViewer.Parameters == null) return;

            DashboardParameters parameters1 = dashboardViewer.Parameters;

            //dashboardViewer.BeginUpdateParameters();
            for (int i = 0; i < parameters1.Count; i++)
            {
                if (parameters1[i].Type == ParameterValueType.DateTime)
                {
                    
                    if (parameters1[i].Name.ToLower().IndexOf("start") >= 0 || parameters1[i].Name.ToLower().IndexOf("from") >= 0)
                    {
                        parameters1[i].SelectedValue = startDateEdit.EditValue;
                    }
                    else if (parameters1[i].Name.ToLower().IndexOf("end") >= 0 || parameters1[i].Name.ToLower().IndexOf("to") >= 0)
                    {
                        parameters1[i].SelectedValue = endDateEdit.EditValue;
                    }
                    else
                    {
                        parameters1[i].SelectedValue = DateTime.Today;
                    }
                }
            }
            //dashboardViewer.EndUpdateParameters();
        }
        private void btnSearch_Click(object sender, EventArgs e)
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

        private void dashboardViewer_CustomParameters(object sender, DevExpress.DashboardCommon.CustomParametersEventArgs e)
        {
            foreach (DashboardParameter param in e.Parameters)
            {
                if (param.Type.Name == "DateTime")
                {
                    if (param.Name.ToLower().IndexOf("start") >= 0 || param.Name.ToLower().IndexOf("from") >= 0)
                    {
                        param.Value = startDateEdit.EditValue;
                    }
                    else if (param.Name.ToLower().IndexOf("end") >= 0 || param.Name.ToLower().IndexOf("to") >= 0)
                    {
                        param.Value = endDateEdit.EditValue;
                    }
                    else
                    {
                        param.Value = DateTime.Today;
                    }
                }
            }
        }
    }
}