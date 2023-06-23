using VTMES3_RE.Common;
using VTMES3_RE.Models;
using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using static DevExpress.XtraEditors.Mask.MaskSettings;

namespace VTMES3_RE
{
    public partial class frmLogin : DevExpress.XtraEditors.XtraForm
    {
        // 이전 로그인 ID 파일 저장
        private string InitFile = Application.StartupPath + @"\login.ini";

        [DllImport("kernel32")]  //C,C++로 개발한 Native DLL파일을 호출한다.
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]  //C,C++로 개발한 Native DLL파일을 호출한다.
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        clsCode code = new clsCode();   

        public frmLogin()
        {
            InitializeComponent();

            // Camstar Employee 테이블 바인딩
            searchEmpNo.Properties.DataSource = code.GetUser();
            searchEmpNo.Properties.DisplayMember = "사용자명";
            searchEmpNo.Properties.ValueMember = "ID";

            if (System.IO.File.Exists(InitFile))
            {
                searchEmpNo.EditValue = IniReadValue("LOGIN", "LoginID").Trim();
            }

        }
        //Login폼이 로드할때 시작되는 로직.
        private void frmLogin_Load(object sender, EventArgs e)
        {
            if ((searchEmpNo.EditValue ?? "").ToString() != "")
            {
                searchEmpNo.Focus();
            }
        }

        // 로그인
        private void btnLogin_Click(object sender, EventArgs e)
        {
            SetLogin();
        }

        /// <summary>
        /// 로그인 사용자 인증 및 사용자 정보 설정
        /// </summary>
        private void SetLogin()
        {
            if ((searchEmpNo.EditValue ?? "").ToString() == "")
            {
                MessageBox.Show("사용자를 선택하세요", "로그인 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                searchEmpNo.Focus();
                return;
            }
            if (txtPW.Text.Equals(""))
            {
                MessageBox.Show("Password를 입력하세요", "로그인 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPW.Focus();
                return;
            }

            // 사용자 정보 가져오기
            DataRowView dr = code.GetUserInfo((searchEmpNo.EditValue ?? "").ToString());

            if (dr == null)
            {
                dr = code.AddUser((searchEmpNo.EditValue ?? "").ToString(), txtPW.Text);
            }

            // 패스워드 인증
            if (clsCommon.SHA256Hash(txtPW.Text) == clsCommon.getString(dr["Password"]))
            {
                WrGlobal.EmployeeId = (dr["EmployeeId"] ?? "").ToString();
                WrGlobal.LoginID = (dr["EmployeeName"] ?? "").ToString();
                WrGlobal.LoginUserNM = (dr["FullName"] ?? "").ToString();

                WrGlobal.FactoryId = (dr["FactoryId"] ?? "").ToString();
                WrGlobal.FactoryName = (dr["FactoryName"] ?? "").ToString();
                WrGlobal.Camstar_RoleName = (dr["ESigRoleGroupName"] ?? "").ToString();

                //팀 설정
                code.SetEmployeeTeam();
                //권한 세팅
                code.SetUserAuthor();
                //세션 설정
                code.CreateSession();
            }
            else
            {
                MessageBox.Show("비밀번호가 일치 하지 않습니다. 다시 입력하세요");
                txtPW.Text = "";
                txtPW.Focus();
                return;
            }

            // 로그인 ID 파일 저장
            WritePrivateProfileString("LOGIN", "LoginID", (searchEmpNo.EditValue ?? "").ToString(), InitFile);

            this.Close();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        /// <summary>
        /// 이전 로그인 정보 가져오기
        /// </summary>
        /// <param name="Section">항목명</param>
        /// <param name="Key">키명</param>
        /// <returns></returns>
        private string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, InitFile);
            return temp.ToString();
        }

        private void txtPW_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Return))
            {
                btnLogin.PerformClick();
            }
        }

    }
}