using DevExpress.XtraBars.Navigation;
using DevExpress.XtraSplashScreen;
using VTMES3_RE.Common;
using VTMES3_RE.Models;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using DevExpress.Skins;
using DevExpress.Skins.XtraForm;
using DevExpress.Utils;
using VTMES3_RE.Properties;
using static DevExpress.XtraEditors.Mask.MaskSettings;

namespace VTMES3_RE
{
    public partial class frmMain : DevExpress.XtraEditors.XtraForm
    {
        // 기초 코드 모델 생성
        clsCode code = new clsCode();
        public frmMain()
        {
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Location = System.Windows.Forms.Screen.GetBounds(MousePosition).Location;

            InitializeComponent();

            Icon = VTMES3_RE.Properties.Resources.rayence_icon;
            mainAccordion.Footer.Visible = false;

            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                this.Text = this.Text + " (" + System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() + ")";
            }

            setUserAccordionMenu();
        }

        /// <summary>
        /// 사용자 메뉴 설정
        /// </summary>
        private void setUserAccordionMenu()
        {
            //DevExpress.XtraBars.Navigation.AccordionControlElement parentGroup = mainAccordion.Elements.add

            DevExpress.XtraBars.Navigation.AccordionControlElement rootGroup = null;
            DevExpress.XtraBars.Navigation.AccordionControlElement parentGroup1 = null;
            DevExpress.XtraBars.Navigation.AccordionControlElement parentGroup2 = null;
            DevExpress.XtraBars.Navigation.AccordionControlElement parentGroup3 = null;
            DevExpress.XtraBars.Navigation.AccordionControlElement addGroup = null;
            DevExpress.XtraBars.Navigation.AccordionControlElement eleItem = null;

            // 전체 사용자 메뉴 가져오기
            DataView dv = code.GetEmployeeMenuList();   

            foreach (DataRowView drv in dv)
            {
                //if (drv["GroupLevel"].ToString() == "0")
                //{
                //    if (!WrGlobal.AuthorList.Exists(s => s.Substring(0, 2) == drv["GroupCode1"].ToString())) continue;
                //}
                //else if (drv["GroupLevel"].ToString() == "1")
                //{
                //    if (!WrGlobal.AuthorList.Exists(s => s.Substring(0, 4) == drv["GroupCode1"].ToString() + drv["GroupCode2"].ToString())) continue;
                //}
                //else if (drv["GroupLevel"].ToString() == "2")
                //{
                //    if (!WrGlobal.AuthorList.Exists(s => s.Substring(0, 6) == drv["GroupCode1"].ToString() + drv["GroupCode2"].ToString() + drv["GroupCode3"].ToString())) continue;
                //}
                //else if (drv["GroupLevel"].ToString() == "3")
                //{
                //    if (!WrGlobal.AuthorList.Exists(s => s.Substring(0, 6) == drv["GroupCode1"].ToString() + drv["GroupCode2"].ToString() + drv["GroupCode3"].ToString() + drv["GroupCode4"].ToString())) continue;
                //}
                //else if (drv["GroupLevel"].ToString() == "4")
                //{
                //    if (!WrGlobal.AuthorList.Exists(s => s.Substring(0, 6) == drv["GroupCode1"].ToString() + drv["GroupCode2"].ToString() + drv["GroupCode3"].ToString() + drv["GroupCode4"].ToString() + drv["GroupCode5"].ToString())) continue;
                //}

                // 루트 그룹 생성
                if (drv["GroupLevel"].ToString() == "0")
                {
                    rootGroup = mainAccordion.Elements.Add();
                    rootGroup.Name = drv["ID"].ToString();
                    rootGroup.Text = drv["Title_" + WrGlobal.Language].ToString();
                    rootGroup.ImageOptions.SvgImage = svgImages[drv["ImageKey"].ToString()];
                    rootGroup.ImageOptions.SvgImageSize = new Size(24, 24);
                    rootGroup.Appearance.Normal.ForeColor = Color.LightCyan;
                }
                else
                {
                    // 하위 그룹 생성
                    if (drv["GroupYn"].ToString() == "1")
                    {   // 그룹
                        addGroup = new DevExpress.XtraBars.Navigation.AccordionControlElement(DevExpress.XtraBars.Navigation.ElementStyle.Group);
                        addGroup.Name = drv["ID"].ToString();
                        addGroup.Text = drv["Title_" + WrGlobal.Language].ToString();

                        if (drv["ImageKey"].ToString() != "")
                        {
                            addGroup.ImageOptions.SvgImage = svgImages[drv["ImageKey"].ToString()];
                            addGroup.ImageOptions.SvgImageSize = new Size(24, 24);
                            addGroup.Appearance.Normal.ForeColor = Color.LightCyan;
                        }
                        addGroup.Expanded = drv["ExpandYn"].ToString() == "0" ? false : true;

                        if (drv["GroupLevel"].ToString() == "1")
                        {
                            rootGroup.Elements.Add(addGroup);
                            parentGroup1 = null;
                            parentGroup1 = addGroup;
                        }
                        else if (drv["GroupLevel"].ToString() == "2")
                        {
                            parentGroup1.Elements.Add(addGroup);
                            parentGroup2 = null;
                            parentGroup2 = addGroup;
                        }
                        else if (drv["GroupLevel"].ToString() == "3")
                        {
                            parentGroup2.Elements.Add(addGroup);
                            parentGroup3 = null;
                            parentGroup3 = addGroup;
                        }

                        addGroup = null;
                    }
                    else
                    {   // 메뉴 아이템 생성
                        eleItem = new DevExpress.XtraBars.Navigation.AccordionControlElement(DevExpress.XtraBars.Navigation.ElementStyle.Item);
                        eleItem.Name = drv["ID"].ToString();
                        eleItem.Text = drv["Title_" + WrGlobal.Language].ToString();
                        eleItem.Tag = drv["ProjectName"].ToString() + "|" + drv["FolderName"].ToString() + "|" + drv["FormName"].ToString();
                        if (drv["ImageKey"].ToString() != "")
                        {
                            eleItem.ImageOptions.SvgImage = svgImages[drv["ImageKey"].ToString()];
                            eleItem.ImageOptions.SvgImageSize = new Size(24, 24);
                        }

                        if (drv["GroupLevel"].ToString() == "1")
                        {
                            rootGroup.Elements.Add(eleItem);
                        }
                        else if (drv["GroupLevel"].ToString() == "2")
                        {
                            parentGroup1.Elements.Add(eleItem);
                        }
                        else if (drv["GroupLevel"].ToString() == "3")
                        {
                            parentGroup2.Elements.Add(eleItem);
                        }
                        else if (drv["GroupLevel"].ToString() == "4")
                        {
                            parentGroup3.Elements.Add(eleItem);
                        }

                        eleItem = null;
                    }
                }

            }

        }

        /// <summary>
        /// 메뉴 아이템 클릭 이벤트
        /// </summary>
        private void mainAccordion_ElementClick(object sender, ElementClickEventArgs e)
        {
            if (e.Element.Style == DevExpress.XtraBars.Navigation.ElementStyle.Group) return;
            if (e.Element.Tag == null) return;

            string[] arrTag = e.Element.Tag.ToString().Split(new char[] { '|' });

            if (!IsOpen(arrTag[2], e.Element.Name))
            {
                string temp = arrTag[0] + "." + arrTag[1] + "." + arrTag[2] + ",RYMES3";
                WrGlobal.OpeningMenuId = e.Element.Name;
                var frm = Activator.CreateInstance(Type.GetType(temp)) as Form;
                frm.Name = arrTag[2];
                frm.Text = e.Element.Text;
                frm.Tag = e.Element.Name;
                doOpenForm(frm);
            }
        }

        /// <summary>
        /// MDI 폼 오픈
        /// </summary>
        /// <param name="frm">오픈할 폼</param>
        private void doOpenForm(Form frm)
        {
            if (frm.GetType().Name == "frmPasswordChange")
            {
                Rectangle r = this.RectangleToScreen(this.Bounds);
                frm.Left = r.Left + (r.Width - this.Width) / 2;
                frm.Top = r.Top + (r.Height - this.Height) / 4;

                if (frm.ShowDialog() == DialogResult.OK)
                {

                }
            }
            else
            {
                code.UseSession((frm.Tag ?? "").ToString(), "Open");

                frm.MdiParent = this;
                frm.Closed += new EventHandler(MDIChildrenCleanup);
                frm.Dock = DockStyle.Fill;
                frm.WindowState = FormWindowState.Maximized;
                frm.Show();
                Application.DoEvents();
                frm.BringToFront();
                
            }

        }//end function

        /// <summary>
        /// 기존 오픈된 폼인지 확인
        /// </summary>
        /// <param name="frmName">폼명</param>
        /// <param name="tagName">태그명</param>
        /// <returns></returns>
        private bool IsOpen(string frmName, string tagName)
        {
            bool m_fReturn = false;
            foreach (var m_form in this.MdiChildren)
            {
                if (m_form.Name.Equals(frmName) && (m_form.Tag ?? "").ToString() == tagName)
                {
                    m_form.Focus();
                    m_fReturn = true;
                    break;
                }//end if

            }//end foreach
            return m_fReturn;
        }//end function
        // 닫을 폼 클리닝
        private void MDIChildrenCleanup(object sender, EventArgs e)
        {
            ((Form)sender).Dispose();
        }
        
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (WrGlobal.SessionNo != "")
            {
                code.CloseSession();
            }
        }
        // 패스워드 변경 폼 오픈
        private void btnChangePw_Click(object sender, EventArgs e)
        {
            var frm = Activator.CreateInstance(Type.GetType(string.Format("{0}.View.Systems.frmPasswordChange", WrGlobal.ProJectName)), WrGlobal.LoginID) as Form;
            doOpenForm(frm);
        }
        // 로그 아웃 처리
        private void btnLogout_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(string.Format("사용자({0}) : 로그아웃 하겠습니까?", WrGlobal.LoginID), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.OK)
            {
                Application.Restart();
            }//end if
            
        }
        // 폼 상단바 세팅
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Skin currentSkin = FormSkins.GetSkin(this.LookAndFeel);
            SkinElement element = currentSkin[FormSkins.SkinFormCaption];
            element.ContentMargins.Top = 6;
            element.ContentMargins.Bottom = 6;

            //element.Color.SolidImageCenterColor = Color.FromArgb(255, 33, 80, 126);
            this.LookAndFeel.UpdateStyleSettings();
        }
        // 폼 상단바 세팅
        protected override DevExpress.Skins.XtraForm.FormPainter CreateFormBorderPainter()
        {
            return new MyFormPainter(this, LookAndFeel, Resources.rayence_icon, new Size(60, 20));
        }

        public class MyFormPainter : FormPainter
        {
            private readonly Icon _icon;
            private readonly Size _size;

            public MyFormPainter(Control owner, ISkinProvider provider) : base(owner, provider) { }

            public MyFormPainter(Control owner, ISkinProvider provider, Icon icon, Size size) : base(owner, provider)
            {
                _icon = icon;
                _size = size;
            }

            protected override Size GetIconSize() { return _size; }
            protected override Icon GetIcon() { return _icon; }
            // 폼 상단바 텍스트 세팅
            protected override void DrawText(DevExpress.Utils.Drawing.GraphicsCache cache)
            {
                string text = Text;
                if (text == null || text.Length == 0 || this.TextBounds.IsEmpty) return;
                AppearanceObject appearance = new AppearanceObject(GetDefaultAppearance());
                appearance.Font = new Font("Segoe UI", 12, FontStyle.Bold);
                appearance.TextOptions.Trimming = Trimming.EllipsisCharacter;
                appearance.ForeColor = Color.FromArgb(0, 100, 180);
                Rectangle r = RectangleHelper.GetCenterBounds(TextBounds, new Size(TextBounds.Width, appearance.CalcDefaultTextSize(cache.Graphics).Height));
                DrawTextShadow(cache, appearance, r);
                cache.DrawString(text, appearance.Font, appearance.GetForeBrush(cache), r, appearance.GetStringFormat());
            }
        }
    }
}