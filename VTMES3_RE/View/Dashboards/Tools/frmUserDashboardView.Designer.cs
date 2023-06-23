namespace VTMES3_RE.View.Dashboards.Tools
{
    partial class frmUserDashboardView
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmUserDashboardView));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.tileNavPane = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.navTitle = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdDisplay = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdDashboardDesign = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdClose = new DevExpress.XtraBars.Navigation.NavButton();
            this.layoutControlItem14 = new DevExpress.XtraLayout.LayoutControlItem();
            this.dashboardViewer = new DevExpress.DashboardWin.DashboardViewer(this.components);
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dashboardViewer)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.dashboardViewer);
            this.layoutControl1.Controls.Add(this.tileNavPane);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1051, 653);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem14,
            this.layoutControlItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1051, 653);
            this.Root.TextVisible = false;
            // 
            // tileNavPane
            // 
            this.tileNavPane.AllowGlyphSkinning = true;
            this.tileNavPane.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.tileNavPane.Appearance.Options.UseFont = true;
            this.tileNavPane.AppearanceHovered.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.tileNavPane.AppearanceHovered.Options.UseFont = true;
            this.tileNavPane.AppearanceSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.tileNavPane.AppearanceSelected.Options.UseFont = true;
            this.tileNavPane.Buttons.Add(this.navTitle);
            this.tileNavPane.Buttons.Add(this.cmdDisplay);
            this.tileNavPane.Buttons.Add(this.cmdDashboardDesign);
            this.tileNavPane.Buttons.Add(this.cmdClose);
            // 
            // tileNavCategory1
            // 
            this.tileNavPane.DefaultCategory.Name = "tileNavCategory1";
            // 
            // 
            // 
            this.tileNavPane.DefaultCategory.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileNavPane.Location = new System.Drawing.Point(12, 12);
            this.tileNavPane.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tileNavPane.Name = "tileNavPane";
            this.tileNavPane.Size = new System.Drawing.Size(1027, 46);
            this.tileNavPane.TabIndex = 12;
            this.tileNavPane.Text = "tileNavPane";
            // 
            // navTitle
            // 
            this.navTitle.Appearance.Font = new System.Drawing.Font("맑은 고딕", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.Appearance.Options.UseFont = true;
            this.navTitle.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.AppearanceHovered.Options.UseFont = true;
            this.navTitle.AppearanceSelected.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.navTitle.AppearanceSelected.Options.UseFont = true;
            this.navTitle.Caption = "사용자 대시보드";
            this.navTitle.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("navTitle.ImageOptions.SvgImage")));
            this.navTitle.Name = "navTitle";
            // 
            // cmdDisplay
            // 
            this.cmdDisplay.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Left;
            this.cmdDisplay.Caption = "검 색";
            this.cmdDisplay.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdDisplay.ImageOptions.Image")));
            this.cmdDisplay.Name = "cmdDisplay";
            // 
            // cmdDashboardDesign
            // 
            this.cmdDashboardDesign.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdDashboardDesign.Caption = "대시보드 작성";
            this.cmdDashboardDesign.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdDashboardDesign.ImageOptions.Image")));
            this.cmdDashboardDesign.Name = "cmdDashboardDesign";
            this.cmdDashboardDesign.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdDashboardDesign_ElementClick);
            // 
            // cmdClose
            // 
            this.cmdClose.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdClose.Caption = "닫 기";
            this.cmdClose.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdClose.ImageOptions.Image")));
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdClose_ElementClick);
            // 
            // layoutControlItem14
            // 
            this.layoutControlItem14.Control = this.tileNavPane;
            this.layoutControlItem14.ControlAlignment = System.Drawing.ContentAlignment.TopLeft;
            this.layoutControlItem14.CustomizationFormText = "layoutControlItem1";
            this.layoutControlItem14.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem14.MaxSize = new System.Drawing.Size(0, 50);
            this.layoutControlItem14.MinSize = new System.Drawing.Size(104, 50);
            this.layoutControlItem14.Name = "layoutControlItem14";
            this.layoutControlItem14.OptionsPrint.AppearanceItem.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItem.Options.UseFont = true;
            this.layoutControlItem14.OptionsPrint.AppearanceItemControl.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItemControl.Options.UseFont = true;
            this.layoutControlItem14.OptionsPrint.AppearanceItemText.Font = new System.Drawing.Font("굴림", 9F);
            this.layoutControlItem14.OptionsPrint.AppearanceItemText.Options.UseFont = true;
            this.layoutControlItem14.Size = new System.Drawing.Size(1031, 50);
            this.layoutControlItem14.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem14.Text = "layoutControlItem1";
            this.layoutControlItem14.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem14.TextVisible = false;
            // 
            // dashboardViewer
            // 
            this.dashboardViewer.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(210)))), ((int)(((byte)(210)))), ((int)(((byte)(210)))));
            this.dashboardViewer.Appearance.Options.UseBackColor = true;
            this.dashboardViewer.AsyncMode = false;
            this.dashboardViewer.Location = new System.Drawing.Point(12, 62);
            this.dashboardViewer.Name = "dashboardViewer";
            this.dashboardViewer.Size = new System.Drawing.Size(1027, 579);
            this.dashboardViewer.TabIndex = 13;
            this.dashboardViewer.DashboardLoaded += new DevExpress.DashboardWin.DashboardLoadedEventHandler(this.dashboardViewer_DashboardLoaded);
            this.dashboardViewer.ConfigureDataConnection += new DevExpress.DashboardCommon.DashboardConfigureDataConnectionEventHandler(this.dashboardViewer_ConfigureDataConnection);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.dashboardViewer;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 50);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(1031, 583);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // frmUserDashboardView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1051, 653);
            this.Controls.Add(this.layoutControl1);
            this.Name = "frmUserDashboardView";
            this.Text = "frmUserDashboardView";
            this.Load += new System.EventHandler(this.frmUserDashboardView_Load);
            this.Shown += new System.EventHandler(this.frmUserDashboardView_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dashboardViewer)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.DashboardWin.DashboardViewer dashboardViewer;
        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane;
        private DevExpress.XtraBars.Navigation.NavButton navTitle;
        private DevExpress.XtraBars.Navigation.NavButton cmdDisplay;
        private DevExpress.XtraBars.Navigation.NavButton cmdDashboardDesign;
        private DevExpress.XtraBars.Navigation.NavButton cmdClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem14;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}