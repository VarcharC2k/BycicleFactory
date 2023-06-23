﻿namespace VTMES3_RE.View.CamstarInf
{
    partial class frmDynamicComponentIssue
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDynamicComponentIssue));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.tileNavPane = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.navTitle = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdSearch = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdExecute = new DevExpress.XtraBars.Navigation.NavButton();
            this.cmdClose = new DevExpress.XtraBars.Navigation.NavButton();
            this.excelSheetControl = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.statusStrip11 = new System.Windows.Forms.StatusStrip();
            this.lblMemo = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblCnt = new System.Windows.Forms.ToolStripStatusLabel();
            this.pgsBar = new System.Windows.Forms.ToolStripProgressBar();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem14 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).BeginInit();
            this.statusStrip11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.tileNavPane);
            this.layoutControl1.Controls.Add(this.excelSheetControl);
            this.layoutControl1.Controls.Add(this.statusStrip11);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(952, 513);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
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
            this.tileNavPane.Buttons.Add(this.cmdSearch);
            this.tileNavPane.Buttons.Add(this.cmdExecute);
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
            this.tileNavPane.Size = new System.Drawing.Size(928, 46);
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
            this.navTitle.Caption = "ComponentIssue 관리";
            this.navTitle.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("navTitle.ImageOptions.SvgImage")));
            this.navTitle.Name = "navTitle";
            // 
            // cmdSearch
            // 
            this.cmdSearch.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdSearch.Caption = "조회";
            this.cmdSearch.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdInit.ImageOptions.Image")));
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdSearch_ElementClick);
            // 
            // cmdExecute
            // 
            this.cmdExecute.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdExecute.Caption = "제 출";
            this.cmdExecute.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdExecute.ImageOptions.Image")));
            this.cmdExecute.Name = "cmdExecute";
            this.cmdExecute.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdExecute_ElementClick);
            // 
            // cmdClose
            // 
            this.cmdClose.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Right;
            this.cmdClose.Caption = "닫 기";
            this.cmdClose.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("cmdClose.ImageOptions.Image")));
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.ElementClick += new DevExpress.XtraBars.Navigation.NavElementClickEventHandler(this.cmdClose_ElementClick);
            // 
            // excelSheetControl
            // 
            this.excelSheetControl.Location = new System.Drawing.Point(12, 62);
            this.excelSheetControl.Name = "excelSheetControl";
            this.excelSheetControl.Size = new System.Drawing.Size(928, 414);
            this.excelSheetControl.TabIndex = 16;
            this.excelSheetControl.Text = "spreadsheetControl1";
            // 
            // statusStrip11
            // 
            this.statusStrip11.AutoSize = false;
            this.statusStrip11.BackColor = System.Drawing.Color.LightGray;
            this.statusStrip11.Dock = System.Windows.Forms.DockStyle.None;
            this.statusStrip11.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblMemo,
            this.lblCnt,
            this.pgsBar});
            this.statusStrip11.Location = new System.Drawing.Point(12, 480);
            this.statusStrip11.Name = "statusStrip11";
            this.statusStrip11.Size = new System.Drawing.Size(928, 21);
            this.statusStrip11.TabIndex = 5;
            this.statusStrip11.Text = "statusStrip1";
            // 
            // lblMemo
            // 
            this.lblMemo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.lblMemo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lblMemo.Name = "lblMemo";
            this.lblMemo.Size = new System.Drawing.Size(0, 16);
            this.lblMemo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCnt
            // 
            this.lblCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(709, 16);
            this.lblCnt.Spring = true;
            this.lblCnt.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pgsBar
            // 
            this.pgsBar.Name = "pgsBar";
            this.pgsBar.Padding = new System.Windows.Forms.Padding(1);
            this.pgsBar.Size = new System.Drawing.Size(202, 15);
            this.pgsBar.Step = 100;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem14,
            this.layoutControlItem2,
            this.layoutControlItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(952, 513);
            this.Root.TextVisible = false;
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
            this.layoutControlItem14.Size = new System.Drawing.Size(932, 50);
            this.layoutControlItem14.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem14.Text = "layoutControlItem1";
            this.layoutControlItem14.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem14.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.excelSheetControl;
            this.layoutControlItem2.ControlAlignment = System.Drawing.ContentAlignment.TopLeft;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem2";
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 50);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(932, 418);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.statusStrip11;
            this.layoutControlItem3.ControlAlignment = System.Drawing.ContentAlignment.TopLeft;
            this.layoutControlItem3.CustomizationFormText = "layoutControlItem2";
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 468);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(0, 25);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(104, 25);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(932, 25);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.Text = "layoutControlItem2";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // frmDynamicComponentIssue
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(952, 513);
            this.Controls.Add(this.layoutControl1);
            this.Name = "frmDynamicComponentIssue";
            this.Text = "frmCMOS_ContainerAttr";
            this.Load += new System.EventHandler(this.frmCMOS_ContainerAttr_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tileNavPane)).EndInit();
            this.statusStrip11.ResumeLayout(false);
            this.statusStrip11.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem14)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane;
        private DevExpress.XtraBars.Navigation.NavButton navTitle;
        private DevExpress.XtraBars.Navigation.NavButton cmdExecute;
        private DevExpress.XtraBars.Navigation.NavButton cmdClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem14;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl excelSheetControl;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private System.Windows.Forms.StatusStrip statusStrip11;
        private System.Windows.Forms.ToolStripStatusLabel lblMemo;
        private System.Windows.Forms.ToolStripStatusLabel lblCnt;
        private System.Windows.Forms.ToolStripProgressBar pgsBar;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraBars.Navigation.NavButton cmdSearch;
    }
}