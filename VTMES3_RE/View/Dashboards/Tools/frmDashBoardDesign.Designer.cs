
namespace VTMES3_RE.View.Dashboards.Tools
{
    partial class frmDashBoardDesign
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
            this.dashboardBarAndDockingController1 = new DevExpress.DashboardWin.Native.DashboardBarAndDockingController(this.components);
            this.dashboardDesigner = new DevExpress.DashboardWin.DashboardDesigner();
            ((System.ComponentModel.ISupportInitialize)(this.dashboardBarAndDockingController1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dashboardDesigner)).BeginInit();
            this.SuspendLayout();
            // 
            // dashboardBarAndDockingController1
            // 
            this.dashboardBarAndDockingController1.PropertiesBar.AllowLinkLighting = false;
            this.dashboardBarAndDockingController1.PropertiesDocking.ViewStyle = DevExpress.XtraBars.Docking2010.Views.DockingViewStyle.Classic;
            // 
            // dashboardDesigner
            // 
            this.dashboardDesigner.AsyncMode = true;
            this.dashboardDesigner.BarAndDockingController = this.dashboardBarAndDockingController1;
            this.dashboardDesigner.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dashboardDesigner.Location = new System.Drawing.Point(0, 0);
            this.dashboardDesigner.Margin = new System.Windows.Forms.Padding(5);
            this.dashboardDesigner.Name = "dashboardDesigner";
            this.dashboardDesigner.Size = new System.Drawing.Size(1078, 683);
            this.dashboardDesigner.TabIndex = 0;
            this.dashboardDesigner.ConfigureDataConnection += new DevExpress.DashboardCommon.DashboardConfigureDataConnectionEventHandler(this.dashboardDesigner_ConfigureDataConnection);
            this.dashboardDesigner.DashboardSaving += new DevExpress.DashboardWin.DashboardSavingEventHandler(this.dashboardDesigner_DashboardSaving);
            this.dashboardDesigner.DashboardClosing += new DevExpress.DashboardWin.DashboardClosingEventHandler(this.dashboardDesigner_DashboardClosing);
            // 
            // frmDashBoardDesign
            // 
            this.AllowFormGlass = DevExpress.Utils.DefaultBoolean.False;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1078, 683);
            this.Controls.Add(this.dashboardDesigner);
            this.Name = "frmDashBoardDesign";
            this.Text = "대쉬보드 디자이너";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmDashBoardDesign_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dashboardBarAndDockingController1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dashboardDesigner)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.DashboardWin.DashboardDesigner dashboardDesigner;
        private DevExpress.DashboardWin.Native.DashboardBarAndDockingController dashboardBarAndDockingController1;
    }
}