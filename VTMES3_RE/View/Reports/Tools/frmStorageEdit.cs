using DevExpress.DashboardWin.Design;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.Common;

namespace VTMES3_RE.View.Reports.Tools
{
    /// <summary>
    /// 레포트 스토리지에서 해당 레포트명 저장
    /// </summary>
    public partial class frmStorageEdit : DevExpress.XtraEditors.XtraForm
    {
        public frmStorageEdit()
        {
            InitializeComponent();
        }

        private void StorageEditorForm_Load(object sender, EventArgs e)
        {
            if (ReportItemListBox.Items.Count > 0 && string.IsNullOrEmpty(ReportItemListBox.Text))
                ReportItemListBox.SelectedIndex = 0;
        }

        private void ReportNameEdit_TextChanged(object sender, EventArgs e)
        {
            btnOK.Enabled = !string.IsNullOrEmpty(ReportNameEdit.Text);
        }

        private void ReportItemListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReportNameEdit.Text = ReportItemListBox.SelectedItem.ToString();
        }
    }
}