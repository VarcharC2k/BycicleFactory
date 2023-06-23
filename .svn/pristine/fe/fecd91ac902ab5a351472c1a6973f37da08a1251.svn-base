using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using DevExpress.XtraReports.UI;
using DevExpress.XtraReports.Extensions;
using DevExpress.ReportServer.ServiceModel.DataContracts;
using DevExpress.XtraReports.UserDesigner;
using VTMES3_RE.View.Reports.Tools;
// ...

namespace VTMES3_RE.Common
{
    public class DataSetReportStorage : ReportStorageExtension {
        StorageDataSet dataSet;
        StorageDataSetTableAdapters.ReportItemTableAdapter adapter = new StorageDataSetTableAdapters.ReportItemTableAdapter();

        public DataSetReportStorage() {
        }

        StorageDataSet DataSet {
            get {
                if (dataSet == null) {
                    dataSet = new StorageDataSet();

                    adapter.Fill(dataSet.ReportItem, WrGlobal.CorpID);
                    // Populate a dataset from an XML file specified in fileName.
                }
                return dataSet;
            }
        }
        StorageDataSet.ReportItemDataTable ReportStorage {
            get {
                return DataSet.ReportItem;
            }
        }

        public override bool CanSetData(string url) {
            // Always return true to confirm that the SetData method is available.
            return true;
        }
        public override bool IsValidUrl(string url) {
            return !string.IsNullOrEmpty(url);
        }
        public override byte[] GetData(string url) {
            // Get a dataset row containing the report.
            StorageDataSet.ReportItemRow row = FindRow(url);
            if (row != null)
                return System.Text.Encoding.UTF8.GetBytes(row.LayoutData.Replace("?<?", "<?"));
            return new byte[] { };
        }
        StorageDataSet.ReportItemRow FindRow(string url) {
            DataRow[] result = ReportStorage.Select(string.Format("ReportName = '{0}'", url));
            if (result.Length > 0)
                return result[0] as StorageDataSet.ReportItemRow;
            return null;
        }
        public override void SetData(XtraReport report, string url) {
            StorageDataSet.ReportItemRow row = FindRow(url);
            // Write the report to a corresponding row in the dataset.

            byte[] m_Buffer;
            string m_XML = "";

            if (row == null)
            {
                row = ReportStorage.NewReportItemRow();
                row["CorpId"] = "RY";
                row["ReportName"] = url;

                using (MemoryStream ms = new MemoryStream())
                {
                    report.SaveLayoutToXml(ms);
                    m_Buffer = ms.ToArray();

                    m_XML = System.Text.Encoding.UTF8.GetString(m_Buffer);
                    m_XML = m_XML.Replace("?<?", "<?");

                    row["LayoutData"] = m_XML;
                }

                ReportStorage.Rows.Add(row);
            }
            else
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    report.SaveLayoutToXml(ms);
                    m_Buffer = ms.ToArray();

                    m_XML = System.Text.Encoding.UTF8.GetString(m_Buffer);
                    m_XML = m_XML.Replace("?<?", "<?");

                    row["LayoutData"] = m_XML;
                }
            }

            adapter.Update(ReportStorage);
        }
        byte[] GetBuffer(XtraReport report) {
            using (MemoryStream stream = new MemoryStream()) {
                report.SaveLayout(stream);
                return stream.ToArray();
            }
        }
        public override string GetNewUrl() {
            // Show the report selection dialog and return a URL for a selected report.
            frmStorageEdit form = CreateForm();
            form.ReportNameEdit.Enabled = false;
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                return form.ReportNameEdit.Text;
            return string.Empty;
        }
        frmStorageEdit CreateForm() {
            frmStorageEdit form = new frmStorageEdit();
            foreach (string item in GetUrls())
                form.ReportItemListBox.Items.Add(item);
            return form;
        }
        public override string SetNewData(XtraReport report, string defaultUrl) {
            frmStorageEdit form = CreateForm();
            form.ReportNameEdit.Text = defaultUrl;
            form.ReportItemListBox.Enabled = false;
            // Show the save dialog to get a URL for a new report.
            if (form.ShowDialog() == DialogResult.OK) {
                string url = form.ReportNameEdit.Text;
                if (!string.IsNullOrEmpty(url) && !form.ReportItemListBox.Items.Contains(url))
                {
                    TypeDescriptor.GetProperties(typeof(XtraReport))["DisplayName"].SetValue(report, url);
                    SetData(report, url);
                    return url;
                }
                else
                {
                    MessageBox.Show("Incorrect report name", "Error",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }
            return string.Empty;
        }

        // The following code is intended to support selection of a value for 
        // the Report Source Url property of Subreport controls.
        // (Use this code to avoid assigning the master report as a 
        // detail report's source.)

        public override bool GetStandardUrlsSupported(ITypeDescriptorContext context) {
            // Always return true to confirm that the GetStandardUrls method is available.
            return true;
        }
        public override string[] GetStandardUrls(ITypeDescriptorContext context) {
            if (context != null && context.Instance is XRSubreport) {
                XRSubreport xrSubreport = context.Instance as XRSubreport;
                if (xrSubreport.RootReport !=
                    null && xrSubreport.RootReport.Extensions.TryGetValue("StorageID", out storageID)) {
                    List<string> result = GetUrlsCore(CanPassId);
                    return result.ToArray();
                }
            }
            return GetUrls();
        }
        string storageID;
        bool CanPassId(string id) {
            return id != storageID;
        }
        string[] GetUrls() {
            return GetUrlsCore(null).ToArray();
        }
        List<string> GetUrlsCore(Predicate<string> method) {
            List<string> list = new List<string>();
            foreach (StorageDataSet.ReportItemRow row in ReportStorage.Rows)
                if (method == null || method(row.ReportId.ToString()))
                    list.Add(row.ReportName);
            return list;
        }
    }
}
