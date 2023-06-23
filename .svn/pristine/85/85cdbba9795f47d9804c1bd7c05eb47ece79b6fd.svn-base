using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraReports.Wizards.Templates;
using DevExpress.XtraSpreadsheet.Model;
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
using VTMES3_RE.Models;

namespace VTMES3_RE.View.CheckSheet
{
    public partial class frmMachineCheckSheet : DevExpress.XtraEditors.XtraForm
    {
        // 점검표 템플릿 파일 경로 
        clsWork work = new clsWork();
        string folderName = Application.StartupPath + @"\Templates\MachineCheckSheet";
        bool isNewDocument = true;

        public frmMachineCheckSheet()
        {
            InitializeComponent();
        }

        private void frmMachineCheckSheet_Load(object sender, EventArgs e)
        {
            DirectoryInfo templateDi = new DirectoryInfo(folderName);
            DirectoryInfo[] directories = templateDi.GetDirectories();

            // 팀별 파일 리스트 가져오기
            DataTable fileDt = new DataTable();
            fileDt.Columns.Add("TeamName", typeof(string));
            fileDt.Columns.Add("SheetName", typeof(string));

            foreach (DirectoryInfo di in directories)
            {
                foreach (FileInfo file in di.GetFiles())
                {
                    DataRow dr = fileDt.NewRow();
                    dr.BeginEdit();
                    dr["TeamName"] = di.Name;
                    dr["SheetName"] = file.Name;
                    dr.EndEdit();
                    fileDt.Rows.Add(dr);
                }
            }
            // 파일리스트 바인딩
            gcTemplateFile.DataSource = fileDt;

            // 점검년도 목록 콤보 설정
            DataView yearDv = work.GetMachineCheckSheetYear();
            foreach (DataRowView drv in yearDv)
            {
                checkYearColComboBox.Items.Add(drv["CheckYear"].ToString());
            }
        }
        // 선택된 로우 스타일 설정
        private void gvTemplateFile_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (e.RowHandle == view.FocusedRowHandle)
            {
                //Apply the appearance of the SelectedRow
                e.Appearance.Assign(view.PaintAppearance.SelectedRow);
                e.Appearance.Options.UseForeColor = true;
            }//end if
        }
        // 점검표 변경 이벤트 처리 -> 해당 점검표 내역 점검표 리스트에 표시
        private void gvTemplateFile_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            //if (!gvTemplateFile.IsDataRow(e.FocusedRowHandle)) return;
            gvMachineCheckSheet.FocusInvalidRow();

            string teamName = (gvTemplateFile.GetFocusedRowCellValue("TeamName") ?? "").ToString();
            string sheetName = (gvTemplateFile.GetFocusedRowCellValue("SheetName") ?? "").ToString();
            machineCheckSheetTableAdapter.FillBySelect(this.iFRYDataSet.MachineCheckSheet, teamName, sheetName);

        }
        // 점검표 추가 클릭 -> 선택된 점검표 템플릿을 불러옴
        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (!gvTemplateFile.IsDataRow(gvTemplateFile.FocusedRowHandle)) return;

            string teamName = (gvTemplateFile.GetFocusedRowCellValue("TeamName") ?? "").ToString();
            string sheetName = (gvTemplateFile.GetFocusedRowCellValue("SheetName") ?? "").ToString();

            machineCheckSheetBindingSource.AddNew();
            //DataRowView drv = (DataRowView)machineCheckSheetBindingSource.Current;
            //drv["TeamName"] = teamName;
            //drv["SheetName"] = sheetName;
            //drv["CheckYear"] = DateTime.Today.Year;
            //drv["CheckMonth"] = DateTime.Today.Month;

            gvMachineCheckSheet.SetFocusedRowCellValue("TeamName", teamName);
            gvMachineCheckSheet.SetFocusedRowCellValue("SheetName", sheetName);
            gvMachineCheckSheet.SetFocusedRowCellValue("CheckYear", DateTime.Today.Year);
            gvMachineCheckSheet.SetFocusedRowCellValue("CheckMonth", DateTime.Today.Month);

            excelSheetControl.LoadDocument(folderName + "\\" + teamName + "\\" + sheetName);
        }
        // 점검표 리스트 포커스 변경 이벤트 -> 해당 점검표 DB에서 읽어와 엑셀 시트에 표시
        private void gvMachineCheckSheet_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (!gvTemplateFile.IsDataRow(gvTemplateFile.FocusedRowHandle)) return;

            if (!gvMachineCheckSheet.IsDataRow(e.FocusedRowHandle))
            {
                if (!isNewDocument) excelSheetControl.CreateNewDocument();
                return;
            }

            DataRowView drv = (DataRowView)machineCheckSheetBindingSource.Current;

            if (drv != null && drv["Doc"] != DBNull.Value)
            {
                using (MemoryStream stream = new MemoryStream((byte[])drv["Doc"]))
                {
                    excelSheetControl.LoadDocument(stream, DocumentFormat.Xlsx);
                }

                isNewDocument = false;
            }
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
        // 점검표 저장 인벤트 처리
        private void btnSave_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)machineCheckSheetBindingSource.Current;

            if (drv == null)
            {
                MessageBox.Show("점검표를 선택하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                this.Validate();
                using (MemoryStream stream = new MemoryStream())
                {
                    excelSheetControl.SaveDocument(stream, DocumentFormat.Xlsx);
                    drv["Doc"] = stream.ToArray();
                }
                machineCheckSheetBindingSource.EndEdit();
                machineCheckSheetTableAdapter.Update(iFRYDataSet.MachineCheckSheet);

                MessageBox.Show("저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 점검표 파일로 저장
        private void cmdDownload_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (gvMachineCheckSheet.IsDataRow(gvMachineCheckSheet.FocusedRowHandle))
            {
                if (fbDialog.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream stream = new FileStream(fbDialog.SelectedPath + "\\" + gvMachineCheckSheet.GetFocusedRowCellValue("CheckYear") + "_" + gvMachineCheckSheet.GetFocusedRowCellValue("CheckMonth") + "_" + gvMachineCheckSheet.GetFocusedRowCellValue("SheetName"),
                            FileMode.Create, FileAccess.ReadWrite))
                    {
                        excelSheetControl.SaveDocument(stream, DocumentFormat.Xlsx);
                    }
                    MessageBox.Show(fbDialog.SelectedPath + "\\" + gvMachineCheckSheet.GetFocusedRowCellValue("CheckYear") + "_" + gvMachineCheckSheet.GetFocusedRowCellValue("CheckMonth") + "_" + gvMachineCheckSheet.GetFocusedRowCellValue("SheetName") + Environment.NewLine + "저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}