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
using VTMES3_RE.IFRYDataSetTableAdapters;

namespace VTMES3_RE.View.WorkManager
{
    public partial class frmEmployeeWorkTimeDef : DevExpress.XtraEditors.XtraForm
    {
        public frmEmployeeWorkTimeDef()
        {
            InitializeComponent();
        }

        private void frmEmployeeWorkTimeDef_Load(object sender, EventArgs e)
        {
            this.employeeWorkTimeDefTableAdapter.Fill(this.iFRYDataSet.EmployeeWorkTimeDef);
        }

        // 등록
        private void cmdAdd_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            foreach(DataRowView drv in employeeWorkTimeDefBindingSource)
            {
                if ((drv["Gubun"] ?? "").ToString() == "")
                {
                    return;
                }
            }

            employeeWorkTimeDefBindingSource.AddNew();

            DataRowView newItem = (DataRowView)employeeWorkTimeDefBindingSource.Current;
            newItem["Gubun"] = "";
            newItem["RegularYn"] = "N";

        }
        // 저장
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                foreach (DataRowView drv in employeeWorkTimeDefBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        if ((drv["Gubun"] ?? "").ToString() == "")
                        {
                            MessageBox.Show("구분을 입력하세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (drv["RegularYn"].ToString() == "Y")
                        {
                            if (drv["StartTime"] == DBNull.Value)
                            {
                                MessageBox.Show("정규 근무는 시작 시간이 필수 입력 항목입니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            if (drv["WorkHour"] == DBNull.Value || drv["WorkHour"].ToString() == "")
                            {
                                MessageBox.Show("정규 근무는 근무 시간이 필수 입력 항목입니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        drv["CreId"] = WrGlobal.LoginID;
                        drv["CreDt"] = DateTime.Now;
                    }
                    else if (drv.Row.RowState == DataRowState.Modified)
                    {
                        if ((drv["Gubun"] ?? "").ToString() == "")
                        {
                            MessageBox.Show("구분을 입력하세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (drv["RegularYn"].ToString() == "Y")
                        {
                            if (drv["StartTime"] == DBNull.Value)
                            {
                                MessageBox.Show("정규 근무는 시작 시간이 필수 입력 항목입니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            if (drv["WorkHour"] == DBNull.Value || drv["WorkHour"].ToString() == "")
                            {
                                MessageBox.Show("정규 근무는 근무 시간이 필수 입력 항목입니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModDt"] = DateTime.Now;
                    }
                }

                employeeWorkTimeDefBindingSource.EndEdit();
                employeeWorkTimeDefTableAdapter.Update(iFRYDataSet.EmployeeWorkTimeDef);

                MessageBox.Show("저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 삭제
        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            string gubun = (gvEmployeeWorkTimeDef.GetFocusedRowCellValue("Gubun") ?? "").ToString();

            if (MessageBox.Show(string.Format("선택한 '{0}' 항목을 삭제하시겠습니까?", gubun), "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }//end if

            employeeWorkTimeDefBindingSource.RemoveCurrent();
            employeeWorkTimeDefTableAdapter.Update(iFRYDataSet.EmployeeWorkTimeDef);

            MessageBox.Show("삭제되었습니다.", "삭제", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        // 닫기
        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
    }
}