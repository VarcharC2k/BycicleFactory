using DevExpress.Charts.Native;
using DevExpress.CodeParser;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraRichEdit.Import.OpenXml;
using DevExpress.XtraSplashScreen;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using static DevExpress.Data.Mask.Internal.RegExMaskMath.DraftFiniteAutomaton<T>;
using VTMES3_RE.Common;
using VTMES3_RE.Models;

namespace VTMES3_RE.View.WorkManager
{
    // 이순희 PM 작성
    public partial class frmCMOS_FtDataUpload : DevExpress.XtraEditors.XtraForm
    {
        clsWork work = new clsWork();

        public frmCMOS_FtDataUpload()
        {
            InitializeComponent();
            
            if (WrGlobal.Camster_Common == null)
            {
                WrGlobal.Camster_Common = new CamstarCommon();
            }
        }

        private void rg_PART_EditValueChanged(object sender, EventArgs e)
        {
            if (rg_PART.EditValue.ToString() == "IOS")
            {
                rg_EQUIP.Enabled = true;
            }
            else
            {
                rg_EQUIP.Enabled = false;
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                string fileExt = string.Empty;
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user
                {
                    filePath = file.FileName; //get the path of the file
                    fileExt = Path.GetExtension(filePath); //get the file extension

                    if (fileExt.CompareTo(".csv") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        lbl_filePath.Text = filePath;
                    }
                    else
                    {
                        MessageBox.Show("Please choose .csv file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR");
            }
        }        

        private void cmos_grid_view(string[] rows)
        {
            DataTable table = new DataTable();

            int cnt = 0;
            foreach (string d in rows)
            {
                string[] split = d.Split(',');

                if (split.Length > 0)
                {
                    if (cnt == 0)
                    {
                        foreach (string colName in split)
                        {                            
                            table.Columns.Add(colName);
                        }
                        // 마지막 열에 Pass_Fail, Result 칼럼 생성
                        table.Columns.Add("Pass_Fail");
                        table.Columns.Add("기준정보(오류)");
                        table.Columns.Add("Result");
                        table.Columns.Add("CamMessage");
                        table.Columns.Add("CamResult");
                    }
                    else
                    {
                        table.Rows.Add(split);
                    }
                }
                cnt ++;
            }

            // Grid 초기화
            gcCheckDetail.DataSource = null;
            gvCheckDetail.Columns.Clear();

            // GRID 에 데이터 바인딩
            gcCheckDetail.DataSource = table;
            gvCheckDetail.BestFitColumns();            
        }

        private void ios_grid_view(string[] rows)
        {
            DataTable table = new DataTable();

            int cnt = 0;
            foreach (string d in rows)
            {
                string[] split = d.Split(',');

                if (split.Length > 0)
                {
                    if (cnt == 0 )
                    {
                        // 칼럼이 01.xxx 02.xxx 로 생성 되기 위함
                        int col_no = 1;
                        foreach (string colName in split)
                        {
                            table.Columns.Add( string.Format("{0:00}.", col_no) + colName);
                            col_no++;
                        }
                        // 마지막 열에 Pass_Fail, Result 칼럼 생성
                        table.Columns.Add("Pass_Fail");
                        table.Columns.Add("기준정보(오류)");
                        table.Columns.Add("Result");
                        table.Columns.Add("CamMessage");
                        table.Columns.Add("CamResult");
                    }
                    //else if(cnt == 1)
                    //{
                    //    // 칼럼의 index 번호( 0, 1, 2,,, )
                    //    int col_idx = 0;
                    //    foreach (string colName in split)
                    //    {
                    //        table.Columns[col_idx].ColumnName = table.Columns[col_idx].ColumnName + " (" + colName + ")";
                    //        col_idx++;
                    //    }
                    //}
                    else
                    {
                        table.Rows.Add(split);
                    }
                }
                cnt = cnt + 1;
            }

            // Grid 초기화
            gcCheckDetail.DataSource = null;
            gvCheckDetail.Columns.Clear();
            
            // GRID 에 데이터 바인딩
            gcCheckDetail.DataSource = table;
            gvCheckDetail.BestFitColumns();
        }

        // 측정값의 상한치, 하한치 검증후 Pass/Fail 등의 정보를 Grid 에 업데이트
        // key_no = Container 번호, equip = IOS 의 호기 번호, cell = row를 읽어 column 배열에 담음, i = 엑셀의 행번호
        private string Spec_High_Low_Check(string key_no, int equip, string[] cell, int i)
        {
            // key 값 으로 SAP Code 를 찾고, SAP Code 의 상항치, 하한치 값을 가져오기
            DataView dv = new DataView();
            dv = work.KeynoHighLow(key_no);

            // SAP Code 의 상항치, 하한치 값으로 Spec Out 을 찾기
            int pass_cnt = 0;       // Pass_Fail 항목 ( pass 갯수 )
            int fail_cnt = 0;       // Pass_Fail 항목 ( fail 갯수 )
            int std_err_cnt = 0;    // 기준정보(오류) 항목 ( 기준정보 틀린 갯수 )
            int ng_cnt = 0;         // Result 항목 ( 기준 정보가 없는 경우, fail 이 난 경우등 ,,, )
            string udcd = "";       // UDCD Name 을 반환

            // Spec 기준 값이 있으면 로직 수행
            if (dv.Table.Rows.Count > 0)
            {
                // udcd 이름 가져오기
                udcd = dv.Table.Rows[0][4].ToString();
                
                foreach (DataRow dr in dv.Table.Rows)
                {
                    // DISPLAY_AREA 가 xx/yy 인경우 xx 는 1호기, yy는 2호기를 뜻한다.
                    // xx, yy 는 업로드 엑셀 시트의 칼럼명을 나타 낸다.
                    // xx, yy를 char 로 변환후 -65 를 하면, 0, 1, 2, 3,,,, 등 배열의 순번으로 사용 가능 하다.
                    string[] display_areas = dr["DISPLAY_AREA"].ToString().ToUpper().Split('/');
                    string display_area = display_areas[equip - 1].Trim();     //////////////error

                    int col = 0;

                    if (display_area.Length == 1)
                    {
                        char c = display_area[0];
                        col = Convert.ToInt32(c) - 65;
                    }
                    else if (display_area.Length == 2)
                    {
                        char c1 = display_area[0];
                        col = (Convert.ToInt32(c1) - 64) * 26;

                        char c2 = display_area[1];
                        col += Convert.ToInt32(c2) - 65;
                    }

                    // Spec out check
                    // 
                    if (cell[col].ToString().Length > 0)
                    {
                        if (dr["LOW_VALUE"].ToString().Length > 0)
                        {
                            if (dr["HIGH_VALUE"].ToString().Length > 0)
                            {
                                // High, Low 둘다 있는 경우
                                // High 값 비교 구분자 잘못되어 있어 방향 수정
                                if (Convert.ToDouble(dr["LOW_VALUE"].ToString()) <= Convert.ToDouble(cell[col].ToString()) && Convert.ToDouble(cell[col].ToString()) <= Convert.ToDouble(dr["HIGH_VALUE"].ToString()))
                                {
                                    pass_cnt++;
                                }
                                else
                                {
                                    fail_cnt++;
                                }
                            }
                            else
                            {
                                // Low 만 있는 경우
                                if (Convert.ToDouble(dr["LOW_VALUE"].ToString()) <= Convert.ToDouble(cell[col].ToString()))
                                {
                                    pass_cnt++;
                                }
                                else
                                {
                                    fail_cnt++;
                                }
                            }
                        }
                        // High 만 있는 경우
                        else if (dr["HIGH_VALUE"].ToString().Length > 0)
                        {
                            // High 만 있는 경우
                            if (Convert.ToDouble(cell[col].ToString()) <= Convert.ToDouble(dr["HIGH_VALUE"].ToString()))
                            {
                                pass_cnt++;
                            }
                            else
                            {
                                fail_cnt++;
                            }
                        }
                        else
                        {
                            ng_cnt++;
                            std_err_cnt++;
                        }
                    }
                    else
                    {
                        ng_cnt++;
                        std_err_cnt++;
                    }
                }
            }
            else
            {
                ng_cnt++;
            }

            // Pass_Fail 항목에 표시
            if (fail_cnt > 0)
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Pass_Fail", "Fail");
            }
            else if (pass_cnt > 0)
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Pass_Fail", "Pass");
            }
            else
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Pass_Fail", "NG");
            }

            // 기준정보(오류) 항목에 표시  
            gvCheckDetail.SetRowCellValue(i - 1, "기준정보(오류)", std_err_cnt);

            // Result 항목에 표시
            if (ng_cnt > 0)
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Result", "NG");
            }
            else if (std_err_cnt > 0)
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Result", "NG");
            }
            else
            {
                gvCheckDetail.SetRowCellValue(i - 1, "Result", "검증성공");
            }

            return udcd;
        }

        // Result 에 NG 인경우
        // 해당 Row 를 Orange 색상으로 표시
        private void gvCheckDetail_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView gridView = sender as GridView; ;

            if (e.RowHandle < 0) return;

            if ((gridView.GetRowCellValue(e.RowHandle, "Result") ?? "").ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Orange;
                e.HighPriority = true;
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            string part = "";
            int equip = 1;

            // part = IOS, CMOS
            // equip = 1, 2
            if (rg_PART.EditValue is null)
            {
                MessageBox.Show("파트를 선택해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                part = rg_PART.EditValue.ToString();
            }

            if (rg_PART.EditValue.ToString() == "IOS" && rg_EQUIP.EditValue is null)
            {
                MessageBox.Show("업로드 할 데이터를 추출한 장비를 선택해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (rg_PART.EditValue.ToString() == "IOS" && !(rg_EQUIP.EditValue is null))
            {
                equip = Convert.ToInt16(rg_EQUIP.EditValue);
            }

            if (string.IsNullOrEmpty(lbl_filePath.Text))
            {
                MessageBox.Show("업로드 할 데이터 파일을 선택해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // key = Container 번호, keys = 엑셀 1Row
            string key = "";
            string keys = "";
            try
            {
                SplashScreenManager.ShowForm(this, typeof(frmWait), true, true, false);
                RemoveBlankRowsFromCVSFile(lbl_filePath.Text);

                Filehandle fh = new Filehandle(lbl_filePath.Text);

                // CMOS, IOS 업로드 파일의 인코딩이 달라서, Head 칼럼 고정 적용
                string[] rows = fh.Cmos_Ios_TextFileReadLine(part);

                //key (Wafer No or Serial No) 가 있는 col 찾기
                int keyCol = 0;
                //IOS의 경우 SerialNo로 판단을 하지만 Wafer 번호가 Container이므로 Container 검색을 위하여 Wafer Col 생성
                int WaferCol = 0;
                string[] cells = rows[0].Split(',');
                for (int i = 0; i < cells.Length; i++)
                {

                    //Serial과 WaferNo가 같이 있는 경우 Serial이 우선 순번에 있으므로 WaferNo가 키값이 되도록 하기 위하여 SerialNo 제거
                    if (cells[i].Trim().Equals("WaferNo") || cells[i].Trim().Equals("Barcode"))
                    {
                        keyCol = i;
                        key = "wafer";
                        break;
                    }
                    else if (cells[i].Trim().Equals("SerialNo"))
                    {
                        keyCol = i;
                        WaferCol = i + 1;
                        key = "serial";
                        break;
                    }
                }

                // CMOS, IOS 파일 검증
                if (part == "CMOS")
                {
                    if (key != "wafer")
                    {
                        MessageBox.Show("CMOS 업로드 파일이 아닙니다.\r\n파일을 확인 해주세요.");
                        return;
                    }
                }
                else if (part == "IOS")
                {
                    if (key != "serial")
                    {
                        MessageBox.Show("IOS 업로드 파일이 아닙니다.\r\n파일을 확인 해주세요.");
                        return;
                    }
                }

                // keyCol (Wafer No or Serial No) 뽑기
                int startRow = 1;
                int specRow = 1;
                string[] specCells;

                for (int i = 1; i < rows.Length; i++)
                {
                    if (rows[i].Contains("Spec") || rows[i].Contains("spec"))
                    {
                        specRow = i;                        // spec 이 있는 row 번호
                        specCells = rows[i].Split(',');     // spec row 를 읽어 specCells 배열에 담기
                        startRow++;                        // Start Row 번호
                    }
                    else
                    {
                        keys += rows[i].Split(',')[keyCol].Trim() + ",";
                    }
                }

                progressBarControl1.Properties.Maximum = rows.Length - startRow;
                progressBarControl1.Position = 0;
                labelControl1.Text = progressBarControl1.Position.ToString() + " / " + progressBarControl1.Properties.Maximum.ToString();

                labelControl1.Refresh();

                SplashScreenManager.CloseForm(false);

                string ErrorMsg = "";

                // CMOS 일 경우
                // csv 데이터 돌면서...값 찾아서 하나씩 집어넣기
                if (part == "CMOS")
                {
                    // grid 에 dt 바인딩
                    cmos_grid_view(rows);

                    // CMOS 업로드시 ID_KEY 값이 필요함
                    // 테이블의 ID_KEY 칼럼의 Max + 1 값을 받아옮
                    string max_1 = work.CmosFtDataCmosID_KEY();

                    // Error 일경우 빠져 나가기
                    if (max_1 == "") { return; }

                    // 
                    for (int i = 1; i < rows.Length; i++)
                    {
                        string[] cell = rows[i].Split(',');         // 1줄 읽어서 배열에 담기
                        string key_no = cell[keyCol].Trim();        // key 값 가져오기( Container 번호 )
                        string[] cam_return = null;                 // Camstar 에서 결과 Return 받아오기 위한
                        string udcd_Name = "";                      // UDCD Name 변수 정의

                        // 상한치, 하한치 체크 로직
                        // 결과 값 3가지를 Grid 에 Update
                        // udcd name 을 return 받아온다.
                        udcd_Name = Spec_High_Low_Check(key_no, equip, cell, i);

                        // Test 용 
                        //cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail("TEST1", "127955-05", "False", true);

                        // 검증성공인 경우만 DB Insert
                        if (gvCheckDetail.GetRowCellValue(i - 1, "Result").ToString() == "검증성공")
                        {
                            // Camstar 의 E Procedure 를 실행
                            // UDCD Name 이 Task List 로 사용
                            // Task 의 Pass / Fail 적용
                            if (gvCheckDetail.GetRowCellValue(i - 1, "Pass_Fail").ToString() == "Pass")
                            {
                                cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail(udcd_Name, key_no, "True", false);
                            }
                            else
                            {
                                cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail(udcd_Name, key_no, "False", false);
                            }

                            // Message 칼럼에 cam_return 메시지 적용
                            gvCheckDetail.SetRowCellValue(i - 1, "CamResult", cam_return[0]);
                            gvCheckDetail.SetRowCellValue(i - 1, "CamMessage", cam_return[1]);

                            if (cam_return[1] == "성공!")
                            {
                                // FT_CSV_INSERT 테이블에 원본 데이터 업로드
                                bool bl = work.FT_CSV_INSERT(part, "", rows[i].ToString());

                                // Error 일경우 빠져 나가기
                                if (!bl) { return; }

                                // 1 Row 를 읽어서 Procedure 로 전달
                                // Procedure 에서 30 칼럼으로 분리후 물리 칼럼에 저장
                                bool bl2 = work.CmosFtDataCmos(rows[i].ToString(), max_1);

                                // Error 일경우 빠져 나가기
                                if (!bl2) { return; }
                                // barcode 로 Camstar 의 StepEntryTxnId 를 찾고,
                             
                                //트랜잭션 순서 변경
                                // CmosFtDataCmos 테이블에 barcode 와 ID_KEY 가 같은 데이터를 찾아 StepEntryTxnId Update
                                bool bl3 = work.CmosFtDataCmos_StepEntryTxnId(key_no, max_1);

                                // Error 일경우 빠져 나가기
                                if (!bl3) { return; }
                            }
                        }

                        // 진행바 Update
                        progressBarControl1.Position = i;
                        progressBarControl1.Refresh();
                        labelControl1.Text = progressBarControl1.Position.ToString() + " / " + progressBarControl1.Properties.Maximum.ToString();
                        labelControl1.Refresh();
                    }
                }

                // IOS 일 경우
                // csv 데이터 돌면서...값 찾아서 하나씩 집어넣기
                if (part == "IOS")
                {
                    // grid 에 dt 바인딩
                    ios_grid_view(rows);

                    // IOS 업로드시 ID_KEY 값이 필요함
                    // Master 테이블의 ID_KEY 칼럼의 Max + 1 값을 받아옮
                    string max_1 = work.CmosFtDataIosID_KEY();

                    // Error 일경우 빠져 나가기
                    if (max_1 == "") { return; }

                    for (int i = 1; i < rows.Length; i++)
                    {   
                        string[] cell = rows[i].Split(',');         // 1줄 읽어서 배열에 담기
                        //string key_no = cell[keyCol].Trim();        // key 값 가져오기( Container 번호 )
                        string key_no = cell[WaferCol].Trim();      //Container 번호는 Wafer이므로 WaferCol로 변경
                        string[] cam_return = null;                 // Camstar 에서 결과 Return 받아오기 위한
                        string udcd_Name = "";                      // UDCD Name 변수 정의

                        // 상한치, 하한치 체크 로직
                        // 결과 값 3가지를 Grid 에 Update
                        // Spec 행은 수행 하지 않음
                        if (specRow != i)
                        {
                            udcd_Name = Spec_High_Low_Check(key_no, equip, cell, i);
                        }
                        else
                        {
                            // FT_CSV_INSERT 테이블에 원본 데이터 업로드
                            bool bl = work.FT_CSV_INSERT(part, equip.ToString(), rows[i].ToString());

                            // Error 일경우 빠져 나가기
                            if (!bl) { return; }
                        }

                        // Test 용 
                        //cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail("TEST1", "127955-05", "False", true);

                        // 검증성공인 경우만 DB Insert
                        // 결과가 검증성공으로 날아오므로 검증성공으로 변경
                        if (gvCheckDetail.GetRowCellValue(i - 1, "Result").ToString() == "검증성공" && specRow != i)
                        {
                            // Camstar 의 E Procedure 를 실행
                            // UDCD Name 이 Task List 로 사용
                            // Task 의 Pass / Fail 적용
                            if (gvCheckDetail.GetRowCellValue(i - 1, "Pass_Fail").ToString() == "Pass")
                            {
                                cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail(udcd_Name, key_no, "True", false);
                            }
                            else
                            {
                                cam_return = WrGlobal.Camster_Common.ExecuteTaskPassFail(udcd_Name, key_no, "False", false);
                            }

                            // Message 칼럼에 cam_return 메시지 적용
                            gvCheckDetail.SetRowCellValue(i - 1, "CamResult", cam_return[0]);
                            gvCheckDetail.SetRowCellValue(i - 1, "CamMessage", cam_return[1]);

                            if (cam_return[1] == "성공!")
                            {
                                // FT_CSV_INSERT 테이블에 원본 데이터 업로드
                                bool bl = work.FT_CSV_INSERT(part, equip.ToString(), rows[i].ToString());

                                // Error 일경우 빠져 나가기
                                if (!bl) { return; }

                                bool bl2;

                                // Spec Row 인 경우 Master 테이블에 Insert
                                if (specRow == i)
                                {
                                    bl2 = work.CmosFtDataIosM(rows[i].ToString(), equip.ToString(), max_1);
                                }
                                // Data Row 인경우 Detail 테이블에 Insert
                                else
                                {
                                    bl2 = work.CmosFtDataIosD(rows[i].ToString(), equip.ToString(), max_1);

                                    // SerialNo 로 Camstar 의 StepEntryTxnId 를 찾고,
                                    // CmosFtDataIosD 테이블에 barcode 와 ID_KEY 가 같은 데이터를 찾아 StepEntryTxnId Update
                                    bool bl3 = work.CmosFtDataIosD_StepEntryTxnId(key_no, max_1);

                                    // Error 일경우 빠져 나가기
                                    if (!bl3) { return; }
                                }

                                // Error 일경우 빠져 나가기
                                if (!bl2) { return; }
                            }


                        }

                        // 진행바 Update
                        progressBarControl1.Position = i;
                        progressBarControl1.Refresh();
                        labelControl1.Text = progressBarControl1.Position.ToString() + " / " + progressBarControl1.Properties.Maximum.ToString();
                        labelControl1.Refresh();
                    }
                }

                SplashScreenManager.CloseForm(false);

                if (!string.IsNullOrEmpty(ErrorMsg))
                {
                    MessageBox.Show(ErrorMsg, "DB Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                SplashScreenManager.CloseForm(false);
                MessageBox.Show(ex.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

            }
        }

        public void RemoveBlankRowsFromCVSFile(string filepath)
        {
            if (filepath == null || filepath.Length == 0)
            {
                throw new ArgumentNullException("filepath");
            }

            if (!File.Exists(filepath))
            {
                throw new FileNotFoundException("Could not find CVS file.", filepath);
            }

            String tempFile = Path.GetTempFileName();

            using (StreamReader reader = new StreamReader(filepath))
            using (StreamWriter writer = new StreamWriter(tempFile))
            {
                String line = null;

                while ((line = reader.ReadLine()) != null)
                {
                    if (!String.IsNullOrWhiteSpace(line.Replace(',', ' ')))
                    {
                        writer.WriteLine(line);
                    }
                }
            }

            File.Delete(filepath);
            File.Move(tempFile, filepath);
        }

        private void frmCMOS_FtDataUpload_Load(object sender, EventArgs e)
        {
            //StreamReader sr = new StreamReader(@"C:\Users\test\Downloads\1.csv", Encoding.GetEncoding("euc-kr"));
            //string line = sr.ReadToEnd();
        }


    }
}