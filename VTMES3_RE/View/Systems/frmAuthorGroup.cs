using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using VTMES3_RE.Common;
using VTMES3_RE.Models;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors.ViewInfo;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.CodeParser;

namespace VTMES3_RE.View.Systems
{
    public partial class frmAuthorGroup : DevExpress.XtraEditors.XtraForm
    {
        // clsCode 모델 생성
        clsCode user = new clsCode();

        public frmAuthorGroup()
        {
            InitializeComponent();

        }

        // 권한 그룹 바인딩
        private void frmAuthorGroup_Load(object sender, EventArgs e)
        {
            //if (WrGlobal.AuthorList.Contains((this.Tag ?? "").ToString() + "02"))
            //{
            //    cmdInsert.Visible = true;
            //    cmdSave.Visible = true;
            //    cmdDelete.Visible = true;
            //}

            this.codeAuthorGroupTableAdapter.Fill(this.codeDataSet.CodeAuthorGroup, WrGlobal.CorpID);
        }

        // 신규 권한 그룹 등록 이벤트 -> 신규 입력 로우 추가후 기본값 설정
        private void cmdInsert_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            string newItemYn = "N";
            foreach (DataRowView item in codeAuthorGroupBindingSource)
            {
                if (item["GroupCode"].ToString() == "")
                {
                    newItemYn = "Y";
                    break;
                }
            }

            if (newItemYn == "Y") return;

            try
            {
                codeAuthorGroupBindingSource.AddNew();
                DataRowView newItem = (DataRowView)codeAuthorGroupBindingSource.Current;

                newItem["CorpId"] = WrGlobal.CorpID;
                newItem["GroupCode"] = "";
                newItem["GroupName"] = "";

                newItem["CreDt"] = DateTime.Now;
                newItem["CreId"] = WrGlobal.LoginID;
                newItem["CreIP"] = WrGlobal.ClientHostName;

                codeAuthorGroupBindingSource.EndEdit();

                gvCodeAuthorGroup.MoveLast();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // 권한 그룹 변경시 권한 리스트에 부여된 권한 항목 체크 표시
        private void gvCodeAuthorGroup_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (gvCodeAuthorGroup.FocusedRowHandle < 0) return;

            this.menuGroupTableAdapter.FillByAuthorGroup(codeDataSet.MenuGroup, WrGlobal.CorpID, gvCodeAuthorGroup.GetRowCellValue(e.FocusedRowHandle, "GroupCode").ToString());
            MenuTreeList.ExpandAll();
        }

        // 저장 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();

                foreach (DataRowView drv in codeAuthorGroupBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        drv["GroupCode"] = user.GetNewAuthorGroupCode();
                    }
                    else if (drv.Row.RowState == DataRowState.Modified)
                    {
                        drv["ModDt"] = DateTime.Now;
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModIP"] = WrGlobal.ClientHostName;
                    }
                }
                // 권한 그룹 저장
                codeAuthorGroupBindingSource.EndEdit();
                this.codeAuthorGroupTableAdapter.Update(this.codeDataSet.CodeAuthorGroup);

               
                List<TreeListNode> Listnodes = MenuTreeList.GetNodeList();
                List<string> queryList = new List<string>();

                // 권한 항목 테이블(CodeAuthorGroupDetail) 에서 해당 권한 그룹 전체 삭제
                queryList.Add(string.Format("DELETE FROM {0}_ReportDB.dbo.CodeAuthorGroupDetail WHERE CorpId = '{0}' and GroupCode = '{1}'", WrGlobal.CorpID, gvCodeAuthorGroup.GetFocusedRowCellValue("GroupCode").ToString()));

                // 트리 노드별 체크된 항목 CodeAuthorGroupDetail 에 Insert 쿼리문  queryList에 추가
                foreach (TreeListNode node in Listnodes)
                { 
                    if (node.GetValue("GroupYn").ToString() == "1") continue;

                    if (node.GetValue("Author1").ToString() == "1")
                    {   // 조회 권한 Insert
                        queryList.Add(string.Format("INSERT INTO {0}_ReportDB.dbo.CodeAuthorGroupDetail(CorpId, GroupCode, MenuID, AuthorCode, UseYn, CreId, CreIP, CreDt) VALUES("
                                        + "'{0}', '{1}', {2}, '{3}', 'Y', '{4}', '{5}', getdate())",
                                        WrGlobal.CorpID, gvCodeAuthorGroup.GetFocusedRowCellValue("GroupCode").ToString(), node.GetValue("Id"), node.GetValue("Code1"), WrGlobal.LoginID, WrGlobal.ClientHostName));
                    }

                    if (node.GetValue("Author2").ToString() == "1")
                    {   // 수정 권한 Insert 
                        queryList.Add(string.Format("INSERT INTO {0}_ReportDB.dbo.CodeAuthorGroupDetail(CorpId, GroupCode, MenuID, AuthorCode, UseYn, CreId, CreIP, CreDt) VALUES("
                                        + "'{0}', '{1}', {2}, '{3}', 'Y', '{4}', '{5}', getdate())",
                                        WrGlobal.CorpID, gvCodeAuthorGroup.GetFocusedRowCellValue("GroupCode").ToString(), node.GetValue("Id"), node.GetValue("Code2"), WrGlobal.LoginID, WrGlobal.ClientHostName));
                    }
                }

                // queryList의 권한 항목 쿼리문 일괄 처리
                if (user.ExecuteAuthorGroupDetailQueryList(queryList))
                {
                    MessageBox.Show("저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                this.menuGroupTableAdapter.FillByAuthorGroup(codeDataSet.MenuGroup, WrGlobal.CorpID, gvCodeAuthorGroup.GetFocusedRowCellValue("GroupCode").ToString());
                MenuTreeList.ExpandAll();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 권한 그룹 삭제 이벤트
        private void cmdDelete_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            if (codeAuthorGroupBindingSource.Current == null) return;

            if (MessageBox.Show("선택한 자료를 삭제 하시겠습니까?", "삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            try
            {
                DataRowView drv = (DataRowView)codeAuthorGroupBindingSource.Current;

                string groupcode = drv["GroupCode"].ToString();

                // 권한 그룹의 권한 항목 삭제 -> CodeAuthorGroupDetail 테이블
                user.ExecuteAuthorGroupDetailDeleteByGroupCode(groupcode);
                // 권한 그룹 삭제 처리
                codeAuthorGroupBindingSource.RemoveCurrent();
                codeAuthorGroupTableAdapter.Update(codeDataSet.CodeAuthorGroup);

                MessageBox.Show("삭제되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }

        // 선택된 로우 및 루트 폴더 이미지 설정
        private void MenuTreeList_GetStateImage(object sender, DevExpress.XtraTreeList.GetStateImageEventArgs e)
        {
            if (e.Node != null)
            {
                if (e.Node.Selected)
                {
                    if (e.Node.HasChildren)
                    {
                        e.NodeImageIndex = 3;
                    }
                    else
                    {
                        e.NodeImageIndex = 4;
                    }//end if
                }
                else
                {
                    if (e.Node.HasChildren)
                    {
                        e.NodeImageIndex = 1;
                    }
                    else
                    {
                        e.NodeImageIndex = 0;
                    }//end if
                }//end if

            }//end if
        }
        // 조회 체크 및 해재, 폴더 체크/ 해제시 하위 항목 동일 적용
        private void chkAuthor1_CheckedChanged(object sender, EventArgs e)
        {
            string groupYn = (MenuTreeList.GetFocusedRowCellValue("GroupYn") ?? "").ToString();
            if (groupYn != "1") return;

            TreeListNode node = MenuTreeList.FindNodeByKeyID((MenuTreeList.GetFocusedRowCellValue("Id") ?? ""));
            SetNodeCheck(node.Nodes, "Author1", ((CheckEdit)sender).Checked);
            node.SetValue("Author1", ((CheckEdit)sender).Checked ? "1" : "0");
        }
        // 편집 체크 및 해재, 폴더 체크/ 해제시 하위 항목 동일 적용
        private void chkAuthor2_CheckedChanged(object sender, EventArgs e)
        {
            string groupYn = (MenuTreeList.GetFocusedRowCellValue("GroupYn") ?? "").ToString();
            if (groupYn != "1") return;

            TreeListNode node = MenuTreeList.FindNodeByKeyID((MenuTreeList.GetFocusedRowCellValue("Id") ?? ""));
            SetNodeCheck(node.Nodes, "Author2", ((CheckEdit)sender).Checked);
            node.SetValue("Author2", ((CheckEdit)sender).Checked ? "1" : "0");
        }

        /// <summary>
        /// 선택된 노드 및 하위 노드 체크/해제 적용
        /// </summary>
        /// <param name="nodes">작업할 노드</param>
        /// <param name="colName">컬럼명</param>
        /// <param name="isChecked">체크여부</param>
        private void SetNodeCheck(TreeListNodes nodes, string colName, bool isChecked)
        {
            if (nodes.Count == 0) return;

            foreach (TreeListNode node in nodes)
            {
                if (node.GetValue("GroupYn").ToString() == "1")
                {
                    node.SetValue(colName, isChecked ? "1" : "0");
                    SetNodeCheck(node.Nodes, colName, isChecked);
                }
                else
                {
                    node.SetValue(colName, isChecked ? "1" : "0");
                }
            }
        }
    }
}