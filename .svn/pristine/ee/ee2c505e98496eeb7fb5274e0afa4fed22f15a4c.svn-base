using DevExpress.DashboardCommon;
using DevExpress.Utils.DragDrop;
using DevExpress.XtraTreeList.Nodes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using VTMES3_RE.Common;
using VTMES3_RE.Models;
using VTMES3_RE.View.Dashboards.Tools;
using VTMES3_RE.View.Reports.Tools;

namespace VTMES3_RE.View.Systems
{
    // 메뉴 설정 폼
    public partial class frmSetMenu : DevExpress.XtraEditors.XtraForm
    {
        clsCode code = new clsCode();

        public frmSetMenu()
        {
            InitializeComponent();
        }

        private void frmSetMenu_Load(object sender, EventArgs e)
        {
            DisplayMenu();
        }
        // 메뉴 구조 TreeList에 바인딩
        private void DisplayMenu()
        {
            this.menuGroupTableAdapter.FillByList(codeDataSet.MenuGroup, WrGlobal.CorpID);
            MenuTreeList.ExpandAll();
        }

        // 저장 버튼 클릭 이벤트
        private void cmdSave_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                this.Validate();
                // 노드에 대한 Level 번호 재정의
                SetNodeLevel(MenuTreeList.Nodes);

                int cnt = 0;

                List<TreeListNode> Listnodes = MenuTreeList.GetNodeList();
                // 메뉴 순번(RowNum) 재정의
                foreach (TreeListNode node in Listnodes)
                {
                    MenuTreeList.SetRowCellValue(node, "RowNum", ++cnt);
                }
                // 입력, 수정 정보 설정
                foreach (DataRowView drv in menuGroupBindingSource)
                {
                    if (drv.Row.RowState == DataRowState.Added)
                    {
                        drv["CreId"] = WrGlobal.LoginID;
                        drv["CreIP"] = WrGlobal.ClientHostName;
                        drv["CreDt"] = DateTime.Now;
                    }
                    else
                    {
                        drv["ModId"] = WrGlobal.LoginID;
                        drv["ModIP"] = WrGlobal.ClientHostName;
                        drv["ModDt"] = DateTime.Now;
                    }
                    //drv["RowNum"] = ++cnt;
                }
                menuGroupBindingSource.EndEdit();

                // 메뉴 변경 내역 저장
                this.menuGroupTableAdapter.Update(codeDataSet.MenuGroup);

                // 신규 폼에대한 DashboardItem 테이블 Row 생성
                code.SetDashboadMenuItem();
                // 메뉴 생성 및 삭제에 대한 권한 항목 생성 및 삭제
                code.SetCodeAuthorByMenu();

                MessageBox.Show("저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 메뉴 추가
        private void cmdInsert_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            try
            {
                // 신규 ROw추가 후 기본 입력값 설정
                menuGroupBindingSource.AddNew();
                DataRowView newItem = (DataRowView)menuGroupBindingSource.Current;

                newItem["CorpId"] = WrGlobal.CorpID;
                newItem["Id"] = code.GetNextSequence("MenuId_Seq").ToString();
                newItem["ParentId"] = "10000";
                newItem["ProjectName"] = WrGlobal.ProJectName;
                newItem["GroupLevel"] = 1;
                newItem["GroupYn"] = "0";
                newItem["ExpandYn"] = "0";
                newItem["PopupYn"] = "N";
                newItem["UseYn"] = "Y";

                newItem["CreDt"] = DateTime.Now;
                newItem["CreId"] = WrGlobal.LoginID;
                newItem["CreIP"] = WrGlobal.ClientHostName;

                menuGroupBindingSource.EndEdit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // nodes 에 대한 Level 정의 및 그룹 여부 설정
        private void SetNodeLevel(TreeListNodes nodes)
        {
            if (nodes.Count == 0) return;

            foreach(TreeListNode node in nodes)
            {
                node.SetValue("GroupLevel", node.Level);
                if (node.Nodes.Count > 0)
                {
                    node.SetValue("GroupYn", "1");
                    SetNodeLevel(node.Nodes);
                }
                else
                {
                    node.SetValue("GroupYn", "0");
                }
            }
        }
        // 선택된 로우 및 폴더 이미지 설정
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
        // 폴더 메뉴 삭제 버튼 숨김, 하위 메뉴 삭제버튼 표시
        private void MenuTreeList_CustomNodeCellEdit(object sender, DevExpress.XtraTreeList.GetCustomNodeCellEditEventArgs e)
        {
            if (e.Column.FieldName == "NodeDel")
            {
                if (e.Node.ParentNode == null)
                {
                    e.RepositoryItem = btnNoDelete;
                }
                else
                {
                    e.RepositoryItem = btnNodeDelete;
                }
            }
        }
        // 노드 삭제 버튼 클릭 이벤트
        private void btnNodeDelete_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            if (MessageBox.Show("선택한 항목을 삭제 하시겠습니까?\r\n하위항목이 존재하는 경우 하위 항목도 같이 삭제됩니다.", "항목삭제", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                List<TreeListNode> nodes = new List<TreeListNode>();
                nodes.Add(MenuTreeList.FocusedNode);
                DropNode(nodes);
            }
        }

        // 노드 삭제 및 하위 노드 삭제
        void DropNode(IEnumerable<TreeListNode> nodes)
        {
            List<TreeListNode> _nodes = new List<TreeListNode>(nodes);
            foreach (TreeListNode node in _nodes)
            {
                if (node.HasChildren)
                    DropNode(node.Nodes);
                DataRowView rowView = MenuTreeList.GetRow(node.Id) as DataRowView;
                if (rowView == null)
                    return;
                MenuTreeList.Nodes.Remove(node);
            }
        }

        private void cmdClose_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            this.Close();
        }
        // 대시보드 작성 버튼 클릭 이벤트
        private void cmdDashboardDesign_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            TreeListNode fNode = MenuTreeList.FocusedNode;
            // 메뉴ID로 대시보드 항목 가져오기
            DataRowView drv = code.IsExistDashboardItem((fNode.GetValue("Id") ?? "").ToString());
            // 저장 전 신규 등록 메뉴는 대시보드를 작성할수 없음
            if (drv == null)
            {
                MessageBox.Show("신규 등록된 메뉴는 저장 후 작성하세요.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 대시보드 폼 오픈
            frmDashBoardDesign form = new frmDashBoardDesign(drv);
            form.ShowDialog();
        }
        // 대시보드 보기 버튼 클릭 이벤트
        private void cmdDashboardView_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            TreeListNode fNode = MenuTreeList.FocusedNode;
            // 메뉴ID로 대시보드 항목 가져오기
            DataRowView drv = code.IsExistDashboardItem((fNode.GetValue("Id") ?? "").ToString());
            // 저장 전 신규 등록 메뉴는 대시보드 보기 할수 없음
            if (drv == null || drv["XML"].ToString() == "")
            {
                MessageBox.Show("대시보드가 작성되지 않았습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 대시보드 폼 오픈
            frmDashBoardView form = new frmDashBoardView(drv);
            form.ShowDialog();
        }
        // 메뉴 조회 버튼 클릭 이벤트 -> 전체 메뉴 재조회
        private void cmdDisplay_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            DisplayMenu();
        }
    }
}