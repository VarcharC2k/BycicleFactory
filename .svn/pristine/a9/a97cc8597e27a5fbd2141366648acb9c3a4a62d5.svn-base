using DevExpress.XtraEditors.Repository;
using VTMES3_RE.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraEditors.TextEditController.Win32;
using System.Reflection;

namespace VTMES3_RE.Models
{
    public class clsCode
    {
        Database db = new Database();
        string query = "";
        // 사용 안함
        public DataView GetAuthorViewByGroupCode(string groupcode)
        {
            query = string.Format("exec {0}_ReportDB.dbo.Code_GetCodeAuthorByGroupCode '{0}', '{1}'", WrGlobal.CorpID, groupcode);

            return db.GetDataView("사용자권한", query);
        }
        // 신규 권한 그룹 코드 생성
        public string GetNewAuthorGroupCode()
        {
            query = string.Format("SELECT CONVERT(NVARCHAR(4), ISNULL(MAX(GroupCode), 1000) + 1) GroupCode FROM {0}_ReportDB.dbo.CodeAuthorGroup Where CorpId = '{0}'", WrGlobal.CorpID);
            DataRowView drv = db.GetDataRecord(query);

            return drv["GroupCode"].ToString();
        }

        /// <summary>
        /// 권한 그룹 삭제 처리
        /// </summary>
        /// <param name="groupcode">권한그룹 코드</param>
        public void ExecuteAuthorGroupDetailDeleteByGroupCode(string groupcode)
        {
            // 권한 그웁의 권한 항목 삭제
            db.ExecuteQuery(string.Format("DELETE FROM {0}_ReportDB.dbo.CodeAuthorGroupDetail WHERE CorpId = '{0}' and GroupCode = '{1}'", WrGlobal.CorpID, groupcode));
            // 사용자에게 정의된 권한 그룹에서 삭제된 권한 그룹 제거
            db.ExecuteQuery(string.Format("UPDATE {0}_ReportDB.dbo.CodeUser SET AuthorGroup = REVERSE(STUFF(REVERSE(Replace(Replace(AuthorGroup, ' ', '') + ',', '{1},', '')), 1, 1, '')) Where CorpId = '{0}'", WrGlobal.CorpID, groupcode));
        }
        // 로그인자의 메뉴 리스트 가져와 DataView로 리턴
        public DataView GetEmployeeMenuList()
        {
            query = string.Format("exec {0}_ReportDB.dbo.Code_GetEmployeeMenuList '{0}', '{1}', '{2}'", WrGlobal.CorpID, WrGlobal.LoginID, WrGlobal.ProgramType);
            return db.GetDataView("MenuGroup", query);
        }

        /// <summary>
        /// 사용자 정보 가져오기
        /// </summary>
        /// <param name="uid">로그인 ID</param>
        /// <returns></returns>
        public DataRowView GetUserInfo(string uid)
        {
            query = string.Format("SELECT cu.*, emp.EmployeeId, emp.FullName, fa.FactoryId, fa.FactoryName, rg.ESigRoleGroupName "
                    + "FROM {0}_ReportDB.dbo.CodeUser cu "
                        + "Inner Join CAMDBsh.Employee emp on cu.EmployeeName = emp.EmployeeName "
                        + "Inner Join CAMDBsh.SessionValues sv on emp.EmployeeId = sv.EmployeeId "
                        + "LEFT Join CAMDBsh.Factory fa on sv.FactoryId = fa.FactoryId "
                        + "LEFT JOIN CAMDBsh.ESigRoleGroup rg ON emp.ESigRoleGroupId = rg.ESigRoleGroupId "
                    + "WHERE cu.CorpId = '{0}' and cu.EmployeeName = '{1}'", WrGlobal.CorpID, uid);
            return db.GetDataRecord(query);
        }
        /// <summary>
        /// 사용자 추가
        /// </summary>
        /// <param name="uid">사용자ID</param>
        /// <param name="upw">비밀번호</param>
        /// <returns></returns>
        public DataRowView AddUser(string uid, string upw)
        {
            query = string.Format("Insert Into {0}_ReportDB.dbo.CodeUser(CorpId, EmployeeName, Password) Values('{0}', '{1}', '{2}')", WrGlobal.CorpID, uid, clsCommon.SHA256Hash(upw));
            db.ExecuteQuery(query);

            return GetUserInfo(uid);
        }
        // 사용자 정보 가져오기, Lookupedit 에 바인딩용
        public DataView GetUser()
        {
            query = string.Format("Select Employee.EmployeeName ID, Employee.FullName 사용자명, Factory.FactoryName 사업부 "
                            + "From CAMDBsh.Employee "
                                + "Inner Join CAMDBsh.SessionValues on Employee.EmployeeId = SessionValues.EmployeeId "
                                + "LEFT Join CAMDBsh.Factory on SessionValues.FactoryId = Factory.FactoryId "
                            + "Where Factory.FactoryName in('Rayence') OR Factory.FactoryName IS NULL "
                            + "Order By Factory.FactoryName, Employee.FullName");
            return db.GetDataView("사용자", query);
        }

        // 로그인 자의 권한 리스트 설정
        public void SetUserAuthor()
        {
            // 권한 그룹 가져오기
            query = string.Format("SELECT AuthorGroup FROM {0}_ReportDB.dbo.CodeUser WHERE CorpId = '{0}' and EmployeeName = '{1}'", WrGlobal.CorpID, WrGlobal.LoginID);
            DataRowView drv = db.GetDataRecord(query);

            if (drv[0] == DBNull.Value) return;

            drv[0] = drv[0].ToString().Replace(" ", "");

            string[] arrGroup = drv[0].ToString().Split(new char[] { ',' });
            string inGroupStr = "";
            // 권한 그룹을 '그룹코드1','그룹코드2' 로 재배열
            foreach (string item in arrGroup)
            {
                if (item == "") continue;

                if (inGroupStr != "") inGroupStr = inGroupStr + ",";
                inGroupStr = inGroupStr + "''" + item + "''";
            }
            // 권한 목록 가져오기 프로시저 호출
            query = string.Format("exec {0}_ReportDB.dbo.Code_GetAuthorByAuthorGroups '{0}', '{1}', '{2}'", WrGlobal.CorpID, WrGlobal.ProgramType, inGroupStr);
            DataView dv = db.GetDataView("권한", query);

            // 권한 목록 WrGlobal.AuthorList에 등록
            foreach (DataRowView view in dv)
            {
                clsAuthor author = new clsAuthor(view["MenuID"].ToString(), view["ParentID"].ToString(), view["AuthorCode"].ToString());
                WrGlobal.AuthorList.Add(author);
            }
        }
        // 사용자의 팀명 설정
        public void SetEmployeeTeam()
        {
            WrGlobal.TeamName = "";
            query = string.Format("exec {0}_ReportDB.dbo.Code_GetEmployeeTeamList '{0}', '{1}'", WrGlobal.CorpID, WrGlobal.LoginID);
            DataView dv = db.GetDataView("TeamList", query);

            foreach (DataRowView drv in dv)
            {
                if (WrGlobal.TeamName != "") WrGlobal.TeamName += ",";
                WrGlobal.TeamName += "'" + drv["FilterTagName"].ToString() + "'";    
            }
        }
        // 권한 그룹 정보 가져오기
        public DataView GetGroupAuthorList()
        {
            query = string.Format("SELECT GroupCode, GroupName FROM {0}_ReportDB.dbo.CodeAuthorGroup Where CorpId = '{0}' ORDER BY GroupName", WrGlobal.CorpID);
            return db.GetDataView("CodeAuthorGroup", query);
        }

        /// <summary>
        /// 비밀번호 변경 
        /// </summary>
        /// <param name="frpwd">기존비밀번호</param>
        /// <param name="topwd">신규비밀번호</param>
        /// <returns></returns>
        public bool ChangePassword(string frpwd, string topwd)
        {
            bool bRet = false;

            query = string.Format("SELECT * FROM {0}_ReportDB.dbo.CodeUser WHERE CorpId = '{0}' and EmployeeName = '{1}'", WrGlobal.CorpID, WrGlobal.LoginID);

            DataRowView drv = db.GetDataRecord(query);

            if (clsCommon.SHA256Hash(frpwd) == clsCommon.getString(drv["Password"]))
            {
                query = string.Format("UPDATE {0}_ReportDB.dbo.CodeUser SET Password = '{2}' WHERE CorpId = '{0}' and EmployeeName = '{1}'", WrGlobal.CorpID, WrGlobal.LoginID, clsCommon.SHA256Hash(topwd));
                bRet = db.ExecuteQuery(query);
            }
            else
            {
                bRet = false;
            }

            return bRet;
        }

        /// <summary>
        /// 권한 항목 등록 쿼리 일괄 실행
        /// </summary>
        /// <param name="list">쿼리리스트</param>
        /// <returns></returns>
        public bool ExecuteAuthorGroupDetailQueryList(List<string> list)
        {
            return db.ExecuteQueryList(list) > 0 ? true : false;
        }
        /// <summary>
        /// 메뉴 생성 시퀀스의 다음 시퀀스 번호 가져오기
        /// </summary>
        /// <param name="seqName"></param>
        /// <returns></returns>
        public int GetNextSequence(string seqName)
        {
            int nextSeq = 1;
            query = string.Format("SELECT NEXT VALUE FOR {0}_ReportDB.dbo.{1}", WrGlobal.CorpID, seqName);
            DataRowView drv = db.GetDataRecord(query);

            if (drv != null)
            {
                nextSeq = Convert.ToInt32(drv[0]);
            }

            return nextSeq;
        }
        // 신규 대시보드 생성 
        public void SetDashboadMenuItem()
        {
            query = string.Format("Insert Into {0}_ReportDB.dbo.DashBoardItem(CorpId, MenuId, CreId, CreIP, CreDt) "
                            + "Select mg.CorpId, mg.Id MenuId, '{1}', HOST_NAME(), getdate() "
                                + "From {0}_ReportDB.dbo.MenuGroup mg "
                                    + "Left Join {0}_ReportDB.dbo.DashBoardItem bi on mg.CorpId = bi.CorpId and mg.Id = bi.MenuId "
                                + "Where mg.CorpId = '{0}' and bi.MenuId is null", 
                                WrGlobal.CorpID, WrGlobal.LoginID);

            db.ExecuteQuery(query);
        }
        /// <summary>
        /// 기존 대시보드 Row 가져오기
        /// </summary>
        /// <param name="menuId">메뉴ID</param>
        /// <returns></returns>
        public DataRowView IsExistDashboardItem(string menuId)
        {
            query = string.Format("Select bi.*, mg.Title_ko MenuName, pmg.Title_ko ParentMenuName "
                                + "From {0}_ReportDB.dbo.DashBoardItem bi "
                                    + "Inner Join {0}_ReportDB.dbo.MenuGroup mg on bi.CorpId = mg.CorpId and bi.MenuId = mg.Id "
                                    + "Left Join {0}_ReportDB.dbo.MenuGroup pmg on mg.CorpId = pmg.CorpId and mg.ParentId = pmg.Id "
                                + "Where bi.CorpId = '{0}' and bi.MenuId = '{1}'",
                            WrGlobal.CorpID, menuId);

            return db.GetDataRecord(query);
        }
        // 메뉴 목록 추가 및 삭제에 대하여 권한 항목 테이블 재설정 
        public void SetCodeAuthorByMenu()
        {
            query = string.Format("exec {0}_ReportDB.dbo.Code_SetCodeAuthorByMenu '{0}'",
                                WrGlobal.CorpID);

            db.ExecuteQuery(query);
        }
        //세션 생성
        public void CreateSession()
        {
            WrGlobal.SessionNo = DateTime.Now.ToString("yyyyMMddHHmmss");

            query = string.Format("exec {0}_ReportDB.dbo.Code_CreateSession '{0}', N'{1}', '{2}', '{3}'", WrGlobal.CorpID, WrGlobal.LoginID, WrGlobal.SessionNo, "CS");
            db.ExecuteQuery(query);
        }
        //세션 종료
        public void CloseSession()
        {
            query = string.Format("exec {0}_ReportDB.dbo.Code_CloseSession '{0}', N'{1}', '{2}'", WrGlobal.CorpID, WrGlobal.LoginID, WrGlobal.SessionNo);
            db.ExecuteQuery(query);
        }

        /// <summary>
        /// 기능 폼 오픈 내역 저장
        /// </summary>
        /// <param name="menuId">오픈 메뉴ID</param>
        /// <param name="useType">구분값</param>
        public void UseSession(string menuId, string useType)
        {
            query = string.Format("exec {0}_ReportDB.dbo.Code_UseSession '{0}', N'{1}', '{2}', '{3}', N'{4}'", WrGlobal.CorpID, WrGlobal.LoginID, WrGlobal.SessionNo, menuId, useType);
            db.ExecuteQuery(query);
        }
    }
}
