using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.CodeParser;
using DevExpress.DataAccess.Sql;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DevExpress.XtraReports.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using VTMES3_RE.Common;
using static DevExpress.Diagram.Core.Native.Either;
using static DevExpress.Utils.Drawing.Helpers.NativeMethods;

namespace VTMES3_RE.Models
{
    public class clsWork
    {
        Database db = new Database();
        string query = "";

        // CMOS FT DATA 업로드시 원본 데이터를 보관 하기 위함
        public bool FT_CSV_INSERT(string part, string equip, string data)
        {
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_FT_CSV_INSERT N'{0}', N'{1}', N'{2}', '{3}'", part, equip, data, WrGlobal.LoginID));
        }

        // CMOS FT DATA 의 CMOS 데이터 Insert
        public bool CmosFtDataCmos(string datas, string max_1)
        {
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataCmos N'{0}', '{1}', '{2}'", datas, max_1, WrGlobal.LoginID));
        }

        // CMOS 업로드시 ID_KEY 값이 필요함
        // 테이블( CmosFtDataCmos )의 ID_KEY 칼럼의 Max + 1 값을 받아옮
        public String CmosFtDataCmosID_KEY()
        {
            //string query;
            query = string.Format("SELECT ISNULL(MAX(ID_KEY),0) + 1 FROM IFRY.dbo.CmosFtDataCmos");

            DataRowView drv = db.GetDataRecord(query);

            if (drv == null)
            {
                return "";
            }
            else
            {
                return drv[0].ToString();
            }
        }

        // CMOS 업로드시 ID_KEY 값이 필요함
        // 테이블( CmosFtDataCmos )의 ID_KEY 칼럼의 Max + 1 값을 받아옮
        public String getCheckDetail()
        {
            //string query;
            query = string.Format("SELECT ISNULL(DataPointName) FROM IFRY.dbo.CheckDetail");

            DataRowView drv = db.GetDataRecord(query);

            if (drv == null)
            {
                return "";
            }
            else
            {
                return drv[0].ToString();
            }
        }

        // IOS 업로드시 ID_KEY 값이 필요함
        // Master 테이블( CmosFtDataIosM )의 ID_KEY 칼럼의 Max + 1 값을 Return
        public String CmosFtDataIosID_KEY()
        {
            //string query;
            query = string.Format("SELECT ISNULL(MAX(ID_KEY),0) + 1 FROM IFRY.dbo.CmosFtDataIosM");

            DataRowView drv = db.GetDataRecord(query);

            if (drv == null)
            {
                return "";
            }
            else
            {
                return drv[0].ToString();
            }
        }

        // key_no 로 Camstar 의 StepEntryTxnId 를 찾고,
        // CmosFtDataCmos 테이블에 key_no 와 ID_KEY 가 같은 데이터를 찾아 StepEntryTxnId Update
        public bool CmosFtDataCmos_StepEntryTxnId(string key_no, string max_1)
        {
            //return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataCmos_StepEntryTxnId N'{0}', '{1}', '{2}'", key_no, max_1, WrGlobal.LoginID));
            //해당 프로시저에 작업자명 받는 변수 없음
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataCmos_StepEntryTxnId N'{0}', '{1}'", key_no, max_1));
        }

        // key_no 로 Camstar 의 StepEntryTxnId 를 찾고,
        // CmosFtDataIosD 테이블에 key_no 와 ID_KEY 가 같은 데이터를 찾아 StepEntryTxnId Update
        public bool CmosFtDataIosD_StepEntryTxnId(string key_no, string max_1)
        {
            //return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataIosD_StepEntryTxnId N'{0}', '{1}', '{2}'", key_no, max_1, WrGlobal.LoginID));
            //해당 프로시저에 작업자명 받는 변수 없음
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataIosD_StepEntryTxnId N'{0}', '{1}'", key_no, max_1));
        }

        // CMOS FT DATA 의 IOS 의 Master 데이터 Insert
        public bool CmosFtDataIosM(string datas, string equip, string max_1)
        {
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataIosM N'{0}', '{1}', '{2}', '{3}'", datas, equip, max_1, WrGlobal.LoginID));
        }

        // CMOS FT DATA 의 IOS 의 Detail 데이터 Insert
        public bool CmosFtDataIosD(string datas, string equip, string max_1)
        {
            return db.ExecuteQuery(string.Format("exec IFRY.dbo.RY_Proc_CmosFtDataIosD N'{0}', '{1}', '{2}', '{3}'", datas, equip, max_1, WrGlobal.LoginID));
        }

        // key 값 으로 SAP Code 를 찾고, SAP Code 의 상항치, 하한치 값을 가져오기
        public DataView KeynoHighLow(string key_no)
        {
            return db.GetDataView("HighLow", string.Format("exec IFRY.dbo.RY_Proc_KeynoHighLow N'{0}'", key_no));
        }

        // UDCD의 User Data Name, Revision, Name 의 Parameter로, deleteItemByIndex 를 가져오기
        public DataView DeleteItemByIndex(string UserDataName, string Revision, string Name)
        {
            return db.GetDataView("rownum", string.Format("exec IFRY.dbo.RY_Proc_UDCD_DeleteItemByIndex N'{0}', N'{1}', N'{2}'", UserDataName, Revision, Name));
        }

        /// <summary>
        /// 쿼리에 대한 결과값 DataView로 리턴
        /// </summary>
        /// <param name="tblName">테이블명</param>
        /// <param name="_query">쿼리문</param>
        /// <returns></returns>
        public DataView GetDataViewByQuery(string tblName, string _query)
        {
            return db.GetDataView(tblName, _query);
        }
        /// <summary>
        /// 쿼리 실행
        /// </summary>
        /// <param name="_query">쿼리문</param>
        public void ExecuteQry(string _query)
        {
            db.ExecuteQuery(_query);
        }
        /// <summary>
        /// 쿼리 리스트 일괄 실행
        /// </summary>
        /// <param name="queryList">쿼리 목록</param>
        /// <returns></returns>
        public int ExecuteQryList(List<string> queryList)
        {
            return db.ExecuteQueryList(queryList);
        }
        // Camstar 리소스 그룹 목록 조회
        public DataView GetResourceGroup()
        {
            query = string.Format("Select ResourceGroupName 그룹명 "
                                  + "From CAMDBsh.ResourceGroup rg WITH(NOLOCK) "
                                + "INNER JOIN CAMDBsh.FilterTag ft WITH(NOLOCK) ON ft.FilterTagName = 'batch' AND rg.FilterTags LIKE '%' + ft.FilterTagId + '%' "
                                + "Where ResourceGroupName like 'CsI_%' "
                                + "Order by ResourceGroupName"); 
            return db.GetDataView("설비그룹", query);
        }

        // Camstar 리소스 그룹 목록 조회
        public DataView GetResourceGroup2()
        {
            query = string.Format("Select ResourceGroupName 그룹명 "
                            + "From CAMDBsh.ResourceGroup WITH(NOLOCK) Where ResourceGroupName like 'CMOS_%' "
                            + "Order by ResourceGroupName");
            return db.GetDataView("공정그룹", query);
        }

        /// <summary>
        /// 리소스 그룹 내 리소스 목록 조회
        /// </summary>
        /// <param name="groupName">리소스 그룹명</param>
        /// <returns></returns>
        public DataView GetResourceDef(string groupName)
        {
            query = string.Format("SELECT RD.ResourceName 설비명 "
                            + "FROM CAMDBsh.ResourceGroup AS RG WITH(NOLOCK) "
                                + "INNER JOIN CAMDBsh.ResourceGroupEntries AS RGE WITH(NOLOCK) ON RGE.ResourceGroupId = RG.ResourceGroupId "
                                + "INNER JOIN CAMDBsh.ResourceDef AS RD WITH(NOLOCK) ON RD.ResourceId = RGE.EntriesId "
                            + "WHERE RG.ResourceGroupName = N'{0}'", groupName);
            return db.GetDataView("설비", query);
        }

        //Camstar ProductName 
        public DataView GetProductName()
        {
            query = string.Format("SELECT pb.ProductName 제품코드 "
                        + "FROM CAMDBsh.ProductBase pb "
                        + "INNER JOIN IFRY.dbo.MES2_ITEM_MASTER im ON im.ITEM_CODE = pb.ProductName "
                        + "WHERE im.FA_ID = 'CMOS'"
                );
            return db.GetDataView("제품코드", query);
        }

        /// <summary>
        /// 리소스 명으로 시작하는 UDCD 조회
        /// </summary>
        /// <param name="resourceName">리소스명</param>
        /// <returns></returns>
        public DataView GetDataCollection(string resourceName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like N'{0}%' "
                                + "Order By cdb.DataCollectionDefName", resourceName);
            return db.GetDataView("DataCollection", query);
        }

        public DataView GetDataCollection2(string resourceName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like '%{0}%' "
                                + "Order By cdb.DataCollectionDefName", resourceName);
            return db.GetDataView("DataCollection", query);
        }

        public DataView GetDataCollection3(string resourceName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like '{0}' "
                                + "Order By cdb.DataCollectionDefName", resourceName);
            return db.GetDataView("DataCollection", query);
        }

        public DataView GetDataCollection4(string resourceName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like '{0}' "
                                + "Order By cdb.DataCollectionDefName", resourceName);
            return db.GetDataView("DataCollection", query);
        }

        public DataView GetDataCollection4(string resourceName, string ConName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like '%{0}%' "
                                + "AND EXISTS "
                                + "(SELECT ContainerName, ElectronicProcedureName "
                                + "FROM CAMDBsh.Container con "
                                + "INNER JOIN CAMDBsh.CurrentStatus cs ON cs.CurrentStatusId = con.CurrentStatusId "
                                + "INNER JOIN CAMDBsh.Spec ConSpec ON ConSpec.SpecId = cs.SpecId "
                                + "INNER JOIN CAMDBsh.Product p On p.ProductId = con.ProductId "
                                + "INNER JOIN CAMDBsh.BillOfProcessBase bopb ON bopb.BillOfProcessBaseId = p.BillOfProcessBaseId "
                                + "INNER JOIN CAMDBsh.BillOfProcess bop ON bop.BillOfProcessId = CASE WHEN p.BillOfProcessId = '0000000000000000' THEN bopb.RevOfRcdId ELSE p.BillOfProcessId END "
                                + "INNER JOIN CAMDBsh.BillOfProcessOverride bopo ON bopo.BillOfProcessId = bop.BillOfProcessId AND bopo.SpecId = ConSpec.SpecId "
                                + "INNER JOIN CAMDBsh.ElectronicProcedure ep ON ep.ElectronicProcedureId = bopo.ElectronicProcedureId "
                                + "INNER JOIN CAMDBsh.ElectronicProcedureBase epb On ep.ElectronicProcedureBaseId = epb.ElectronicProcedureBaseId "
                                + "WHERE epb.ElectronicProcedureName = cdb.DataCollectionDefName AND ContainerName = '{1}') "
                                + "Order By cdb.DataCollectionDefName", resourceName, ConName);
            return db.GetDataView("DataCollection", query);
        }

        public DataView GetDataCollectionTFT(string resourceName, string ConName)
        {
            query = string.Format("select cdb.DataCollectionDefName 명칭, cd.DataCollectionDefId 코드, cd.DataCollectionDefRevision 리비전 "
                                + "From CAMDBsh.DataCollectionDefBase cdb "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cdb.RevOfRcdId = cd.DataCollectionDefId "
                                + "where cdb.DataCollectionDefName like '%{0}%' "
                                + "AND cdb.DataCollectionDefName Not IN ('CSI_TFT_')"
                                + "AND EXISTS "
                                + "(SELECT ContainerName, ElectronicProcedureName "
                                + "FROM CAMDBsh.Container con "
                                + "INNER JOIN CAMDBsh.CurrentStatus cs ON cs.CurrentStatusId = con.CurrentStatusId "
                                + "INNER JOIN CAMDBsh.Spec ConSpec ON ConSpec.SpecId = cs.SpecId "
                                + "INNER JOIN CAMDBsh.Product p On p.ProductId = con.ProductId "
                                + "INNER JOIN CAMDBsh.BillOfProcessBase bopb ON bopb.BillOfProcessBaseId = p.BillOfProcessBaseId "
                                + "INNER JOIN CAMDBsh.BillOfProcess bop ON bop.BillOfProcessId = CASE WHEN p.BillOfProcessId = '0000000000000000' THEN bopb.RevOfRcdId ELSE p.BillOfProcessId END "
                                + "INNER JOIN CAMDBsh.BillOfProcessOverride bopo ON bopo.BillOfProcessId = bop.BillOfProcessId AND bopo.SpecId = ConSpec.SpecId "
                                + "INNER JOIN CAMDBsh.ElectronicProcedure ep ON ep.ElectronicProcedureId = bopo.ElectronicProcedureId "
                                + "INNER JOIN CAMDBsh.ElectronicProcedureBase epb On ep.ElectronicProcedureBaseId = epb.ElectronicProcedureBaseId "
                                + "WHERE epb.ElectronicProcedureName = cdb.DataCollectionDefName AND ContainerName = '{1}') "
                                + "Order By cdb.DataCollectionDefName", resourceName, ConName);
            return db.GetDataView("DataCollection", query);
        }

        /// <summary>
        /// DataCollectionDefId로 DataPoint 목록 조회
        /// </summary>
        /// <param name="DataCollectionDefId">DataCollectionDefId</param>
        /// <returns></returns>
        public DataView GetDataPointByCollection(string DataCollectionDefId)
        {
            query = string.Format("SELECT dp.DataPointId, dp.DataPointName "
                                    + ",dp.DataType, dp.IsRequired "
                                    + ", ( "
                                        + "SELECT STRING_AGG(og.NamedObjectGroupName, ',') "
                                        + "FROM CAMDBsh.NamedObjectGroupEntries oge "
                                            + "INNER JOIN CAMDBsh.NamedObjectGroup og ON oge.EntriesId = og.NamedObjectGroupId "
                                        + "WHERE oge.NamedObjectGroupId = dp.ObjectGroupId "
                                    + ") NamedObjectGroupName "
                                    + ", cd.INSPECTION_DEFAULT_VALUE dfv "
                                + "FROM CAMDBsh.DataPoint dp "
                                + "INNER JOIN CAMDBsh.DataCollectionDef dcd ON dcd.DataCollectionDefId = dp.DataCollectionDefId "
                                + "INNER JOIN CAMDBsh.DataCollectionDefBase dcdb On dcdb.DataCollectionDefBaseId = dcd.DataCollectionDefBaseId "
                                + "INNER JOIN IFRY.dbo.CheckDetail cd ON cd.DataCollectionDefName = dcdb.DataCollectionDefName AND dcd.DataCollectionDefRevision = cd.DataCollectionDefRevision "
                                + "AND cd.DataPointName = dp.DataPointName "
                                + "WHERE dp.DataCollectionDefId = '{0}' "
                                + "ORDER BY dp.RowPosition, dp.ColumnPosition", DataCollectionDefId);
            return db.GetDataView("DataPoint", query);
        }

        public DataView GetDataPointByCollectionResource(string DataCollectionDefId)
        {
            query = string.Format("SELECT dp.DataPointId, dp.DataPointName "
                                    + ",dp.DataType, dp.IsRequired "
                                    + ", ( "
                                        + "SELECT STRING_AGG(og.NamedObjectGroupName, ',') "
                                        + "FROM CAMDBsh.NamedObjectGroupEntries oge "
                                            + "INNER JOIN CAMDBsh.NamedObjectGroup og ON oge.EntriesId = og.NamedObjectGroupId "
                                        + "WHERE oge.NamedObjectGroupId = dp.ObjectGroupId "
                                    + ") NamedObjectGroupName "
                                + "FROM CAMDBsh.DataPoint dp "
                                + "WHERE dp.DataCollectionDefId = '{0}' "
                                + "ORDER BY dp.RowPosition, dp.ColumnPosition", DataCollectionDefId);
            return db.GetDataView("DataPoint", query);
        }

        /// <summary>
        /// DataCollectionDefId로 해당 Task 목록 조회
        /// </summary>
        /// <param name="DataCollectionDefId">DataCollectionDefId</param>
        /// <returns></returns>
        public DataView GetTaskInfoByCollection(string DataCollectionDefId)
        {
            query = string.Format("select tlb.TaskListName + ' -> ' + ta.TaskItemName TaskName, tlb.TaskListName + '|' + ta.TaskItemName TaskValue "
                            + "From CAMDBsh.TaskItem ta "
                                + "Inner Join CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                + "Inner Join CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId else ta.DataCollectionDefId end "
                            + "Where cd.DataCollectionDefId = '{0}'", DataCollectionDefId);

            return db.GetDataView("TaskList", query);
        }

        public DataView GetTaskInfoByCollection_Origin(string DataCollectionDefId)
        {
            query = string.Format("select tlb.TaskListName + ' -> ' + ta.TaskItemName TaskName, (SELECT TaskListName "
                                   + "FROM CAMDBsh.TaskListBase "
                                   + "WHERE TaskListBaseId IN( "
                                   + "SELECT TaskListBaseId "
                                   + "FROM CAMDBsh.[EProcedureDetail] "
                                   + "WHERE ElectronicProcedureId IN( "
                                   + "SELECT ElectronicProcedureId "
                                   + "FROM[CAMDBsh].[EProcedureDetail] "
                                   + "WHERE TaskListBaseId IN( "
                                   + "SELECT tlb.TaskListBaseId "
                                   + "FROM CAMDBsh.TaskItem ta "
                                   + "INNER JOIN CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                   + "INNER JOIN CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                   + "INNER JOIN CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                   + "INNER JOIN CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId else ta.DataCollectionDefId end "
                                   + "WHERE cd.DataCollectionDefId = '{0}' "
                                   + ") ) AND[Sequence] = 1 "
                                   + ")) +'|' + tlb.TaskListName + '|' + ta.TaskItemName TaskValue "
                                   + "From CAMDBsh.TaskItem ta "
                                   + "Inner Join CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                   + "Inner Join CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                   + "Inner Join CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                   + "Inner Join CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId "
                                   + "else ta.DataCollectionDefId end "
                                   + "Where cd.DataCollectionDefId = '{0}'", DataCollectionDefId);

            return db.GetDataView("TaskList", query);
        }

        /// DataCollectionDefId로 해당 All Task 목록 조회
        public DataView GetAllTaskInfoByCollection(string DataCollectionDefId)
        {
            query = string.Format("SELECT MAX(TaskName) AS TaskName, MAX(TaskListName) + '|' + MAX(TaskValue) AS TaskValue "
                            + "FROM ( "
                                + "select tlb.TaskListName + ' -> ' + ta.TaskItemName TaskName, tlb.TaskListName + '|' + ta.TaskItemName TaskValue, '' as TaskListName "
                                + "From CAMDBsh.TaskItem ta "
                                + "Inner Join CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                + "Inner Join CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId else ta.DataCollectionDefId end "
                                + "Where cd.DataCollectionDefId = '{0}' "
                                + "union all "
                                + "SELECT '' as TaskName, '' as TaskValue, TaskListName "
                                + "FROM CAMDBsh.TaskListBase "
                                + "WHERE TaskListBaseId IN ( "
                                    + "SELECT TaskListBaseId "
                                    + "FROM CAMDBsh.[EProcedureDetail] "
                                    + "WHERE ElectronicProcedureId IN ( "
                                        + "SELECT ElectronicProcedureId "
                                        + "FROM [CAMDBsh].[EProcedureDetail] "
                                        + "WHERE TaskListBaseId IN ( "
                                            + "SELECT tlb.TaskListBaseId "
                                            + "FROM CAMDBsh.TaskItem ta "
                                            + "INNER JOIN CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                            + "INNER JOIN CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                            + "INNER JOIN CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                            + "INNER JOIN CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId else ta.DataCollectionDefId end "
                                            + "WHERE cd.DataCollectionDefId = '{1}' "
                                            + ") "
                                        + ") "
                                        + "AND [Sequence] = 1 "
                                    + ") "
                                + ") A", DataCollectionDefId, DataCollectionDefId);

            return db.GetDataView("TaskList", query);
        }

        public DataView GetTaskInfoByCollection2(string DataCollectionDefId)
        {
            query = string.Format("select tlb.TaskListName + ' -> ' + ta.TaskItemName TaskName, tlb.TaskListName + '|' + ta.TaskItemName TaskValue "
                            + "From CAMDBsh.TaskItem ta "
                                + "Inner Join CAMDBsh.TaskList tl on ta.TaskListId = tl.TaskListId "
                                + "Inner Join CAMDBsh.TaskListBase tlb on tl.TaskListBaseId = tlb.TaskListBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDefBase cdb on ta.DataCollectionDefBaseId = cdb.DataCollectionDefBaseId "
                                + "Inner Join CAMDBsh.DataCollectionDef cd on cd.DataCollectionDefId = case when ta.DataCollectionDefId = '0000000000000000' then cdb.RevOfRcdId else ta.DataCollectionDefId end "
                            + "Where cd.DataCollectionDefId = '{0}'", DataCollectionDefId);

            return db.GetDataView("TaskList", query);
        }
        // 사용안함
        public DataView GetEmployeeWorkTimeDef()
        {
            query = string.Format("Select Gubun 구분 From IFRY.dbo.EmployeeWorkTimeDef Order by SortOrder");

            return db.GetDataView("EmployeeWorkTimeDef", query);
        }
        // 증착 배치번호 조회
        public DataView GetBatchNoDef(string batchno)
        {
            query = string.Format("Select BATCHNO FROM IFRY.dbo.CsiAfterTaskInput where BATCHNO = N'{0}'", batchno);

            return db.GetDataView("CsiAfterTaskInput", query);
        }
        public void GetBatchNoUpdateDef(string batchno,string CL,string CR, string C3, string C4, string Tt, string T5, string SL, string SR, string S3, string S4, string S1, string S2)
        {
            query = string.Format("update IFRY.dbo.CsiAfterTaskInput" +
                " set CsiWeightL = '{0}' , CsiWeightR = '{1}' , TliWeight = '{2}' , ShutterWeightL = '{3}' , ShutterWeightR = '{4}' , SampleThickSPL1 = '{5}' , SampleThickSPL2 = '{6}' " +
                ", CsiWeight3 = '{7}' , CsiWeight4 = '{8}' , TliWeight5 = '{9}' , ShutterWeight3 = '{10}' , ShutterWeight4 = '{11}' where BATCHNO  = N'{12}'", CL,CR,Tt,SL,SR,S1,S2,C3,C4,T5,S3,S4,batchno);

            db.ExecuteQuery(query);
        }
        public DataView GetContainerBatchNo(string container)
        {
            query = string.Format("SELECT DataValue " 
                +"FROM CAMDBsh.RY_EProcedure_TotalRawData ret " 
                +"INNER JOIN CAMDBsh.DataPointHistoryDetail dphd ON dphd.DataPointHistoryId = ret.DataPointHistoryId AND DataName = N'배치번호' " 
                +"WHERE ret.ContainerName = '{0}'", container);

            return db.GetDataView("CsiAfterTaskInput", query);
        }
        public DataView GetDepoBatchNoHistoryDef()
        {
            query = string.Format("Select BatchNo,CsiWeightL,CsiWeightR,CsiWeight3,CsiWeight4,TliWeight,TliWeight5,ShutterWeightL,ShutterWeightR,ShutterWeight3,ShutterWeight4,SampleThickSPL1,SampleThickSPL2 FROM IFRY.dbo.CsiAfterTaskInput");

            return db.GetDataView("CsiAfterTaskInput", query);
        }

        public DataView GetDepoBatchNoHistoryDef2(string start, string end)
        {
            query = string.Format("Select BatchNo,CsiWeightL,CsiWeightR,TliWeight,ShutterWeightL,ShutterWeightR,SampleThickSPL1,SampleThickSPL2 FROM IFRY.dbo.CsiAfterTaskInput" 
                + " WHERE SUBSTRING(BATCHNO,4,8) BETWEEN '{0}' AND '{1}' ",start, end);

            return db.GetDataView("CsiAfterTaskInput", query);
        }
        // 증착 배치 번호 입력
        public void GetBatchNoInsertDef(string batchno)
        {
            query = string.Format("Insert Into IFRY.dbo.CsiAfterTaskInput (BATCHNO) Values ('{0}')", batchno);

            db.ExecuteQuery(query);
        }
        public void GetBatchNoDelteDef(string batchno)
        {
            query = string.Format("Delete From IFRY.dbo.CsiAfterTaskInput where BATCHNO = '{0}'", batchno);

            db.ExecuteQuery(query);
        }
        public void GetBatchNoInsertDef2(string batchno,string CL,string CR, string CR3, string CR4, string Tt, string Tt5, string SL,string SR, string SW3, string SW4, string S1,string S2)
        {
            query = string.Format("Insert Into IFRY.dbo.CsiAfterTaskInput (BATCHNO,CsiWeightL,CsiWeightR," 
                + "TliWeight,ShutterWeightL,ShutterWeightR,SampleThickSPL1,SampleThickSPL2,CsiWeight3,CsiWeight4,TliWeight5,SutterWeight3,SutterWeight4) " 
                + "Values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", batchno,CL,CR,Tt,SL,SR,S1,S2,CR3,CR4,Tt5,SW3,SW4);

            db.ExecuteQuery(query);
        }
        public void GetBatchNoUpdateDef(string batchno, string CL, string CR, string Tt, string SL, string SR, string S1, string S2)
        {
            query = string.Format("Update IFRY.dbo.CsiAfterTaskInput " 
                + "set CsiWeightL = '{0}',CsiWeightR = '{1}', TliWeight = '{2}', ShutterWeightL = '{3}'" 
                + ", ShutterWeightR = '{4}', SampleThickSPL1 = '{5}', SampleThickSPL2 = '{6}'"
                + " where BatchNo = '{7}'", CL, CR, Tt, SL, SR, S1, S2,batchno);

            db.ExecuteQuery(query);
        }
        public DataView GetComponentIssueSelectDef(string Container)
        {
            query = string.Format("select pd.Description + '|' + pb.ProductName + '|' + "
                                 + "case  pr.MaterialTxnLogic when '1' then 'Issue Container (Serial)' "
                                 + "when '2' then 'Issue Container (Lot)'" 
                                 + "when '3' then 'Lot and Stock Point'end as Materials "
                                 + ",fa.FactoryName "
                                 + ",a.ContainerName "
                                 + ",a.Qty "
                                 + ",pr.QtyRequired "
                                 + ",a.Qty * pr.QtyRequired as insertQty "
                                 + ",pb.ProductName "
                                 + "from camdbsh.Container as a "
                                 + "inner join camdbsh.BOM as b on a.BOMId = b.BOMId "
                                 + "inner join camdbsh.ProductMaterialListItem as pr on b.BOMId = pr.BOMId "
                                 + "inner join camdbsh.Product as pd on pr.ProductId = pd.ProductId "
                                 + "inner join camdbsh.ProductBase as pb on pd.ProductBaseId = pb.ProductBaseId "
                                 + "inner join camdbsh.Factory as fa on a.OriginalFactoryId = fa.FactoryId "
                                 + "where a.ContainerName = '{0}'",Container);

            return db.GetDataView("Container", query);
        }
        public DataView GetComponentIssueSelectDef1(string Container)
        {
            query = string.Format("select isnull(max(Qty),0) as Qty from camdbsh.Container where ContainerName = '{0}'", Container);

            return db.GetDataView("Container", query);
        }
        public DataView GetComponentIssueSelectDef2(string product,string Container1,string txn)
        {
            query = string.Format("select "
                                 + "pr.QtyRequired "
                                 + ",a.Qty "
                                 + ",pr.QtyRequired * a.Qty as InsertQty "
                                 + "from camdbsh.Container as a "
                                 + "inner join camdbsh.BOM as b on a.BOMId = b.BOMId "
                                 + "inner join camdbsh.ProductMaterialListItem as pr on b.BOMId = pr.BOMId "
                                 + "inner join camdbsh.Product as pd on pr.ProductId = pd.ProductId "
                                 + "inner join camdbsh.ProductBase as pb on pd.ProductBaseId = pb.ProductBaseId "
                                 + "inner join camdbsh.Factory as fa on a.OriginalFactoryId = fa.FactoryId "
                                 + "where a.ContainerName = '{0}' and pb.ProductName = '{1}' and pr.MaterialTxnLogic ='{2}'", Container1, product,txn);

            return db.GetDataView("Container", query);
        }
        public DataView GetComponentIssueSelectDef3(string product, string Container1)
        {
            query = string.Format("select "
                                   + "pr.MaterialTxnLogic "
                                   + "from camdbsh.Container as a "
                                   + "inner join camdbsh.BOM as b on a.BOMId = b.BOMId "
                                   + "inner join camdbsh.ProductMaterialListItem as pr on b.BOMId = pr.BOMId "
                                   + "inner join camdbsh.Product as pd on pr.ProductId = pd.ProductId "
                                   + "inner join camdbsh.ProductBase as pb on pd.ProductBaseId = pb.ProductBaseId "
                                   + "inner join camdbsh.Factory as fa on a.OriginalFactoryId = fa.FactoryId "
                                   + "where a.ContainerName = '{0}' "
                                   + "and pb.ProductName = '{1}'", Container1,product);

            return db.GetDataView("Container", query);
        }
        // 정상 근무 정보 조회
        public DataView GetEmployeeWorkTimeRegularDef()
        {
            query = string.Format("Select Gubun, StartTime, WorkHour From IFRY.dbo.EmployeeWorkTimeDef Where RegularYn = 'Y'");

            return db.GetDataView("EmployeeWorkTimeRegularDef", query);
        }
        // 근무 구분 조회 
        public String GetEmployeeWorkTimeDefString()
        {
            query = string.Format("Select STRING_AGG(Gubun, ',') WITHIN GROUP(ORDER BY SortOrder) Gubun From IFRY.dbo.EmployeeWorkTimeDef");

            DataRowView drv = db.GetDataRecord(query);

            if (drv == null)
            {
                return "";
            }
            else
            {
                return drv[0].ToString();
            }
        }
        // 공장에 대한 사용자 팀리스트(전체 포함)
        public DataView GetEmployeeTeamList()
        {
            query = string.Format("exec CAMDBsh.RY_VR_Proc_Common_EmployeeTeamList N'{0}'", WrGlobal.FactoryName);

            return db.GetDataView("EmployeeTeamList", query);
        }

        // 공장, 팀에 대한 사용자 리스트 (전체 포함)
        public DataView GetEmployeeListByTeam(string teamName)
        {
            query = string.Format("exec CAMDBsh.RY_VR_Proc_Common_EmployeeListByTeam N'{0}', N'{1}'", WrGlobal.FactoryName, teamName);

            return db.GetDataView("EmployeeList", query);
        }
        // 점검년월 콤보 바인딩
        public DataView GetMachineCheckSheetYear()
        {
            query = string.Format("Select YEAR(GETDATE()) CheckYear Union Select CheckYear From IFRY.dbo.MachineCheckSheet group by CheckYear Order By CheckYear");
            return db.GetDataView("CheckYear", query);
        }
        // 사용 안함
        public XtraReport GetSelectedReport(string url)
        {
            // Return a report by a URL selected in the ListBox.
            if (url == "")
                return null;
            using (MemoryStream stream = new MemoryStream(Program.ReportStorage.GetData(url)))
            {
                
                return XtraReport.FromStream(stream, true);
            }
        }
        // 신규 SapCode 등록 프로시저
        public void MES2_ITEM_MASTER_insert()
        {
            query = string.Format("exec CAMDBsh.RY_VM_Proc_MES2_ITEM_MASTER_insert");
            //db.ExecuteQuery(query);
        }

        // PQC 데이터 불러오기
        public DataView GetPQCDataLoad(string ConName)
        {
            query = string.Format("exec CAMDBsh.RY_VR_Proc_TFT_FT_DATA N'{0}' ", ConName);
            return db.GetDataView("CheckContainer", query);
        }

        // PQC 데이터 불러오기 (23.06.19 김문철) : Ft_Upload 데이터 불러오기 추가
        public DataTable GetCmosPQCDataLoad(string ConName)
        {
            query = string.Format("exec CAMDBsh.RY_VR_Proc_CMOS_FT_DATA_LOAD N'{0}' ", ConName);
            DataSet ds = db.GetDataSet("GetCmosPQCDataLoad", query);

            DataRow[] ftRows = ds.Tables[0].Select("AUTO_YN = 'Y'");

            if (ftRows.Length > 0)
            {
                string[] display_areas = null;
                string display_area = "";

                foreach (DataRow ftRow in ftRows)
                {
                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        display_areas = ftRow["DISPLAY_AREA"].ToString().ToUpper().Split('/');
                        display_area = "";

                        foreach (DataRow cmosRow in ds.Tables[1].Rows)
                        {
                            display_area = display_areas[0].Trim();

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
                            ftRow["VALUE_RESULT_NAME"] = cmosRow[col].ToString();
                        }
                    }

                    if (ds.Tables[2].Rows.Count > 0)
                    {
                        display_areas = ftRow["DISPLAY_AREA"].ToString().ToUpper().Split('/');
                        display_area = "";

                        foreach (DataRow iosRow in ds.Tables[2].Rows)
                        {
                            display_area = display_areas[Convert.ToInt32(iosRow["HoGi"]) - 1].Trim();
                            int col = 0;

                            if (display_area.Length == 1)
                            {
                                if (Convert.ToInt32(iosRow["SoonBun"]) != 1) continue;
                                char c = display_area[0];
                                col = Convert.ToInt32(c) - 65;
                            }
                            else if (display_area.Length == 2)
                            {
                                char c1 = display_area[0];
                                col = (Convert.ToInt32(c1) - 64) * 26;

                                char c2 = display_area[1];
                                col += Convert.ToInt32(c2) - 65;

                                if (col > 26)
                                {   // 순번 : 2 아니면 continue
                                    if (Convert.ToInt32(iosRow["SoonBun"]) != 2) continue;
                                    col -= 27;
                                }
                                else
                                {   // 순번 : 1 아니면 continue
                                    if (Convert.ToInt32(iosRow["SoonBun"]) != 1) continue;
                                }
                            }

                            ftRow["VALUE_RESULT_NAME"] = iosRow[col].ToString();
                        }
                    }
                }
            }

            return ds.Tables[0];
        }

        // OQC 데이터 불러오기
        public DataView GetQCDataLoad(string DataCollectionDefId)
        {
            query = string.Format("exec CAMDBsh.RY_VR_Proc_TFT_QC_DATA N'{0}' ", DataCollectionDefId);
            return db.GetDataView("DataCollectionDefId", query);
        }

        //GetFTContainerName SN로 컨테이너 넘버 가져오기
        public DataView GetFTContainerName(string SerialNumber)
        {
            query = string.Format("select Con.ContainerName AS AY_NAME, " +
                                    " att.AttributeValue ,REPLACE(Con.ContainerName,'AY','QC') AS QC_NAME " +
                                    " from CAMDBsh.Container as Con " +
                                    " INNER JOIN (select * from CAMDBsh.UserAttribute where UserAttributeName = '300.PRODUCT_SN') as att on Con.ContainerId = att.ParentId" +
                                    " WHERE Con.Status = 1 and att.AttributeValue = N'{0}' ", SerialNumber);

            return db.GetDataView("ContainerName", query);
        }


    }
}
