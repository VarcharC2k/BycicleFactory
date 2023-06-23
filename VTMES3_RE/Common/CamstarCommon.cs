using Camstar.XMLClient.API;
using DevExpress.CodeParser;
using DevExpress.DashboardWin.Native;
using DevExpress.Pdf.Native.BouncyCastle.Ocsp;
using DevExpress.Xpo.DB.Helpers;
using DevExpress.XtraMap.Native;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraReports.Serialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VTMES3_RE.Models;
using static DevExpress.Data.Filtering.Helpers.SubExprHelper.ThreadHoppingFiltering;

namespace VTMES3_RE.Common
{
    public class CamstarCommon
    {
        clsWork work = new clsWork();

        csiClient gClient = new csiClient();
        csiConnection gConnection = null;
        csiSession gSession = null;
        csiDocument gDocument = null;
        csiService gService = null;
        string gStrSessionID = "";

        string gHost = WrGlobal.Camstar_Host;
        int gPort = WrGlobal.Camstar_Port;

        CamstarMessage camMessage = new CamstarMessage();

        public bool IsExecuting = false;

        string query = "";
        CamstarDatabase db = new CamstarDatabase();

        // Camstar Connection 설정
        public CamstarCommon()
        {
            try
            {
                gConnection = gClient.createConnection(gHost, gPort);
                //gSession = gConnection.createSession(gUserName, gPassword, gSessionID.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Camstar", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Camstar 서비스 생성
        public void CreateDocumentandService(string documentName, string serviceName)
        {
            if (documentName.Length > 0 && gSession != null)
            {
                gSession.removeDocument(documentName);
            }

            if (gService != null)
            {
                gService = null;
            }

            gDocument = gSession.createDocument(documentName);

            if (serviceName.Length > 0)
            {
                gService = gDocument.createService(serviceName);
            }
        }
        // Request, Response XML 파일 설정
        public void PrintDoc(string strDoc, bool isInputDoc)
        {
            string strDocFileName = "";
            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            if (isInputDoc)
            {
                strDocFileName = "inputDoc.xml";
            }
            else
            {
                strDocFileName = "responseDoc.xml";
            }

            if (File.Exists(path + "\\" + strDocFileName))
            {
                File.Delete(path + "\\" + strDocFileName);
            }

            File.WriteAllText(path + "\\" + strDocFileName, strDoc, Encoding.Default);
        }
        // 에러 체크
        private void ErrorsCheck(csiDocument ResponseDocument)
        {
            csiExceptionData csiexceptiondata;

            if (ResponseDocument.checkErrors())
            {   // 에러
                camMessage.Result = false;
                csiexceptiondata = ResponseDocument.exceptionData();
                camMessage.Message = "Severity: " + csiexceptiondata.getSeverity() + " Description: " + csiexceptiondata.getDescription();
            }
            else
            {   // 정상
                camMessage.Result = true;
                camMessage.Message = "성공!";
            }
        }
        // 신규 세션 생성
        public string CreateSession()
        {
            string PW = string.Empty;
            try
            {
                gStrSessionID = Guid.NewGuid().ToString();
                if(WrGlobal.LoginID.ToString() == "Administrator")
                {
                    PW = "Dentalimageno.1";
                }
                else
                {
                    PW = WrGlobal.LoginID.Substring(0, 1) + "!@" + WrGlobal.LoginID.Substring(WrGlobal.LoginID.Length - 5);
                }
                gSession = gConnection.createSession(WrGlobal.LoginID, PW, gStrSessionID);

                return "";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string DestroySession()
        {
            try
            {
                gConnection.removeSession(gStrSessionID);

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        // 세션 삭제
        public void DestroyConnection()
        {
            gConnection.removeSession(gStrSessionID);
            gClient.removeConnection(gHost, gPort);
        }

        #region Role Function
        public CamstarMessage RoleDelete(string employeeName, int idx)
        {
            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiNamedSubentityList Roles = null;

            try
            {
                CreateDocumentandService("EmployeeMaintTrans", "EmployeeMaint");

                InputData = gService.inputData();
                InputData.namedObjectField("ObjectToChange").setRef(employeeName);

                gService.perform("Load");

                InputData2 = gService.inputData();

                ObjectChanges = InputData2.subentityField("ObjectChanges");
                Roles = ObjectChanges.namedSubentityList("Roles");

                Roles.deleteItemByIndex(idx);

                gService.setExecute();
                gService.requestData().requestField("CompletionMsg");

                PrintDoc(gDocument.asXML(), true);
                ResponseDocument = gDocument.submit();
                PrintDoc(ResponseDocument.asXML(), false);
                ErrorsCheck(ResponseDocument);
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            return camMessage;
        }

        public CamstarMessage RoleAdd(string employeename, string roleName)
        {
            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiSubentity Members = null;

            try
            {
                CreateDocumentandService("EmployeeMaintTrans", "EmployeeMaint");

                InputData = gService.inputData();
                InputData.namedObjectField("ObjectToChange").setRef(employeename);

                gService.perform("Load");

                InputData2 = gService.inputData();

                ObjectChanges = InputData2.subentityField("ObjectChanges");
                Members = ObjectChanges.subentityList("Roles").appendItem();
                Members.namedObjectField("Role").setRef(roleName);
                Members.dataField("PropagateToChildOrgs").setValue(false.ToString());

                gService.setExecute();
                gService.requestData().requestField("CompletionMsg");

                PrintDoc(gDocument.asXML(), true);
                ResponseDocument = gDocument.submit();
                PrintDoc(ResponseDocument.asXML(), false);
                ErrorsCheck(ResponseDocument);
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            return camMessage;
        }
        #endregion

        #region Modeling Import
        // Modeling BOP
        public int BOP_Modeling_Import(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData1 = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiSubentityList BillOfProcessOverrides = null;
            csiSubentity listItem = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("BillOfProcessMaintDoc", "BillOfProcessMaint");

                    //Set InputData 1
                    InputData1 = gService.inputData();
                    InputData1.dataField("SyncName").setValue(dr["BOP Name"].ToString());
                    InputData1.dataField("SyncRevision").setValue(dr["Revision"].ToString());

                    // Set eventName
                    gService.perform("Sync");
                    //gService.perform("NEW")

                    //' Set InputData2
                    InputData2 = gService.inputData();

                    //' Set ObjectChanges
                    ObjectChanges = InputData2.subentityField("ObjectChanges");
                    ObjectChanges.dataField("Name").setValue(dr["BOP Name"].ToString());
                    ObjectChanges.dataField("Revision").setValue(dr["Revision"].ToString());
                    ObjectChanges.dataField("Filtertags").setValue(dr["Filtertag"].ToString());

                    //' Set Data Points
                    BillOfProcessOverrides = ObjectChanges.subentityList("BillOfProcessOverrides");

                    //' Set ListItem Loop
                    listItem = BillOfProcessOverrides.appendItem();

                    //'Set E-Procedure
                    if (dr["E-Procedure Rev"].ToString() == "")
                    {

                        listItem.revisionedObjectField("ElectronicProcedure").setRef(dr["E-Procedure"].ToString(), "", true);
                    }
                    else
                    {
                        listItem.revisionedObjectField("ElectronicProcedure").setRef(dr["E-Procedure"].ToString(), dr["E-Procedure Rev"].ToString(), false);
                    }

                    //' Set Name → 이거 리비전 어케해야될지 모르겠음....; OOTB랑 상의
                    listItem.dataField("Name").setValue(dr["Spec"].ToString());

                    //' Set Spec
                    if (dr["Spec Rev"].ToString() == "")
                    {
                        listItem.revisionedObjectField("Spec").setRef(dr["Spec"].ToString(), "", true);
                    }
                    else
                    {
                        listItem.revisionedObjectField("Spec").setRef(dr["Spec"].ToString(), dr["Spec Rev"].ToString(), false);
                    }

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XML Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);
                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        //if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        if (csiservice != null)
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            //dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        //dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }
        
        // Modeling UDCD Value
        public int UserDataCollectDef_ValueDataPoint(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;            
            csiObject InputData1 = null;
            csiObject InputData2 = null;            
            csiSubentity ObjectChanges = null;
            csiSubentityList DataPoints = null;
            csiSubentity listItem = null;            

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("UserDataCollectionDefMaintDoc", "UserDataCollectionDefMaint");

                    //Set InputData 1
                    InputData1 = gService.inputData();
                    InputData1.dataField("SyncName").setValue(dr["User Data Name"].ToString());
                    InputData1.dataField("SyncRevision").setValue(dr["Revision"].ToString());

                    // Set eventName
                    gService.perform("Sync");
                    //gService.perform("NEW")

                    //' Set InputData2
                    InputData2 = gService.inputData();

                    //' Set ObjectChanges
                    ObjectChanges = InputData2.subentityField("ObjectChanges");
                    ObjectChanges.dataField("FilterTags").setValue(dr["Filter Tags"].ToString());
                    ObjectChanges.dataField("Name").setValue(dr["User Data Name"].ToString());
                    ObjectChanges.dataField("Revision").setValue(dr["Revision"].ToString());
                    ObjectChanges.dataField("DataPointLayout").setValue(dr["DataPointLayout"].ToString());

                    //' Set Data Points
                    DataPoints = ObjectChanges.subentityList("DataPoints");

                    //' Set ListItem Loop
                    listItem = DataPoints.appendItem();
                    listItem.setObjectType("ValueDataPointChanges");
                    listItem.dataField("Name").setValue(dr["Name"].ToString());
                    listItem.dataField("RowPosition").setValue(dr["RowPosition"].ToString());
                    listItem.dataField("ColumnPosition").setValue(dr["ColumnPosition"].ToString());
                    listItem.dataField("DataType").setValue(dr["DataType"].ToString());
                    if (dr["DataType"].ToString() == "7")
                    {
                        listItem.dataField("BooleanTrue").setValue(dr["BooleanTrue"].ToString());
                        listItem.dataField("BooleanFalse").setValue(dr["BooleanFalse"].ToString());
                    }
                    listItem.dataField("IsRequired").setValue(dr["IsRequired"].ToString());
                    if(dr["LowerLimit"].ToString() != "")
                    {
                        listItem.dataField("LowerLimit").setValue(dr["LowerLimit"].ToString());
                    }
                    if(dr["UpperLimit"].ToString() != ""){
                        listItem.dataField("UpperLimit").setValue(dr["UpperLimit"].ToString());
                    }                    
                    listItem.dataField("IsLimitOverrideAllowed").setValue(dr["IsLimitOverrideAllowed"].ToString());
                    listItem.dataField("MapToUserAttribute").setValue(dr["MapToUserAttribute"].ToString());
                    listItem.dataField("AttributeName").setValue(dr["AttributeName"].ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XML Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);
                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        //if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        if (csiservice != null)
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            //dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        //dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch(Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        // Modeling UDCD Value - Delete
        public int UserDataCollectDef_DataPoint_Delete(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData1 = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiSubentityList DataPoints = null;
            csiSubentityList Attributes = null;
            csiSubentity listItem = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Data Point Name 으로 Index 가져오기
                    DataView dv = new DataView();
                    dv = work.DeleteItemByIndex(dr["User Data Name"].ToString(), dr["Revision"].ToString(), dr["Name"].ToString());

                    //if (dv.Table.Rows[0][1] == DBNull.Value)
                    //{
                    //    string a = "";
                    //}

                    if (dv.Count > 0)
                    {
                        // Set Service Type
                        CreateDocumentandService("UserDataCollectionDefMaintDoc", "UserDataCollectionDefMaint");

                        //Set InputData 1
                        InputData1 = gService.inputData();
                        InputData1.revisionedObjectField("ObjectToChange").setRef(dr["User Data Name"].ToString(), dr["Revision"].ToString(), false);

                        // Set eventName
                        gService.perform("Load");
                        //gService.perform("Sync");
                        //gService.perform("NEW")

                        //' Set InputData2
                        InputData2 = gService.inputData();

                        //' Set ObjectChanges
                        ObjectChanges = InputData2.subentityField("ObjectChanges");

                        // 'Delete Data Point Name
                        if(dv.Table.Rows[0][1] == DBNull.Value)
                        {
                            string a = "";
                        }
                        else
                        {
                            int ItemIndex = Convert.ToInt32(dv.Table.Rows[0][1].ToString());
                            DataPoints = ObjectChanges.subentityList("DataPoints");
                            DataPoints.deleteItemByIndex(ItemIndex);
                        }                        

                        // Service Excute and request Completion Msg
                        gService.setExecute();
                        gService.requestData().requestField("CompletionMsg");

                        // Print XML Dcoument
                        PrintDoc(gDocument.asXML(), true);
                        ResponseDocument = gDocument.submit();
                        PrintDoc(ResponseDocument.asXML(), false);
                        ErrorsCheck(ResponseDocument);

                        dr.BeginEdit();
                        if (camMessage.Result)
                        {
                            csiService csiservice = ResponseDocument.getService();
                            //if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                            if (csiservice != null)
                            {
                                csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                                //dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                            }
                            successCnt++;
                        }
                        else
                        {
                            //dr["Container"] = "";
                        }
                        dr["Result"] = camMessage.Message;
                        dr["BoolResult"] = camMessage.Result;
                        dr.EndEdit();
                    }
                    else
                    {
                        dr.BeginEdit();
                        dr["Result"] = "데이터 없음";
                        dr["BoolResult"] = false;
                        dr.EndEdit();
                    }
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        // Modeling UDCD Object
        public int UserDataCollectDef_ObjectDataPoint(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData1 = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiSubentityList DataPoints = null;
            csiSubentity listItem = null;
            csiNamedSubentity ObjectGroup = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("UserDataCollectionDefMaintDoc", "UserDataCollectionDefMaint");

                    //Set InputData 1
                    InputData1 = gService.inputData();
                    InputData1.dataField("SyncName").setValue(dr["User Data Name"].ToString());
                    InputData1.dataField("SyncRevision").setValue(dr["Revision"].ToString());

                    // Set eventName
                    gService.perform("Sync");
                    //gService.perform("NEW")

                    //' Set InputData2
                    InputData2 = gService.inputData();

                    //' Set ObjectChanges
                    ObjectChanges = InputData2.subentityField("ObjectChanges");
                    ObjectChanges.dataField("FilterTags").setValue(dr["Filter Tags"].ToString());
                    ObjectChanges.dataField("DataPointLayout").setValue(dr["DataPointLayout"].ToString());

                    //' Set Data Points
                    DataPoints = ObjectChanges.subentityList("DataPoints");

                    //' Set ListItem Loop
                    listItem = DataPoints.appendItem();
                    listItem.setObjectType("ObjectDataPointChanges");
                    listItem.dataField("Name").setValue(dr["Name"].ToString());
                    listItem.dataField("RowPosition").setValue(dr["RowPosition"].ToString());
                    listItem.dataField("ColumnPosition").setValue(dr["ColumnPosition"].ToString());
                    listItem.dataField("DataType").setValue(dr["DataType"].ToString());
                    listItem.dataField("IsRequired").setValue(dr["IsRequired"].ToString());
                    listItem.dataField("DisplayMode").setValue(dr["DisplayMode"].ToString());
                    ObjectGroup = listItem.namedSubentityField("ObjectGroup");
                    ObjectGroup.setObjectType("NamedObjectGroup");
                    ObjectGroup.setName(dr["ObjectGroup"].ToString());
                    listItem.dataField("ObjectSelValType").setValue(dr["ObjectSelValType"].ToString());
                    listItem.dataField("ObjectType").setValue("5330");
                    ObjectChanges.dataField("Name").setValue(dr["User Data Name"].ToString());
                    ObjectChanges.dataField("Revision").setValue(dr["Revision"].ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XML Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);
                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        //if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        if (csiservice != null)
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            //dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        //dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int UserDataCollectDef_ResourceDataPoint(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData1 = null;
            csiObject InputData2 = null;
            csiSubentity ObjectChanges = null;
            csiSubentityList DataPoints = null;
            csiSubentity listItem = null;
            csiNamedSubentity ObjectGroup = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("UserDataCollectionDefMaintDoc", "UserDataCollectionDefMaint");

                    //Set InputData 1
                    InputData1 = gService.inputData();
                    InputData1.dataField("SyncName").setValue(dr["User Data Name"].ToString());
                    InputData1.dataField("SyncRevision").setValue(dr["Revision"].ToString());

                    // Set eventName
                    gService.perform("Sync");
                    //gService.perform("NEW")

                    //' Set InputData2
                    InputData2 = gService.inputData();

                    //' Set ObjectChanges
                    ObjectChanges = InputData2.subentityField("ObjectChanges");
                    ObjectChanges.dataField("FilterTags").setValue(dr["Filter Tags"].ToString());
                    ObjectChanges.dataField("DataPointLayout").setValue(dr["DataPointLayout"].ToString());

                    //' Set Data Points
                    DataPoints = ObjectChanges.subentityList("DataPoints");

                    //' Set ListItem Loop
                    listItem = DataPoints.appendItem();
                    listItem.setObjectType("ObjectDataPointChanges");
                    listItem.dataField("Name").setValue(dr["Name"].ToString());
                    listItem.dataField("RowPosition").setValue(dr["RowPosition"].ToString());
                    listItem.dataField("ColumnPosition").setValue(dr["ColumnPosition"].ToString());
                    listItem.dataField("DataType").setValue(dr["DataType"].ToString());
                    listItem.dataField("IsRequired").setValue(dr["IsRequired"].ToString());
                    listItem.dataField("DisplayMode").setValue(dr["DisplayMode"].ToString());
                    ObjectGroup = listItem.namedSubentityField("ObjectGroup");
                    //ObjectGroup.setObjectType("NamedObjectGroup");
                    ObjectGroup.setObjectType("ResourceGroup");
                    ObjectGroup.setName(dr["ObjectGroup"].ToString());
                    listItem.dataField("ObjectSelValType").setValue(dr["ObjectSelValType"].ToString());
                    listItem.dataField("ObjectType").setValue("5350");
                    ObjectChanges.dataField("Name").setValue(dr["User Data Name"].ToString());
                    ObjectChanges.dataField("Revision").setValue(dr["Revision"].ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XML Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);
                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        //if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        if (csiservice != null)
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            //dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        //dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        #endregion

        #region Container Function

        public int ContainerStart(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiSubentity Details = null;
            csiSubentity CurrentStatusDetails = null;



            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("StartDoc", "Start");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set CurrentStatusDetails
                    CurrentStatusDetails = InputData.subentityField("CurrentStatusDetails");
                    if (dr["Workflow Rev"].ToString() == "")
                    {
                        CurrentStatusDetails.revisionedObjectField("Workflow").setRef((dr["Workflow"] ?? "").ToString(), "", true);
                    }
                    else
                    {
                        CurrentStatusDetails.revisionedObjectField("Workflow").setRef((dr["Workflow"] ?? "").ToString(), (dr["Workflow Rev"] ?? "").ToString(), true);
                    }

                    // Set Start Details
                    Details = InputData.subentityField("Details");

                    //Set Auto Container Name
                    if ((dr["Container"] ?? "").ToString() == "Auto")
                    {
                        Details.dataField("AutoNumber").setValue("True");
                        Details.dataField("IsContainer").setValue("True");
                    }
                    else
                    {
                        Details.dataField("ContainerName").setValue((dr["Container"] ?? "").ToString());
                    }

                    // Set Start Element
                    Details.namedObjectField("Owner").setRef((dr["Owner"] ?? "").ToString());
                    Details.dataField("Qty").setValue((dr["Qty"] ?? "0").ToString());
                    Details.namedObjectField("StartReason").setRef((dr["StartReason"] ?? "").ToString());
                    Details.namedObjectField("UOM").setRef((dr["UOM"] ?? "").ToString());
                    Details.namedObjectField("Level").setRef((dr["Level"] ?? "").ToString());
                    Details.namedObjectField("PriorityCode").setRef((dr["PriorityCode"] ?? "").ToString());
                    Details.namedObjectField("MfgOrder").setRef((dr["MfgOrder"] ?? "").ToString());
                    if (dr["Product Rev"].ToString() == "")
                    {
                        Details.revisionedObjectField("Product").setRef((dr["Product"] ?? "").ToString(), "", true);
                    }
                    else
                    {
                        Details.revisionedObjectField("Product").setRef((dr["Product"] ?? "").ToString(), (dr["Product Rev"] ?? "").ToString(), true);

                    }
                    Details.dataField("ContainerComment").setValue((dr["Comments"] ?? "").ToString());

                    // Set Factory 
                    InputData.namedObjectField("Factory").setRef((dr["Factory"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int ComponentIssue(DataTable table, string InsertContainer, DataTable table1, string[] product, string[] insertQty)
        {
            int successCnt = 0;
            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiSubentity CalledByTransactionTask = null;
            csiSubentity IssueActualDetails = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                // 엑셀 칼럼의 투입자재를 동적으로 가져오기 위함
                int colMoc = Convert.ToInt32(System.Math.Truncate(Convert.ToDouble(table.Columns.Count) / 3.5));
                int i = 0;
                foreach (DataRow dr in table.Rows)
                {
                   // colMoc 만큼 반복하여 CAMSTAR 로 제출
                   for (int colidx = 0; colMoc > colidx; colidx++)
                   {
                       // 투입자재에 데이터가 있을때만 실행
                       if (dr.Table.Rows[i][colidx * 3 + 3].ToString().Length > 0)
                       {
                           string[] pro = new string[20];
                           string[] issue = new string[20];
                           string[] total = new string[20];
                           string qty = string.Empty;

                           total = dr.Table.Rows[i][colidx * 3 + 2].ToString().Split(new char[] { '|' });

                           pro[0] = total[1].ToString();
                           issue[0] = total[2].ToString();

                           if (issue[0].ToString() == "Issue Container (Serial)")
                           {
                               issue[0] = '1'.ToString();
                           }
                           else if (issue[0].ToString() == "Issue Container (Lot)")
                           {
                               issue[0] = '2'.ToString();
                           }
                           else if (issue[0].ToString() == "Lot and Stock Point")
                           {
                               issue[0] = '3'.ToString();
                           }

                           string Container = dr["Container"].ToString();

                           // 자식의 Container 로 Product 가져오기 - 사용 X
                           // parameter : 자식 Container / 결과 : 자식 SAP Code
                           //DataView dv = work.GetComponentIssueSelectDef1(dr.Table.Rows[0][colidx * 3 + 2].ToString());

                           // parameter : 부모 Container, 자식 SAP Code, issue Control 타입( datatable 의 투입자재x 의 배열변수에 3번쨰 항목에 있음 ) / 
                           //DataView dv2 = work.GetComponentIssueSelectDef3(dv.Table.Rows[0][0].ToString(), dr["Container"].ToString());

                           // LOT : 부모 Container Qty( table에 있음 ) x 부모 BOM의 자식 Product 의 Required Qty
                           DataView dv1 = work.GetComponentIssueSelectDef2(pro[0].ToString(), Container, issue[0].ToString());


                           // Serial : 자식 Container 의 Qty 값 >> parameter: 자식 Container / 결과 : 자식 Container Qty
                           DataView dv2 = work.GetComponentIssueSelectDef1(dr.Table.Rows[i][colidx * 3 + 3].ToString());

                            if (dv2.Table.Rows[0][0].ToString() == "0")
                            {
                                dr.Table.Rows[i][colidx * 3 + 4] = "투입 자재 컨테이너명을 확인해주세요";
                                continue;

                            }
                            //자식 컨테이너에 QtyRequired 와 부모 컨테이너 Qty 비교

                            //if (Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) <= Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                            //{
                            //    qty = dv1.Table.Rows[0][0].ToString();
                            //}
                            //else if (Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) > Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                            //{
                            //    qty = dv2.Table.Rows[0][0].ToString();
                            //}

                            CreateDocumentandService("ComponentIssueTrans", "ComponentIssue");

                           if (issue[0].ToString() == '3'.ToString())
                           {

                               InputData = gService.inputData();

                               CalledByTransactionTask = InputData.subentityField("CalledByTransactionTask");
                               CalledByTransactionTask.setObjectType("InstructionItem");

                               InputData.namedObjectField("Container").setRef(Container);
                               InputData.namedObjectField("Factory").setRef("Rayence");

                               IssueActualDetails = InputData.subentityField("IssueActualDetails").subentityField("__listItem");
                               IssueActualDetails.dataField("FromLot").setValue(dr.Table.Rows[i][colidx * 3 + 3].ToString());

                               IssueActualDetails.revisionedObjectField("Product").setRef(pro[0].ToString(), "", true);

                               IssueActualDetails.dataField("QtyIssued").setValue(dv1.Table.Rows[0][0].ToString());

                           }
                           else
                           {

                               InputData = gService.inputData();

                               CalledByTransactionTask = InputData.subentityField("CalledByTransactionTask");
                               CalledByTransactionTask.setObjectType("InstructionItem");

                               InputData.namedObjectField("Container").setRef(Container);
                               InputData.namedObjectField("Factory").setRef("Rayence");

                               IssueActualDetails = InputData.subentityField("IssueActualDetails").subentityField("__listItem");
                               IssueActualDetails.namedObjectField("FromContainer").setRef(dr.Table.Rows[i][colidx * 3 + 3].ToString());

                               // 부모자재 컨테이너에 issue control 값이 Serial이고 부모자재 qty required 값이 투입자재 qty보다 작을 경우
                               if (issue[0].ToString() == '1'.ToString() && Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) < Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv1.Table.Rows[0][0].ToString());
                                   //IssueActualDetails.namedObjectField("IssueDifferenceReason").setRef("");
                               }
                               // 부모자재 컨테이너에 issue control 값이 Serial이고 부모자재 qty required 값이 투입자재 qty보다 클 경우
                               else if (issue[0].ToString() == '1'.ToString() && Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) > Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv2.Table.Rows[0][0].ToString());
                                   //IssueActualDetails.namedObjectField("IssueDifferenceReason").setRef("");
                               }
                               // 부모자재 컨테이너에 issue control 값이 Serial이고 부모자재 qty required 값이 투입자재 qty보다 같을 경우
                               else if (issue[0].ToString() == '1'.ToString() && Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) == Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv1.Table.Rows[0][0].ToString());
                               }

                               // 부모자재 컨테이너에 issue control 값이 LOT이고 부모자재 qty required 값이 투입자재 qty보다 작거나 같을 경우
                               else if (issue[0].ToString() == '2'.ToString() && Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) <= Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv1.Table.Rows[0][0].ToString());
                               }
                               // 부모자재 컨테이너에 issue control 값이 LOT이고 부모자재 qty required 값이 투입자재 qty보다 클 경우
                               else if (issue[0].ToString() == '2'.ToString() && Convert.ToInt32(dv1.Table.Rows[0][0].ToString()) > Convert.ToInt32(dv2.Table.Rows[0][0].ToString()))
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv2.Table.Rows[0][0].ToString());
                                   //IssueActualDetails.namedObjectField("IssueDifferenceReason").setRef("");
                               }
                               else if (issue[0].ToString() == '4'.ToString())
                               {
                                   IssueActualDetails.dataField("QtyIssued").setValue(dv1.Table.Rows[0][0].ToString());
                                   //IssueActualDetails.namedObjectField("IssueDifferenceReason").setRef("");
                               }

                               IssueActualDetails.revisionedObjectField("Product").setRef(pro[0].ToString(), "", true);

                               InputData.namedObjectField("TaskContainer").setRef(Container);
                           }


                           gService.setExecute();
                           gService.requestData().requestField("CompletionMsg");

                           PrintDoc(gDocument.asXML(), true);
                           ResponseDocument = gDocument.submit();
                           PrintDoc(ResponseDocument.asXML(), false);

                           ErrorsCheck(ResponseDocument);

                           dr.BeginEdit();
                           if (camMessage.Result)
                           {
                               csiService csiservice = ResponseDocument.getService();
                               if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                               {
                                   csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                                   dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                               }
                               successCnt++;
                           }
                           else
                           {
                               //dr["Container"] = "";
                           }


                           dr.Table.Rows[i][colidx * 3 + 4] = camMessage.Message;
                           dr.EndEdit();
                       }


                   }
                    i++;
                    
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int StartTwoLevel(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiSubentity Details = null;
            csiSubentity CurrentStatusDetails = null;
            csiSubentity ChildContainers = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("StartDoc", "Start");

                    // Set Input Data
                    InputData = gService.inputData();

                    // Set CurrentStatusDetails
                    CurrentStatusDetails = InputData.subentityField("CurrentStatusDetails");
                    if (dr["Workflow Rev"].ToString() == "")
                    {
                        CurrentStatusDetails.revisionedObjectField("Workflow").setRef((dr["Workflow"] ?? "").ToString(), "", true);
                    }
                    else
                    {
                        CurrentStatusDetails.revisionedObjectField("Workflow").setRef((dr["Workflow"] ?? "").ToString(), (dr["Workflow Rev"] ?? "").ToString(), true);

                    }

                    // Set Start Details
                    Details = InputData.subentityField("Details");

                    // Set Container Details
                    if ((dr["Container"] ?? "").ToString() == "Auto")
                    {
                        Details.dataField("AutoNumber").setValue("True");
                    }
                    else
                    {
                        Details.dataField("ContainerName").setValue((dr["Container"] ?? "").ToString());
                    }
                    Details.namedObjectField("Level").setRef((dr["Level"] ?? "").ToString());
                    Details.namedObjectField("Owner").setRef((dr["Owner"] ?? "").ToString());
                    Details.namedObjectField("StartReason").setRef((dr["StartReason"] ?? "").ToString());
                    Details.namedObjectField("MfgOrder").setRef((dr["MfgOrder"] ?? "").ToString());
                    if (dr["Product Rev"].ToString() == "")
                    {
                        Details.revisionedObjectField("Product").setRef((dr["Product"] ?? "").ToString(), "", true);
                    }
                    else
                    {
                        Details.revisionedObjectField("Product").setRef((dr["Product"] ?? "").ToString(), (dr["Product Rev"] ?? "").ToString(), true);

                    }
                    Details.namedObjectField("PriorityCode").setRef((dr["PriorityCode"] ?? "").ToString());
                    Details.dataField("ContainerComment").setValue((dr["Comments"] ?? "").ToString());

                    InputData.namedObjectField("Factory").setRef((dr["Factory"] ?? "").ToString());

                    // Set Child Container info
                    Details.dataField("ChildAutoNumber").setValue("True");
                    Details.dataField("ChildCount").setValue((dr["ChildCount"] ?? "0").ToString());
                    Details.dataField("DefaultChildQty").setValue((dr["ChildQty"] ?? "0").ToString());

                    // Set CildContainers 
                    Int32 ChildCount = Int32.Parse((dr["ChildCount"]).ToString());
                    ;

                    for (int i = 1; i <= ChildCount; i++)
                    {
                        ChildContainers = Details.subentityList("ChildContainers").appendItem();
                        ChildContainers.dataField("ContainerName").setValue("");
                        ChildContainers.namedObjectField("Level").setRef((dr["ChildLevel"] ?? "").ToString());
                        ChildContainers.dataField("Qty").setValue((dr["ChildQty"] ?? "0").ToString());
                        ChildContainers.namedObjectField("UOM").setRef((dr["ChildUOM"] ?? "").ToString());
                    }

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        dr["Container"] = "";
                    }

                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int PrintContainerLabel(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("PrintContainerLabelDoc", "PrintContainerLabel");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.dataField("LabelCount").setValue((dr["Label Count"] ?? "").ToString());

                    if (dr["Printer Label Rev"].ToString() == "")
                    {
                        InputData.revisionedObjectField("PrinterLabelDefinition").setRef((dr["Printer Label Definition"] ?? "").ToString(), "", true);
                    }
                    else
                    {
                        InputData.revisionedObjectField("PrinterLabelDefinition").setRef((dr["Printer Label Definition"] ?? "").ToString(), (dr["Printer Label Rev"] ?? "").ToString(), true);
                    }
                    InputData.namedObjectField("PrintQueue").setRef((dr["Print Queue"] ?? "").ToString());
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        dr["Container"] = "";
                    }

                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public CamstarMessage ContainerHoldLoop(string containerName, string holdReasonName)
        {
            csiDocument ResponseDocument = null;
            csiObject InputData = null;

            try
            {
                CreateDocumentandService("HoldDoc", "Hold");

                InputData = gService.inputData();

                InputData.namedObjectField("Container").setRef(containerName);
                InputData.namedObjectField("HoldReason").setRef(holdReasonName);

                gService.setExecute();
                gService.requestData().requestField("CompletionMsg");

                PrintDoc(gDocument.asXML(), true);
                ResponseDocument = gDocument.submit();
                PrintDoc(ResponseDocument.asXML(), false);

                ErrorsCheck(ResponseDocument);
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            return camMessage;
        }

        public CamstarMessage ContainerReleaseLoop(string containerName, string releaseReasonName)
        {
            csiDocument ResponseDocument = null;
            csiObject InputData = null;

            try
            {
                CreateDocumentandService("ReleaseDoc", "Release");

                InputData = gService.inputData();

                InputData.namedObjectField("Container").setRef(containerName);
                InputData.namedObjectField("ReleaseReason").setRef(releaseReasonName);

                gService.setExecute();
                gService.requestData().requestField("CompletionMsg");

                PrintDoc(gDocument.asXML(), true);
                ResponseDocument = gDocument.submit();
                PrintDoc(ResponseDocument.asXML(), false);

                ErrorsCheck(ResponseDocument);
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            return camMessage;
        }


        public int ContainerAttribute(DataTable table)
        {
            int successCnt = 0;

            List<string> arrList = new List<string>();

            int startAttrIdx = table.Columns.IndexOf("Container");
            int endAttrIdx = table.Columns.IndexOf("Attribute Result");

            if (startAttrIdx > -1 || endAttrIdx > -1)
            {
                startAttrIdx = startAttrIdx + 1;
                endAttrIdx = endAttrIdx - 1;

                for (int i = startAttrIdx; i < endAttrIdx + 1; i++)
                {
                    arrList.Add(table.Columns[i].ColumnName);
                }
            }

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiSubentityList ServiceDetails = null;
            csiSubentity listItem = null;
            csiSubentity Attribute = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("ContainerAttrMaintDoc", "ContainerAttrMaint");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    ServiceDetails = InputData.subentityList("ServiceDetails");

                    // Load Product Attribute
                    query = string.Format("Select ua.UserAttributeName, ua.AttributeValue "
                                        + "From (Select * From CAMDBsh.Container Where ContainerName = N'{0}') con "
                                            + "Inner Join CAMDBsh.UserAttribute ua on con.ContainerId = ua.ParentId", (dr["Container"] ?? "").ToString());
                    DataView conAttrDv = db.GetDataView("containerAttribute", query);

                    foreach (DataRowView drv in conAttrDv)
                    {
                        listItem = ServiceDetails.appendItem();
                        Attribute = listItem.subentityField("Attribute");
                        Attribute.setObjectType("UserAttribute");
                        listItem.dataField("Name").setValue((drv["UserAttributeName"] ?? "").ToString());

                        if (arrList.Contains((drv["UserAttributeName"] ?? "").ToString()))
                        {
                            if (dr[(drv["UserAttributeName"] ?? "").ToString()].ToString() == "")
                            {
                                listItem.dataField("AttributeValue").setValue(drv["AttributeValue"].ToString());        
                            }    
                            else
                            {
                                listItem.dataField("AttributeValue").setValue(dr[(drv["UserAttributeName"] ?? "").ToString()].ToString());
                            }
                        }
                        else
                        {
                            listItem.dataField("AttributeValue").setValue(drv["AttributeValue"].ToString());
                        }
                        listItem.dataField("DataType").setValue("4");
                        listItem.dataField("IsExpression").setValue("False");
                    }

                    foreach (string colName in arrList)
                    {
                        conAttrDv.RowFilter = string.Format("UserAttributeName = '{0}'", colName);

                        if (conAttrDv.Count > 0) continue;

                        listItem = ServiceDetails.appendItem();
                        Attribute = listItem.subentityField("Attribute");
                        Attribute.setObjectType("UserAttribute");
                        listItem.dataField("Name").setValue(colName);
                        listItem.dataField("AttributeValue").setValue(dr[colName].ToString());
                        listItem.dataField("DataType").setValue("4");
                        listItem.dataField("IsExpression").setValue("False");
                    }

                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        successCnt++;
                    }

                    dr["Attribute Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
  
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int Close(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiSubentity Details = null;
            csiSubentity CurrentStatusDetails = null;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("CloseDoc", "Close");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Container Name
                    InputData.namedObjectField("Container").setRef(dr["Container"].ToString());

                    // Set Factory 
                    InputData.namedObjectField("Factory").setRef("Rayence");

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        //dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        #endregion

        #region RY_ExcuteTask
        public int WorkStart(string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity CalledByTransactionTask;
            csiParentInfo TaskList = null;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("MoveInDoc", "MoveIn");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Called By Transaction Task
                    CalledByTransactionTask = InputData.namedSubentityField("CalledByTransactionTask");
                    CalledByTransactionTask.setName("작업시작");
                    CalledByTransactionTask.setObjectType("InstructionItem");
                    TaskList = CalledByTransactionTask.parentInfo();
                    TaskList.setRevisionedObjectRef(taskListName, "", true);

                    InputData.dataField("ClearLocation").setValue("False");
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.namedObjectField("Resource").setRef(resourceName);
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    dr["StartResult"] = camMessage.Message;
                    dr["StartBoolResult"] = camMessage.Result;
                    //dr["Result"] = camMessage.Message;
                    //dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        public int WorkStartResult(string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity CalledByTransactionTask;
            csiParentInfo TaskList = null;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("MoveInDoc", "MoveIn");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Called By Transaction Task
                    CalledByTransactionTask = InputData.namedSubentityField("CalledByTransactionTask");
                    CalledByTransactionTask.setName("작업시작");
                    CalledByTransactionTask.setObjectType("InstructionItem");
                    TaskList = CalledByTransactionTask.parentInfo();
                    TaskList.setRevisionedObjectRef(taskListName, "", true);

                    InputData.dataField("ClearLocation").setValue("False");
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.namedObjectField("Resource").setRef(resourceName);
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    //dr["StartResult"] = camMessage.Message;
                    //dr["StartBoolResult"] = camMessage.Result;
                    if(camMessage.Message.ToString() == "성공!")
                    {
                        dr["Result"] = "작업시작 성공!"; //camMessage.Message;

                    }
                    else
                    {
                        dr["Result"] = camMessage.Message;
                    }
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        public int ComboBoxTask(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument;
            csiObject InputData;
            csiRevisionedObject DataCollectionDef;
            csiSubentity ParametricData;
            csiSubentityList DataPointDetails;
            csiSubentity listItem;
            csiNamedSubentity DataPoint;
            csiParentInfo DataPointParentInfo;
            csiNamedSubentity NDOValue;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo;
            csiRevisionedObject TaskList;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    // Set Data Collection Def
                    DataCollectionDef = InputData.revisionedObjectField("DataCollectionDef");
                    DataCollectionDef.setObjectType("UserDataCollectionDef");
                    DataCollectionDef.setRef((dr["UDC"] ?? "").ToString(), "", true);

                    // Set Prametric Data
                    ParametricData = InputData.subentityField("ParametricData");
                    ParametricData.createObject("DataPointSummary");
                    DataPointDetails = ParametricData.subentityList("DataPointDetails");

                    // Set listItem Loop (반복해야하는 부분)
                    listItem = DataPointDetails.appendItem();
                    listItem.setObjectType("DataPointDetails");
                    DataPoint = listItem.namedSubentityField("DataPoint");
                    DataPoint.setName((dr["Task Name"] ?? "").ToString());
                    DataPointParentInfo = DataPoint.parentInfo();
                    DataPointParentInfo.setObjectType("UserDataCollectionDef");
                    DataPointParentInfo.setRevisionedObjectRef((dr["UDC"] ?? "").ToString(), "", true);
                    listItem.dataField("DataType").setValue("5"); // 선택형 Task의 경우 Data Type 5로 고정
                    NDOValue = listItem.namedSubentityField("NDOValue");
                    NDOValue.setName((dr["구분"] ?? "").ToString());
                    NDOValue.setObjectType("NamedObjectGroup");

                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName("Data1"); // 실행하는 Task의 이름
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    DataPointParentInfo.setRevisionedObjectRef((dr["UDC"] ?? "").ToString(), "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef((dr["UDC"] ?? "").ToString(), "", true);


                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int KeyInTask(DataTable table)
        {
            int successCnt = 0;

            csiDocument ResponseDocument;
            csiObject InputData;
            csiRevisionedObject DataCollectionDef;
            csiSubentity ParametricData;
            csiSubentityList DataPointDetails;
            csiSubentity listItem;
            csiNamedSubentity DataPoint;
            csiParentInfo DataPointParentInfo;
            //csiNamedSubentity NDOValue;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo;
            csiRevisionedObject TaskList;

            try
            {
                CreateSession();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    // Set Data Collection Def
                    DataCollectionDef = InputData.revisionedObjectField("DataCollectionDef");
                    DataCollectionDef.setObjectType("UserDataCollectionDef");
                    DataCollectionDef.setRef((dr["UDC"] ?? "").ToString(), "", true);

                    // Set Prametric Data
                    ParametricData = InputData.subentityField("ParametricData");
                    ParametricData.createObject("DataPointSummary");
                    DataPointDetails = ParametricData.subentityList("DataPointDetails");

                    // Set listItem Loop (반복해야하는 부분)
                    listItem = DataPointDetails.appendItem();
                    listItem.setObjectType("DataPointDetails");
                    DataPoint = listItem.namedSubentityField("DataPoint");
                    DataPoint.setName("Setting 두께");
                    DataPointParentInfo = DataPoint.parentInfo();
                    DataPointParentInfo.setObjectType("UserDataCollectionDef");
                    DataPointParentInfo.setRevisionedObjectRef((dr["UDC"] ?? "").ToString(), "", true);
                    listItem.dataField("DataType").setValue("1"); // 1 : int, 4 : string, 9 : decimal
                    listItem.dataField("DataValue").setValue((dr["Setting 두께"] ?? "").ToString()); // 선택형 Task의 경우 Data Type 5로 고정

                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName("Setting 두께"); // 실행하는 Task의 이름
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    DataPointParentInfo.setRevisionedObjectRef((dr["UDC"] ?? "").ToString(), "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef((dr["UDC"] ?? "").ToString(), "", true);


                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Result)
                    {
                        csiService csiservice = ResponseDocument.getService();
                        if (csiservice != null && (dr["Container"] ?? "").ToString() == "Auto")
                        {
                            csiDataField csidatafield = (csiDataField)csiservice.responseData().getResponseFieldByName("CompletionMsg");
                            dr["Container"] = csidatafield.getValue().Split(new char[] { ' ' })[0].Trim();
                        }
                        successCnt++;
                    }
                    else
                    {
                        dr["Container"] = "";
                    }
                    dr["Result"] = camMessage.Message;
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            DestroySession();

            return successCnt;
        }

        public int ExecuteTaskByUDC(string taskListName, string taskItemName, string dataCollectionName, DataTable table, DataView collectionView, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument;
            csiObject InputData;
            csiRevisionedObject DataCollectionDef;
            csiSubentity ParametricData;
            csiSubentityList DataPointDetails;
            csiSubentity listItem;
            csiNamedSubentity DataPoint;
            csiParentInfo DataPointParentInfo;
            csiNamedSubentity NDOValue;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo;
            csiRevisionedObject TaskList;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }      
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    // Set Data Collection Def
                    DataCollectionDef = InputData.revisionedObjectField("DataCollectionDef");
                    DataCollectionDef.setObjectType("UserDataCollectionDef");
                    DataCollectionDef.setRef(dataCollectionName, "", true);

                    // Set Prametric Data
                    ParametricData = InputData.subentityField("ParametricData");
                    ParametricData.createObject("DataPointSummary");
                    DataPointDetails = ParametricData.subentityList("DataPointDetails");

                    foreach (DataRowView drv in collectionView)
                    {
                        if (dr[drv["DataPointName"].ToString()].ToString() == "") continue;

                        listItem = DataPointDetails.appendItem();
                        listItem.setObjectType("DataPointDetails");
                        DataPoint = listItem.namedSubentityField("DataPoint");
                        DataPoint.setName(drv["DataPointName"].ToString());
                        DataPointParentInfo = DataPoint.parentInfo();
                        DataPointParentInfo.setObjectType("UserDataCollectionDef");
                        DataPointParentInfo.setRevisionedObjectRef(dataCollectionName, "", true);
                        listItem.dataField("DataType").setValue(drv["DataType"].ToString()); 

                        if (drv["DataType"].ToString() == "5" && drv["NamedObjectGroupName"].ToString() != "")
                        {
                            NDOValue = listItem.namedSubentityField("NDOValue");
                            NDOValue.setName(dr[drv["DataPointName"].ToString()].ToString());
                            NDOValue.setObjectType("NamedObjectGroup");
                        }
                        else
                        {
                            listItem.dataField("DataValue").setValue(dr[drv["DataPointName"].ToString()].ToString());
                        }
                    }


                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName(taskItemName); // 실행하는 Task의 이름
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    TaskParentInfo.setRevisionedObjectRef(taskListName, "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef(taskListName, "", true);

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    dr["TaskResult"] = camMessage.Message;
                    dr["TaskBoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        // Task 실행 API - Result 수정
        public int ExecuteTaskByUDCResult(string taskListName, string taskItemName, string dataCollectionName, DataTable table, DataView collectionView, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument;
            csiObject InputData;
            csiRevisionedObject DataCollectionDef;
            csiSubentity ParametricData;
            csiSubentityList DataPointDetails;
            csiSubentity listItem;
            csiNamedSubentity DataPoint;
            csiParentInfo DataPointParentInfo;
            csiNamedSubentity NDOValue;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo;
            csiRevisionedObject TaskList;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    // Set Data Collection Def
                    DataCollectionDef = InputData.revisionedObjectField("DataCollectionDef");
                    DataCollectionDef.setObjectType("UserDataCollectionDef");
                    DataCollectionDef.setRef(dataCollectionName, "", true);

                    // Set Prametric Data
                    ParametricData = InputData.subentityField("ParametricData");
                    ParametricData.createObject("DataPointSummary");
                    DataPointDetails = ParametricData.subentityList("DataPointDetails");

                    foreach (DataRowView drv in collectionView)
                    {
                        if (dr[drv["DataPointName"].ToString()].ToString() == "") continue;

                        listItem = DataPointDetails.appendItem();
                        listItem.setObjectType("DataPointDetails");
                        DataPoint = listItem.namedSubentityField("DataPoint");
                        DataPoint.setName(drv["DataPointName"].ToString());
                        DataPointParentInfo = DataPoint.parentInfo();
                        DataPointParentInfo.setObjectType("UserDataCollectionDef");
                        DataPointParentInfo.setRevisionedObjectRef(dataCollectionName, "", true);
                        listItem.dataField("DataType").setValue(drv["DataType"].ToString());

                        if (drv["DataType"].ToString() == "5" && drv["NamedObjectGroupName"].ToString() != "")
                        {
                            NDOValue = listItem.namedSubentityField("NDOValue");
                            NDOValue.setName(dr[drv["DataPointName"].ToString()].ToString());
                            NDOValue.setObjectType("NamedObjectGroup");
                        }
                        else
                        {
                            listItem.dataField("DataValue").setValue(dr[drv["DataPointName"].ToString()].ToString());
                        }
                    }

                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName(taskItemName); // 실행하는 Task의 이름
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    TaskParentInfo.setRevisionedObjectRef(taskListName, "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef(taskListName, "", true);

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    if (camMessage.Message.ToString() == "성공!")
                    {
                        dr["Result"] = "제출 성공!"; //camMessage.Message;

                    }
                    else
                    {
                        dr["Result"] = camMessage.Message;
                    }
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        public int WorkFinishByUDC(string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo = null;
            csiRevisionedObject TaskList;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }      
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    if (dr["TaskBoolResult"].ToString() == "False" || dr["TaskBoolResult"].ToString() == "") continue;

                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());

                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName("작업종료");
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    TaskParentInfo.setRevisionedObjectRef(taskListName, "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef(taskListName, "", true);

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    dr["EndResult"] = camMessage.Message;
                    dr["EndBoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }
 
            return successCnt;
        }

        // 작업종료 API - Result 수정
        public int WorkFinishByUDCResult(string cif ,string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo = null;
            csiRevisionedObject TaskList;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    //if (dr["BoolResult"].ToString() == "False" || dr["BoolResult"].ToString() == "") continue;

                    // Set Service Type
                    CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                    // Set Input Data
                    InputData = gService.inputData();
                    if(cif == "설비별 배치 작업" || cif == "공정별 LOT 작업(CMOS)")
                    {
                        InputData.dataField("Comments").setValue(dr["Comment"].ToString());
                    }
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    

                    // Set Task
                    Task = InputData.namedSubentityField("Task");
                    Task.setName("작업종료");
                    TaskParentInfo = Task.parentInfo();
                    TaskParentInfo.setObjectType("TaskList");
                    TaskParentInfo.setRevisionedObjectRef(taskListName, "", true);

                    // Set Task List
                    TaskList = InputData.revisionedObjectField("TaskList");
                    TaskList.setObjectType("TaskList");
                    TaskList.setRef(taskListName, "", true);

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    //dr["Result"] = camMessage.Message;
                    if (camMessage.Message.ToString() == "성공!")
                    {
                        dr["Result"] = "작업종료 성공!"; //camMessage.Message;

                    }
                    else
                    {
                        dr["Result"] = camMessage.Message;
                    }
                    dr["BoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        public string[] ExecuteTaskPassFail(string taskListName, string Container, string True_False, bool isSessionCreate)
        {
            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity Task;
            csiParentInfo TaskParentInfo = null;
            csiRevisionedObject TaskList;
            string[] ArrMessage = { "", "" };

            try
            {
                //Session이 없는 경우 Session 생성
                if (isSessionCreate == false)
                {
                    CreateSession();
                    isSessionCreate = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // 결과 값을 배열에 넣어서 Return
                ArrMessage[0] = "세션 오류";
                return ArrMessage;
            }

            try
            {
                // Set Service Type
                CreateDocumentandService("ExecuteTaskDoc", "ExecuteTask");

                // Set Input Data
                InputData = gService.inputData();
                InputData.namedObjectField("Container").setRef( Container );

                // Set Pass
                InputData.dataField("Pass").setValue( True_False );

                // Set Task
                Task = InputData.namedSubentityField("Task");
                Task.setName("Pass/Fail");
                TaskParentInfo = Task.parentInfo();
                TaskParentInfo.setObjectType("TaskList");
                TaskParentInfo.setRevisionedObjectRef(taskListName, "", true);

                // Set Task List
                TaskList = InputData.revisionedObjectField("TaskList");
                TaskList.setObjectType("TaskList");
                TaskList.setRef(taskListName, "", true);

                // Service Excute and request Completion Msg
                gService.setExecute();
                gService.requestData().requestField("CompletionMsg");
                gService.requestData().requestField("ACEMessage");
                gService.requestData().requestField("ACEStatus");

                // Print XMl Dcoument
                PrintDoc(gDocument.asXML(), true);
                ResponseDocument = gDocument.submit();
                PrintDoc(ResponseDocument.asXML(), false);

                ErrorsCheck(ResponseDocument);

                // 결과 값을 배열에 넣어서 Return
                ArrMessage[0] = camMessage.Result.ToString();
                ArrMessage[1] = camMessage.Message;

                //dr.BeginEdit();
                //dr["EndResult"] = camMessage.Message;
                //dr["EndBoolResult"] = camMessage.Result;
                //dr.EndEdit();
                //}
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;

                // 결과 값을 배열에 넣어서 Return
                ArrMessage[0] = "false";
                ArrMessage[1] = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return ArrMessage;
        }

        public int NextWorkByUDC(string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity CalledByTransactionTask;
            csiParentInfo TaskList = null;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    if (dr["TaskBoolResult"].ToString() == "False" || dr["TaskBoolResult"].ToString() == "") continue;
                    if (dr["EndBoolResult"].ToString() == "False" || dr["EndBoolResult"].ToString() == "") continue;

                    // Set Service Type
                    CreateDocumentandService("MoveStdDoc", "MoveStd");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Called By Transaction Task
                    CalledByTransactionTask = InputData.namedSubentityField("CalledByTransactionTask");
                    CalledByTransactionTask.setName("다음공정으로");
                    CalledByTransactionTask.setObjectType("InstructionItem");
                    TaskList = CalledByTransactionTask.parentInfo();
                    TaskList.setRevisionedObjectRef(taskListName, "", true);

                    //InputData.dataField("ClearLocation").setValue("false");
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.namedObjectField("Resource").setRef(resourceName);
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    dr["MoveStdResult"] = camMessage.Message;
                    dr["MoveStdBoolResult"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        // 다음공정이동 API - Result 수정
        public int NextWorkByUDCResult(string taskListName, string resourceName, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity CalledByTransactionTask;
            csiParentInfo TaskList = null;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    //if (dr["BoolResult"].ToString() == "False" || dr["BoolResult"].ToString() == "") continue;
                    //if (dr["BoolResult2"].ToString() == "False" || dr["BoolResult2"].ToString() == "") continue;
                    if (dr["BoolResult"].ToString() == "False")
                    {
                        dr["Result2"] = "실패";
                        dr["BoolResult2"] = false;
                        continue;
                    }

                    // Set Service Type
                    CreateDocumentandService("MoveStdDoc", "MoveStd");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Called By Transaction Task
                    CalledByTransactionTask = InputData.namedSubentityField("CalledByTransactionTask");
                    CalledByTransactionTask.setName("다음공정으로");
                    CalledByTransactionTask.setObjectType("InstructionItem");
                    TaskList = CalledByTransactionTask.parentInfo();
                    TaskList.setRevisionedObjectRef(taskListName, "", true);

                    //InputData.dataField("ClearLocation").setValue("false");
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.namedObjectField("Resource").setRef(resourceName);
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    //dr["Result2"] = camMessage.Message;
                    if (camMessage.Message.ToString() == "성공!")
                    {
                        dr["Result2"] = "다음 공정이동 성공!"; //camMessage.Message;

                    }
                    else
                    {
                        dr["Result2"] = camMessage.Message;
                    }
                    dr["BoolResult2"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        public int NextWorkByUDCHeatResult(string taskListName, string resourceName, DataTable table, bool NextHeat, bool isSessionCreate = true)
        {
            int successCnt = 0;

            csiDocument ResponseDocument = null;
            csiObject InputData = null;
            csiNamedSubentity CalledByTransactionTask;
            csiParentInfo TaskList = null;

            try
            {
                if (isSessionCreate)
                {
                    CreateSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "세션 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    //if (dr["BoolResult"].ToString() == "False" || dr["BoolResult"].ToString() == "") continue;
                    //if (dr["BoolResult2"].ToString() == "False" || dr["BoolResult2"].ToString() == "") continue;
                    if (dr["BoolResult"].ToString() == "False")
                    {
                        dr["Result2"] = "실패";
                        dr["BoolResult2"] = false;
                        continue;
                    }

                    // Set Service Type
                    CreateDocumentandService("MoveStdDoc", "MoveStd");

                    // Set Input Data
                    InputData = gService.inputData();

                    //Set Called By Transaction Task
                    CalledByTransactionTask = InputData.namedSubentityField("CalledByTransactionTask");
                    CalledByTransactionTask.setName("다음공정으로");
                    CalledByTransactionTask.setObjectType("InstructionItem");
                    TaskList = CalledByTransactionTask.parentInfo();
                    TaskList.setRevisionedObjectRef(taskListName, "", true);

                    //InputData.dataField("ClearLocation").setValue("false");
                    InputData.namedObjectField("Container").setRef((dr["Container"] ?? "").ToString());
                    InputData.namedObjectField("Resource").setRef(resourceName);
                    InputData.namedObjectField("TaskContainer").setRef((dr["Container"] ?? "").ToString());
                    if (NextHeat == true)
                    {
                        InputData.namedObjectField("Path").setRef("CSI_TFT_HEAT");
                    }

                    // Service Excute and request Completion Msg
                    gService.setExecute();
                    gService.requestData().requestField("CompletionMsg");
                    gService.requestData().requestField("ACEMessage");
                    gService.requestData().requestField("ACEStatus");

                    // Print XMl Dcoument
                    PrintDoc(gDocument.asXML(), true);
                    ResponseDocument = gDocument.submit();
                    PrintDoc(ResponseDocument.asXML(), false);

                    ErrorsCheck(ResponseDocument);

                    dr.BeginEdit();
                    //dr["Result2"] = camMessage.Message;
                    if (camMessage.Message.ToString() == "성공!")
                    {
                        dr["Result2"] = "다음 공정이동 성공!"; //camMessage.Message;

                    }
                    else
                    {
                        dr["Result2"] = camMessage.Message;
                    }
                    dr["BoolResult2"] = camMessage.Result;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                camMessage.Result = false;
                camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }
        #endregion

        #region Data Extract
        //TFT PQC Data 불러오기
        public int PQCDataOpen(string DataCollectionLookUpEdit, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            try
            {
                //arrTask[0] : 작시 작종
                //arrTask[1] : UDCD
                //arrTask[2] : UDCD의 taskName
                //dataTable  : 화면에 뿌려진 엑셀
                //successCnt = WrGlobal.Camster_Common.PQCDataOpen(arrTask[0], resourceName, dataTable, false);
                //DataCollectionLookUpEdit.EditValue : 작업 종류의 id 값 : 001c61800000010e

                //여기에서부터 다시 작업 해야 함
                //FT(PQC) DATA 38 건 : exec [CAMDBsh].[RY_VR_Proc_TFT_FT_DATA] '260-T221F127NK01' //PQC
                //QC(OQC) Data 48 건 : exec [CAMDBsh].[RY_VR_Proc_TFT_QC_DATA] '001c61800000010e' //OQC
                //작업종류는 같아야 한다, 사유는 SAP 코드가 같아야 함



                foreach (DataRow dr in table.Rows)
                {
                    //PQC(FT) Data 가져오기 ex) 컨테이너 260-T221F127NK01 38건을 가져옴
                    DataView DVPQC = work.GetPQCDataLoad(dr["Container"].ToString().Replace("270-", "260-"));
                    int j = 0;
                    //Array[] CfOQC = null;
                    //string[] CfPqcCheckID = new string[1000];
                    //string[] CfPqcUDCD = new string[1000];
                    string[] CfPqcValue = new string[1000];
                    //string[] CfPqcRev = new string[1000];
                    string[] CfPqcData = new string[1000];

                    foreach (DataRow drPQC in DVPQC.Table.Rows)
                    {//QC Data 48건
                        //CfPqcCheckID[j] = drPQC["ChK_ID"].ToString();
                        //CfPqcUDCD[j] = drPQC["DataPointName"].ToString();
                        CfPqcValue[j] = drPQC["VALUE_RESULT_NAME"].ToString();
                        //CfPqcRev[j] = drPQC["DataCollectionDefRevision"].ToString();
                        CfPqcData[j] = drPQC["SPEC_NAME"].ToString();
                        j++;
                    }

                    //OQC Data ex) 컨테이너 270-T221F127NK01 48건을 가져옴
                    DataView DVOQC = work.GetQCDataLoad(DataCollectionLookUpEdit);
                    int i = 0;
                    //Array[] CfOQC = null;
                    //string[] CfOqcCheckID = new string[1000];
                    string[] CfOqcUDCD = new string[1000];
                    //string[] CfOqcValue = new string[1000];
                    //string[] CfOqcRev = new string[1000];
                    string[] CfOqcData = new string[1000];

                    foreach (DataRow drOQC in DVOQC.Table.Rows)
                    {//QC Data 48건
                        //CfOqcCheckID[i] = drOQC["CHK_ID"].ToString();
                        CfOqcUDCD[i] = drOQC["DataPointName"].ToString();
                        //CfOqcRev[i] = drOQC["DataCollectionDefRevision"].ToString();
                        CfOqcData[i] = drOQC["SPEC_NAME"].ToString();

                        //배열내 String 찾기
                        //bool index = CfPqcCheckID.Contains(drOQC["CHK_ID"].ToString());

                        //OQC CHK_ID가 PQC CHK_ID 값과 같으면, PQC 배열에 OQC CHK_ID 값을 포함하고 있으면 배열 요소 값을 찾아라
                        if (CfPqcData.Contains(CfOqcData[i].ToString()) ) // 요조건에 하나도 안걸린다.
                        {
                            //if (i > 46) 
                            //{
                            //    string a = "0";
                            //    int b = 0;
                            //}
                            int IndexPQC = Array.IndexOf(CfPqcData, CfOqcData[i].ToString());//PQC 요소값

                            //int IndexOQC = Array.IndexOf(CfOqcData, CfOqcData[i].ToString());//OQC 요소값
                            //dr[CfOqcUDCD[IndexOQC].ToString()] = CfPqcValue[IndexPQC].ToString();//
                        
                            //요소값을 찾았으면 해당 값을 가져와서 Table에 값을 넣자
                            dr.BeginEdit();
                            dr[CfOqcUDCD[i].ToString()] = CfPqcValue[IndexPQC].ToString();
                            //dr[""];
                            dr.EndEdit();
                            successCnt++;
                        }
                        i++;
                    }

                    dr.BeginEdit();
                    dr["Result"] = "불러오기 성공!"; //camMessage.Message;
                    dr.EndEdit();


                    //Array.IndexOf(drOQC["CHK_ID"], drPQC["CHK_ID"]);





                    //dr.BeginEdit();
                    ////dr["StartResult"] = camMessage.Message;
                    ////dr["StartBoolResult"] = camMessage.Result;
                    //if (camMessage.Message.ToString() == "성공!")
                    //{
                    //    dr["Result"] = "불러오기 성공!"; //camMessage.Message;

                    //}
                    //else
                    //{
                    //    dr["Result"] = camMessage.Message;
                    //}
                    //dr["BoolResult"] = camMessage.Result;
                    //dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                //camMessage.Result = false;
                //camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }

        //CMOS PQC Data 불러오기
        public int CMOSPQCDataOpen(string DataCollectionLookUpEdit, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            try
            {
                //arrTask[0] : 작시 작종
                //arrTask[1] : UDCD
                //arrTask[2] : UDCD의 taskName
                //dataTable  : 화면에 뿌려진 엑셀
                //successCnt = WrGlobal.Camster_Common.PQCDataOpen(arrTask[0], resourceName, dataTable, false);
                //DataCollectionLookUpEdit.EditValue : 작업 종류의 id 값 : 001c61800000010e

                //여기에서부터 다시 작업 해야 함
                //FT(PQC) DATA 38 건 : exec [CAMDBsh].[RY_VR_Proc_TFT_FT_DATA] '260-T221F127NK01' //PQC
                //QC(OQC) Data 48 건 : exec [CAMDBsh].[RY_VR_Proc_TFT_QC_DATA] '001c61800000010e' //OQC
                //작업종류는 같아야 한다, 사유는 SAP 코드가 같아야 함



                foreach (DataRow dr in table.Rows)
                {
                    //PQC(FT) Data 가져오기 ex) 컨테이너 260-T221F127NK01 38건을 가져옴
                    DataTable DTPQC = work.GetCmosPQCDataLoad(dr["Container"].ToString().Replace("QC-", "AY-"));
                  
                    int j = 0;
                    //Array[] CfOQC = null;
                    //string[] CfPqcCheckID = new string[1000];
                    //string[] CfPqcUDCD = new string[1000];
                    string[] CfPqcValue = new string[1000];
                    //string[] CfPqcRev = new string[1000];
                    string[] CfPqcData = new string[1000];

                    foreach (DataRow drPQC in DTPQC.Rows)
                    {//QC Data 48건
                        //CfPqcCheckID[j] = drPQC["ChK_ID"].ToString();
                        //CfPqcUDCD[j] = drPQC["DataPointName"].ToString();
                        CfPqcValue[j] = drPQC["VALUE_RESULT_NAME"].ToString();
                        //CfPqcRev[j] = drPQC["DataCollectionDefRevision"].ToString();
                        CfPqcData[j] = drPQC["SPEC_NAME"].ToString();
                        j++;
                    }

                    //OQC Data ex) 컨테이너 270-T221F127NK01 48건을 가져옴
                    DataView DVOQC = work.GetQCDataLoad(DataCollectionLookUpEdit);
                    int i = 0;
                    //Array[] CfOQC = null;
                    //string[] CfOqcCheckID = new string[1000];
                    string[] CfOqcUDCD = new string[1000];
                    //string[] CfOqcValue = new string[1000];
                    //string[] CfOqcRev = new string[1000];
                    string[] CfOqcData = new string[1000];

                    foreach (DataRow drOQC in DVOQC.Table.Rows)
                    {//QC Data 48건
                        //CfOqcCheckID[i] = drOQC["CHK_ID"].ToString();
                        CfOqcUDCD[i] = drOQC["DataPointName"].ToString();
                        //CfOqcRev[i] = drOQC["DataCollectionDefRevision"].ToString();
                        CfOqcData[i] = drOQC["SPEC_NAME"].ToString();

                        //배열내 String 찾기
                        //bool index = CfPqcCheckID.Contains(drOQC["CHK_ID"].ToString());

                        //OQC CHK_ID가 PQC CHK_ID 값과 같으면, PQC 배열에 OQC CHK_ID 값을 포함하고 있으면 배열 요소 값을 찾아라
                        if (CfPqcData.Contains(CfOqcData[i].ToString())) // 요조건에 하나도 안걸린다.
                        {
                            //if (i > 46)
                            //{
                            //    string a = "0";
                            //    int b = 0;
                            //}
                            int IndexPQC = Array.IndexOf(CfPqcData, CfOqcData[i].ToString());//PQC 요소값

                            //int IndexOQC = Array.IndexOf(CfOqcData, CfOqcData[i].ToString());//OQC 요소값
                            //dr[CfOqcUDCD[IndexOQC].ToString()] = CfPqcValue[IndexPQC].ToString();//

                            //요소값을 찾았으면 해당 값을 가져와서 Table에 값을 넣자
                            dr.BeginEdit();
                            dr[CfOqcUDCD[i].ToString()] = CfPqcValue[IndexPQC].ToString();
                            //dr[""];
                            dr.EndEdit();
                        }
                        i++;
                    }

                    dr.BeginEdit();
                    dr["Result"] = "불러오기 성공!"; //camMessage.Message;
                    dr.EndEdit();


                    //Array.IndexOf(drOQC["CHK_ID"], drPQC["CHK_ID"]);





                    //dr.BeginEdit();
                    ////dr["StartResult"] = camMessage.Message;
                    ////dr["StartBoolResult"] = camMessage.Result;
                    //if (camMessage.Message.ToString() == "성공!")
                    //{
                    //    dr["Result"] = "불러오기 성공!"; //camMessage.Message;

                    //}
                    //else
                    //{
                    //    dr["Result"] = camMessage.Message;
                    //}
                    //dr["BoolResult"] = camMessage.Result;
                    //dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                //camMessage.Result = false;
                //camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }


        //CMOS PQC Data 불러오기 (23.06.19 Dictionary 사용하여 불러오기로 변경)
        public int CMOSPQCDataOpenNew(string DataCollectionLookUpEdit, DataTable table, bool isSessionCreate = true)
        {
            int successCnt = 0;

            try
            {
                foreach (DataRow dr in table.Rows)
                {
                    // FT작업 데이터(OOTB + Upload)
                    DataTable DTPQC = work.GetCmosPQCDataLoad(dr["Container"].ToString().Replace("QC-", "AY-"));
                    // FT작업 데이터(OOTB + Upload)의 SPEC_NAME 별 VALUE_RESULT_NAME 을 Dictionary에 저장
                    Dictionary<string, string> pqcValuePairs = new Dictionary<string, string>();

                    foreach (DataRow drPQC in DTPQC.Rows)
                    {
                        if (!pqcValuePairs.ContainsKey(drPQC["SPEC_NAME"].ToString()))
                        {
                            pqcValuePairs.Add(drPQC["SPEC_NAME"].ToString(), drPQC["VALUE_RESULT_NAME"].ToString());
                        }
                    }

                    // OQC 검사 항목 
                    DataView DVOQC = work.GetQCDataLoad(DataCollectionLookUpEdit);
                    // OQC 검사 항목의 SPEC_NAME 별 DataPointName 을 Dictionary에 저장
                    Dictionary<string, string> oqcValuePairs = new Dictionary<string, string>();

                    foreach (DataRow drOQC in DVOQC.Table.Rows)
                    {
                        // oqcValuePairs 에 SPEC_NAME, DataPointName 추가
                        if (!oqcValuePairs.ContainsKey(drOQC["SPEC_NAME"].ToString()))
                        {
                            oqcValuePairs.Add(drOQC["SPEC_NAME"].ToString(), drOQC["DataPointName"].ToString());
                        }

                        // pqcValuePairs 에 OQC SPEC_NAME이 있으면
                        if (pqcValuePairs.ContainsKey(drOQC["SPEC_NAME"].ToString()))
                        {
                            dr.BeginEdit();
                            // OQC DataPoint 컬럼에 PQC VALUE_RESULT_NAME 입력
                            dr[oqcValuePairs[drOQC["SPEC_NAME"].ToString()]] = pqcValuePairs[drOQC["SPEC_NAME"].ToString()];
                            dr.EndEdit();
                        }
                    }

                    dr.BeginEdit();
                    dr["Result"] = "불러오기 성공!"; //camMessage.Message;
                    dr.EndEdit();
                }
            }
            catch (Exception ex)
            {
                //camMessage.Result = false;
                //camMessage.Message = ex.Message;
            }

            if (isSessionCreate)
            {
                DestroySession();
            }

            return successCnt;
        }
        #endregion

    }
}
