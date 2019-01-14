using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using SapB1MutHelper.Models;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using BoMessageTime = SAPbouiCOM.BoMessageTime;
using BoStatusBarMessageType = SAPbouiCOM.BoStatusBarMessageType;

namespace SapB1MutHelper

{
    public class UdoCreation
    {
        public static bool RegisterUdo(string udoCode, string udoName, BoUDOObjType udoType,
            Dictionary<string, string> fields, string udohTableName = "", string udodTableName = "",
            BoYesNoEnum logOption = BoYesNoEnum.tNO)
        {
            var functionReturnValue = false;
            try
            {
                var vUdoMd = (UserObjectsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                vUdoMd.Code = udoCode;
                vUdoMd.Name = udoName;
                vUdoMd.ObjectType = udoType;
                vUdoMd.TableName = udohTableName;
                vUdoMd.CanDelete = BoYesNoEnum.tYES;
                vUdoMd.CanFind = BoYesNoEnum.tYES;
                vUdoMd.CanCancel = BoYesNoEnum.tYES;
                vUdoMd.CanClose = BoYesNoEnum.tYES;
                vUdoMd.CanCreateDefaultForm = BoYesNoEnum.tNO;
                vUdoMd.EnableEnhancedForm = BoYesNoEnum.tNO;


                if (logOption == BoYesNoEnum.tYES)
                {
                    vUdoMd.CanLog = BoYesNoEnum.tYES;
                    vUdoMd.LogTableName = "A" + udohTableName;
                }

                #region Bul Alanlarının Eklenmesi

                foreach (var item in fields)
                {
                    vUdoMd.FindColumns.ColumnAlias = item.Key;
                    vUdoMd.FindColumns.ColumnDescription = item.Value;
                    vUdoMd.FindColumns.Add();
                }

                #endregion

                #region Görünecek Alanlarının Eklenmesi

                var count = 0;
                foreach (var item in fields)
                {
                    count++;
                    vUdoMd.FormColumns.FormColumnAlias = item.Key;
                    vUdoMd.FormColumns.FormColumnDescription = item.Value;
                    if (count > 1) vUdoMd.FormColumns.Editable = BoYesNoEnum.tYES;

                    vUdoMd.FormColumns.Add();
                }

                #endregion

                if (vUdoMd.Add() == 0)
                {
                    functionReturnValue = true;
                    if (Helper.OCompany.InTransaction)
                        Helper.OCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                    Application.SBO_Application.StatusBar.SetText(
                        "Successfully Registered UDO >" + udoCode + ">" + udoName + " >" +
                        Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Failed to Register UDO >" + udoCode + ">" + udoName + " >" +
                        Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short);
                }

                Marshal.ReleaseComObject(vUdoMd);
                GC.Collect();
                if (true & Helper.OCompany.InTransaction)
                    Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
            }
            catch (Exception)
            {
                if (Helper.OCompany.InTransaction)
                    Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
            }

            return functionReturnValue;
        }

        public static bool MeRegisterUdo(string udoCode, string udoName, BoUDOObjType udoType, string udohTableName,
            string udodTableName = "")
        {
            var functionReturnValue = false;
            var actionSuccess = false;
            try
            {
                var vUdoMd = (UserObjectsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                vUdoMd.Code = udoCode;
                vUdoMd.Name = udoName;
                vUdoMd.ObjectType = udoType;
                vUdoMd.TableName = udohTableName;
                vUdoMd.CanDelete = BoYesNoEnum.tYES;
                vUdoMd.CanFind = BoYesNoEnum.tYES;
                vUdoMd.CanCancel = BoYesNoEnum.tYES;
                vUdoMd.CanClose = BoYesNoEnum.tYES;
                vUdoMd.CanCreateDefaultForm = BoYesNoEnum.tNO;
                vUdoMd.EnableEnhancedForm = BoYesNoEnum.tNO;
                vUdoMd.MenuItem = BoYesNoEnum.tYES;
                vUdoMd.Code = udoCode;
                vUdoMd.Name = udoName;
                vUdoMd.TableName = udohTableName;
                vUdoMd.MenuCaption = udoName;

                if (udoName == "WorkSheet")
                {
                    vUdoMd.CanCreateDefaultForm = BoYesNoEnum.tYES;
                    vUdoMd.EnableEnhancedForm = BoYesNoEnum.tNO;

                    vUdoMd.MenuCaption = "WorkSheet";
                    vUdoMd.FatherMenuID = 11520;
                    vUdoMd.Position = -1;
                }


                vUdoMd.ObjectType = udoType;


                if (vUdoMd.Add() == 0)
                {
                    functionReturnValue = true;
                    if (Helper.OCompany.InTransaction)
                        Helper.OCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                    Application.SBO_Application.StatusBar.SetText(
                        "Successfully Registered UDO >" + udoCode + ">" + udoName + " >" +
                        Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Failed to Register UDO >" + udoCode + ">" + udoName + " >" +
                        Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short);
                }

                Marshal.ReleaseComObject(vUdoMd);
                GC.Collect();
                if ((actionSuccess == false) & Helper.OCompany.InTransaction)
                    Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
            }
            catch (Exception)
            {
                if (Helper.OCompany.InTransaction)
                    Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
            }

            return functionReturnValue;
        }

        public static bool UdoExists(string code)
        {
            GC.Collect();
            var vUdoMd = (UserObjectsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            var vReturnCode = vUdoMd.GetByKey(code);
            Marshal.ReleaseComObject(vUdoMd);
            return vReturnCode;
        }

        public static bool CreateFunction(string functionName, string function)
        {
            var functionReturnValue = false;
            long vRetVal = 0;
            long vErrCode = 0;
            var vErrMsg = "";
            try
            {
                if (!FunctionExists(functionName))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Creating Procedure " + functionName + " ...................", BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);

                    if (vRetVal != 0)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to Create Procedure " + functionName + vErrCode + " " + vErrMsg,
                            BoMessageTime.bmt_Short);

                        return false;
                    }

                    var oRsObjectExists = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRsObjectExists.DoQuery(function);
                    Application.SBO_Application.StatusBar.SetText("[" + functionName + "] - Created Successfully!!!",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                    return true;
                }

                GC.Collect();
                return false;
            }
            catch (Exception)
            {
                // ignored
            }

            return functionReturnValue;
        }

        public static bool CreateTrigger(string triggerName, string trigger)
        {
            long vRetVal = 0;
            long vErrCode = 0;
            var vErrMsg = "";
            try
            {
                if (!TriggerExists(triggerName))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Creating Trigger " + triggerName + " ...................", BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);

                    if (vRetVal != 0)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to Create Procedure " + triggerName + vErrCode + " " + vErrMsg,
                            BoMessageTime.bmt_Short);

                        return false;
                    }

                    var oRsObjectExists = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRsObjectExists.DoQuery(trigger);
                    Application.SBO_Application.StatusBar.SetText("[" + triggerName + "] - Created Successfully!!!",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                    return true;
                }

                GC.Collect();
                return false;
            }
            catch (Exception)
            {
                // ignored
            }

            return false;
        }

        public static bool FunctionExists(string functionName)
        {
            var oFlag = false;
            var oObjectExists =
                "SELECT  1 FROM    Information_schema.Routines WHERE   Specific_schema = 'dbo' AND specific_name = '" +
                functionName + "' ";
            oObjectExists += " AND Routine_Type = 'FUNCTION' ";
            var oRsObjectExists = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRsObjectExists.DoQuery(oObjectExists);
            if (oRsObjectExists.RecordCount == 1) oFlag = true;
            return oFlag;
        }

        public static bool TriggerExists(string triggerName)
        {
            var oFlag = false;
            var oObjectExists =
                "SELECT trigger_name = name, trigger_owner = USER_NAME(uid), table_name = OBJECT_NAME(parent_obj),";
            oObjectExists +=
                " isupdate = OBJECTPROPERTY( id, 'ExecIsUpdateTrigger'), isdelete = OBJECTPROPERTY( id, 'ExecIsDeleteTrigger'),";
            oObjectExists +=
                " isinsert = OBJECTPROPERTY( id, 'ExecIsInsertTrigger'), isafter = OBJECTPROPERTY( id, 'ExecIsAfterTrigger'),";
            oObjectExists +=
                " isinsteadof = OBJECTPROPERTY( id, 'ExecIsInsteadOfTrigger'),status = CASE OBJECTPROPERTY(id, 'ExecIsTriggerDisabled') WHEN 1 THEN 'Disabled' ELSE 'Enabled' END ";
            oObjectExists += " FROM sysobjects WHERE type = 'TR' and name='" + triggerName + "'";
            var oRsObjectExists = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRsObjectExists.DoQuery(oObjectExists);
            if (oRsObjectExists.RecordCount == 1) oFlag = true;
            //oFlag = Convert.ToBoolean(oRsObjectExists.Fields.Item(0).Value);
            return oFlag;
        }

        public static bool RegisterUdoWithChildTable(string udoCode, string udoName, BoUDOObjType udoType,
            Dictionary<string, string> fields, string udohTableName = "", string udodTableName = "",
            BoYesNoEnum logOption = BoYesNoEnum.tNO, List<ChildTable> chList = null)
        {
            {
                var functionReturnValue = false;
                var actionSuccess = false;
                try
                {
                    var vUdoMd = (UserObjectsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                    vUdoMd.Code = udoCode;
                    vUdoMd.Name = udoName;
                    vUdoMd.ObjectType = udoType;
                    vUdoMd.TableName = udohTableName;
                    vUdoMd.CanDelete = BoYesNoEnum.tYES;
                    vUdoMd.CanFind = BoYesNoEnum.tYES;
                    vUdoMd.CanCancel = BoYesNoEnum.tYES;
                    vUdoMd.CanClose = BoYesNoEnum.tYES;
                    vUdoMd.CanCreateDefaultForm = BoYesNoEnum.tNO;
                    vUdoMd.EnableEnhancedForm = BoYesNoEnum.tNO;
                    vUdoMd.CanLog = BoYesNoEnum.tNO;
                    if (logOption == BoYesNoEnum.tYES)
                    {
                        vUdoMd.CanLog = BoYesNoEnum.tYES;
                        vUdoMd.LogTableName = "A" + udohTableName;
                    }

                    #region Bul Alanlarının Eklenmesi

                    foreach (var item in fields)
                    {
                        vUdoMd.FindColumns.ColumnAlias = item.Key;
                        vUdoMd.FindColumns.ColumnDescription = item.Value;
                        vUdoMd.FindColumns.Add();
                    }

                    #endregion

                    #region chList

                    var childNumber = 0;
                    foreach (var ch in chList)
                    {
                        childNumber++;
                        vUdoMd.ChildTables.TableName = ch.TableName;
                        if (childNumber != chList.Count()) vUdoMd.ChildTables.Add();
                        if (childNumber == chList.Count())
                        {
                            var setCurrentLine = 0;
                            foreach (var item in ch.FormColumn)
                            {
                                vUdoMd.FormColumns.SetCurrentLine(setCurrentLine);
                                vUdoMd.FormColumns.SonNumber = childNumber;
                                vUdoMd.FormColumns.FormColumnAlias = item.FormColumnAlias;
                                vUdoMd.FormColumns.FormColumnDescription = item.FormColumnDescription;
                                vUdoMd.FormColumns.Editable = item.Editable;
                                vUdoMd.FormColumns.Add();

                                setCurrentLine++;
                            }
                        }

                        vUdoMd.EnhancedFormColumns.ColumnAlias = "DocEntry";
                        vUdoMd.EnhancedFormColumns.ColumnDescription = "DocEntry";
                        vUdoMd.EnhancedFormColumns.ColumnIsUsed = BoYesNoEnum.tNO;
                        vUdoMd.EnhancedFormColumns.ColumnNumber = 1;
                        vUdoMd.EnhancedFormColumns.ChildNumber = childNumber;
                        vUdoMd.EnhancedFormColumns.Add();

                        vUdoMd.EnhancedFormColumns.ColumnAlias = "LineId";
                        vUdoMd.EnhancedFormColumns.ColumnDescription = "LineId";
                        vUdoMd.EnhancedFormColumns.ColumnIsUsed = BoYesNoEnum.tNO;
                        vUdoMd.EnhancedFormColumns.ColumnNumber = 2;
                        vUdoMd.EnhancedFormColumns.ChildNumber = childNumber;
                        vUdoMd.EnhancedFormColumns.Add();

                        vUdoMd.EnhancedFormColumns.ColumnAlias = "Object";
                        vUdoMd.EnhancedFormColumns.ColumnDescription = "Object";
                        vUdoMd.EnhancedFormColumns.ColumnIsUsed = BoYesNoEnum.tNO;
                        vUdoMd.EnhancedFormColumns.ColumnNumber = 3;
                        vUdoMd.EnhancedFormColumns.ChildNumber = childNumber;
                        vUdoMd.EnhancedFormColumns.Add();

                        vUdoMd.EnhancedFormColumns.ColumnAlias = "LogInst";
                        vUdoMd.EnhancedFormColumns.ColumnDescription = "LogInst";
                        vUdoMd.EnhancedFormColumns.ColumnIsUsed = BoYesNoEnum.tNO;
                        vUdoMd.EnhancedFormColumns.ColumnNumber = 4;
                        vUdoMd.EnhancedFormColumns.ChildNumber = childNumber;
                        vUdoMd.EnhancedFormColumns.Add();


                        var columNumber = 5;
                        foreach (var item in ch.FormColumn.Where(k => k.FormColumnAlias != "DocEntry"))
                        {
                            vUdoMd.EnhancedFormColumns.ColumnAlias = item.FormColumnAlias;
                            vUdoMd.EnhancedFormColumns.ColumnDescription = item.FormColumnDescription;
                            vUdoMd.EnhancedFormColumns.ColumnIsUsed = item.Editable;
                            vUdoMd.EnhancedFormColumns.Editable = item.Editable;
                            vUdoMd.EnhancedFormColumns.ColumnNumber = columNumber;
                            vUdoMd.EnhancedFormColumns.ChildNumber = childNumber;
                            vUdoMd.EnhancedFormColumns.Add();

                            columNumber++;
                        }
                    }

                    #endregion

                    #region Görünecek Alanlarının Eklenmesi

                    var count = 0;
                    foreach (var item in fields)
                    {
                        count++;
                        vUdoMd.FormColumns.FormColumnAlias = item.Key;
                        vUdoMd.FormColumns.FormColumnDescription = item.Value;
                        if (count > 1) vUdoMd.FormColumns.Editable = BoYesNoEnum.tYES;

                        vUdoMd.FormColumns.Add();
                    }

                    #endregion


                    if (vUdoMd.Add() == 0)
                    {
                        functionReturnValue = true;
                        if (Helper.OCompany.InTransaction)
                            Helper.OCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                        Application.SBO_Application.StatusBar.SetText(
                            "Successfully Registered UDO >" + udoCode + ">" + udoName + " >" +
                            Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to Register UDO >" + udoCode + ">" + udoName + " >" +
                            Helper.OCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short);
                    }

                    Marshal.ReleaseComObject(vUdoMd);
                    vUdoMd = null;
                    GC.Collect();
                    if ((actionSuccess == false) & Helper.OCompany.InTransaction)
                        Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                catch (Exception)
                {
                    if (Helper.OCompany.InTransaction)
                        Helper.OCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                return functionReturnValue;
            }
        }
    }
}