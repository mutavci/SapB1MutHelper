using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SapB1MutHelper.Models;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using BoMessageTime = SAPbouiCOM.BoMessageTime;
using BoStatusBarMessageType = SAPbouiCOM.BoStatusBarMessageType;

namespace SapB1MutHelper
{
    public class TableCreation
    {
        public static long VRetVal;
        public static int VErrCode;
        public static string VErrMsg = "";


        public static bool CreateTable(string tableName, string tableDesc, BoUTBTableType tableType)
        {
            const bool functionReturnValue = false;

            try
            {
                if (!TableExists(tableName))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Creating Table " + tableName + " ...................", BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);
                    var vUserTableMd = (UserTablesMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserTables);
                    vUserTableMd.TableName = tableName;
                    vUserTableMd.TableDescription = tableDesc;
                    vUserTableMd.TableType = tableType;
                    VRetVal = vUserTableMd.Add();
                    if (VRetVal != 0)
                    {
                        Helper.OCompany.GetLastError(out VErrCode, out VErrMsg);
                        Application.SBO_Application.StatusBar.SetText(
                            "Failed to Create Table " + tableDesc + VErrCode + " " + VErrMsg, BoMessageTime.bmt_Short);
                        Marshal.ReleaseComObject(vUserTableMd);
                        return false;
                    }

                    Application.SBO_Application.StatusBar.SetText(
                        "[" + tableName + "] - " + tableDesc + " Created Successfully!!!", BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);
                    Marshal.ReleaseComObject(vUserTableMd);
                    return true;
                }

                GC.Collect();
                return false;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("" + ":> " + ex.Message + " @ " + ex.Source);
            }

            return functionReturnValue;
        }

        public static bool ColumnExists(string tableName, string fieldId)
        {
            try
            {
                var rs = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                var oFlag = true;
                rs.DoQuery("Select 1 from \"CUFD\" Where \"TableID\"='" + tableName.Trim() + "' and \"AliasID\"='" +
                           fieldId.Trim() + "'");
                if (rs.EoF)
                    oFlag = false;
                Marshal.ReleaseComObject(rs);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return oFlag;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message);
            }

            return true;
        }


        public static bool TableExists(string tableName)
        {
            return ((UserTablesMD)Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserTables)).GetByKey(tableName);
         /*   Marshal.ReleaseComObject(oTables);*/
            /*      Marshal.ReleaseComObject(oTables);*/
        }


        public static bool CreateUserFields(string tableName, string fieldName, string fieldDescription,
            BoFieldTypes type, long size = 0, BoFldSubTypes subType = BoFldSubTypes.st_None, string linkedTable = "",
            string defaultValue = "", List<ComboList> combolist = null, bool mandatory = false)
        {
            try
            {
                if (tableName.StartsWith("@"))
                    if (!ColumnExists(tableName, fieldName))
                    {
                        UserFieldsMD vUserField;
                        vUserField = (UserFieldsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                        vUserField.TableName = tableName;
                        vUserField.Name = fieldName;
                        vUserField.Description = fieldDescription;
                        vUserField.Type = type;
                        if (mandatory) vUserField.Mandatory = BoYesNoEnum.tYES;
                        if (type != BoFieldTypes.db_Date)
                            if (size != 0)
                                if (type == BoFieldTypes.db_Numeric)
                                    vUserField.EditSize = 11;
                                else
                                    vUserField.Size = (int) size;
                        if (subType != BoFldSubTypes.st_None) vUserField.SubType = subType;
                        if (!string.IsNullOrEmpty(linkedTable))
                            vUserField.LinkedTable = linkedTable;
                        if (!string.IsNullOrEmpty(defaultValue))
                            vUserField.DefaultValue = defaultValue;


                        if (combolist != null)
                            foreach (var item in combolist)
                            {
                                vUserField.ValidValues.Value = item.Value;
                                vUserField.ValidValues.Description = item.Description;
                                vUserField.ValidValues.Add();
                            }


                        VRetVal = vUserField.Add();
                        if (VRetVal != 0)
                        {
                            Helper.OCompany.GetLastError(out VErrCode, out VErrMsg);
                            Application.SBO_Application.StatusBar.SetText(
                                "Failed to add UserField masterid" + VErrCode + " " + VErrMsg,
                                BoMessageTime.bmt_Short);
                            Marshal.ReleaseComObject(vUserField);
                            return false;
                        }
                        else
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                "[" + tableName + "] - " + fieldDescription + " added successfully!!!",
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            Marshal.ReleaseComObject(vUserField);
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }

                if (tableName.StartsWith("@") == false)
                    if (!UdfExists(tableName, fieldName))
                    {
                        var vUserField = (UserFieldsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                        vUserField.TableName = tableName;
                        vUserField.Name = fieldName;
                        vUserField.Description = fieldDescription;
                        vUserField.Type = type;
                        if (type != BoFieldTypes.db_Date)
                            if (size != 0)
                                if (type == BoFieldTypes.db_Numeric)
                                    vUserField.EditSize = 11;
                                else
                                    vUserField.Size = (int) size;
                        if (subType != BoFldSubTypes.st_None) vUserField.SubType = subType;

                        //#region Geçerli Değerler

                        if (combolist != null)
                            foreach (var item in combolist)
                            {
                                vUserField.ValidValues.Value = item.Value;
                                vUserField.ValidValues.Description = item.Description;
                                vUserField.ValidValues.Add();
                            }

                        if (!string.IsNullOrEmpty(linkedTable))
                            vUserField.LinkedTable = linkedTable;
                        VRetVal = vUserField.Add();
                        if (VRetVal != 0)
                        {
                            Helper.OCompany.GetLastError(out VErrCode, out VErrMsg);
                            Application.SBO_Application.StatusBar.SetText(
                                "Failed to add UserField " + fieldDescription + " - " + VErrCode + " " + VErrMsg,
                                BoMessageTime.bmt_Short);
                            Marshal.ReleaseComObject(vUserField);
                            return false;
                        }
                        else
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                " & TableName & - " + fieldDescription + " added successfully!!!",
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            Marshal.ReleaseComObject(vUserField);
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                GC.Collect();
            }

            return true;
        }

        public static bool UdfExists(string tableName, string fieldId)
        {
            try
            {
                var rs = (Recordset) Helper.OCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                var oFlag = true;
                rs.DoQuery("Select 1 from \"CUFD\" Where \"TableID\"='" + tableName.Trim() + "' and \"AliasID\"='" +
                           fieldId.Trim() + "'");
                if (rs.EoF)
                    oFlag = false;
                Marshal.ReleaseComObject(rs);
                GC.Collect();
                return oFlag;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message);
            }

            return true;
        }


        public static void CreateUserFieldsFloat()
        {
            try
            {
                UserFieldsMD userFields;
                userFields = (UserFieldsMD) Helper.OCompany.GetBusinessObject(BoObjectTypes.oUserFields);
                userFields.TableName = "@OPTION";
                userFields.Name = "price1";
                userFields.Description = "price1 ack";
                userFields.Type = BoFieldTypes.db_Float;
                userFields.SubType = BoFldSubTypes.st_Price;
                userFields.EditSize = 20;
                var errCode = userFields.Add();

                if (errCode != 0)
                {
                    string errMsg;
                    Helper.OCompany.GetLastError(out errCode, out errMsg);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }
    }
}