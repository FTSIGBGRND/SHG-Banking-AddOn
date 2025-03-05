using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Data.Sql;
using System.Data.SqlClient;
//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : DI.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class DI : UserControl
    {
        private static long ErrNumber;
        private static string ErrMsg;
        private static long RetVal;
        private static SqlCommand mySqlCommand;

        public static SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
        public static SqlConnection mySqlConnection;

        public static long getnextkey(string as_code)
        {
            return getnextkey(as_code, false);
        }
        public static long getnextkey(string as_code, bool ab_transaction)
        {
            SAPbobsCOM.Recordset oRecordset;
            long ll_nkey;
            string ls_nkey;
            if (ab_transaction) oCompany.StartTransaction();

            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("select u_nkey from [@ftnkey] where code = '" + as_code + "'");
            if (oRecordset.RecordCount > 0)
            {
                ll_nkey = Int64.Parse(oRecordset.Fields.Item("u_nkey").Value.ToString());
                ls_nkey = System.Convert.ToString(ll_nkey + 1);
                oRecordset.DoQuery("Update [@ftnkey] set u_nkey=" + ls_nkey + " where code = '" + as_code + "' and  u_nkey=" + ll_nkey.ToString());
                oRecordset.DoQuery("select u_nkey from [@ftnkey] where code = '" + as_code + "' and  u_nkey=" + ls_nkey);
                if (oRecordset.RecordCount > 0)
                {
                    if (ab_transaction) DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    oRecordset = null;
                    return ll_nkey + 1;
                }
                return 0;
            }
            else
            {
                if (!DI.executeQuery("insert into [@ftnkey] (code,name,u_nkey) values('" + as_code + "','" + as_code + "',1)"))
                {
                    UI.SBO_Application.MessageBox("Error : " + DI.oCompany.GetLastErrorCode() + "/n Error Description : " +
                                                   DI.oCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    return 0;
                }
                oRecordset.DoQuery("select u_nkey from [@ftnkey] where code = '" + as_code + "' and u_nkey=1");
                if (oRecordset.RecordCount > 0)
                {
                    if (ab_transaction) DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    oRecordset = null;
                    return 1;
                }
                return 0;

            }
        }
        public static bool uploadReportType()
        {
            SAPbobsCOM.ReportTypesService rptTypeService;
            SAPbobsCOM.ReportType newType;
            SAPbobsCOM.Recordset oRecordSet;
            DataTable oDataRepType;
            DataRow[] oDataRT;

            string repname = "", reptype = "";

            oDataRepType = new DataTable("REPTYPE");
            oDataRepType.Columns.Add("row", typeof(System.String));
            oDataRepType.Columns.Add("RepCode", typeof(System.String));
            oDataRepType.Columns.Add("RepType", typeof(System.String));

            try
            {

                oDataRepType.Clear();
                oDataRepType.Rows.Add("0", "FTOOCW", "Check Writing Add On");

                rptTypeService = (SAPbobsCOM.ReportTypesService)DI.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);

                oDataRT = oDataRepType.Select();

                if (oDataRT.Length > 0)
                {
                    for (int i = 0; i < oDataRT.Length; i++)
                    {
                        repname = oDataRepType.Rows[i]["RepCode"].ToString();
                        reptype = oDataRepType.Rows[i]["RepType"].ToString();

                        oRecordSet = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            oRecordSet.DoQuery("select Code from RTYP where NAME = '" + repname + "'");
                            if (oRecordSet.RecordCount == 0)
                            {
                                newType = (SAPbobsCOM.ReportType)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);
                                newType.TypeName = repname;
                                newType.AddonName = reptype;
                                newType.AddonFormType = reptype;
                                newType.MenuID = reptype + "01";

                                rptTypeService.AddReportType(newType);

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                                oRecordSet = null;
                                GC.Collect();
                            }
                            else
                            {

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                                oRecordSet = null;
                                GC.Collect();
                            }
                        }
                        catch (Exception e)
                        {
                            GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                            oRecordSet = null;
                            GC.Collect();
                            return false;
                        }

                    }
                }
            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);

                GC.Collect();
                return false;
            }
            GC.Collect();
            return true;
        }
        public static bool uploadReportLayout(string as_repname, string as_reppath)
        {
            SAPbobsCOM.ReportTypesService rptTypeService;
            SAPbobsCOM.ReportType newType;
            SAPbobsCOM.ReportTypeParams newTypeParam;
            SAPbobsCOM.ReportLayoutsService rptService;
            SAPbobsCOM.ReportLayout newReport;
            SAPbobsCOM.ReportLayoutParams newReportParam;
            SAPbobsCOM.BlobParams oBlobParams;
            SAPbobsCOM.BlobTableKeySegment oKeySegment;
            SAPbobsCOM.Blob oBlob;
            SAPbobsCOM.Recordset oRecordSet;
            SAPbobsCOM.Recordset oRecordSet1;

            FileStream oFile;
            int fileSize = 0;
            byte[] buf;
            string typecode = "";

            try
            {
                rptTypeService = (SAPbobsCOM.ReportTypesService)DI.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);

                try
                {
                    oRecordSet = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordSet.DoQuery("select Code from RTYP where '" + as_repname + "' like Name + '%'");
                    if (oRecordSet.RecordCount > 0)
                    {
                        typecode = oRecordSet.Fields.Item("Code").Value.ToString();

                        oRecordSet1 = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet1.DoQuery("select DocCode from RDOC where DocName = '" + as_repname + "' and TypeCode = '" + typecode + "'");
                        if (oRecordSet1.RecordCount == 0)
                        {
                            newTypeParam = (SAPbobsCOM.ReportTypeParams)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportTypeParams);
                            newTypeParam.TypeCode = typecode;

                            newType = rptTypeService.GetReportType(newTypeParam);
                            rptService = (SAPbobsCOM.ReportLayoutsService)DI.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);

                            newReport = (SAPbobsCOM.ReportLayout)rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
                            newReport.Author = DI.oCompany.UserName;
                            newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
                            newReport.Name = as_repname;
                            newReport.TypeCode = newTypeParam.TypeCode;

                            newReportParam = rptService.AddReportLayout(newReport);

                            newType = rptTypeService.GetReportType(newTypeParam);
                            newType.DefaultReportLayout = newReportParam.LayoutCode;

                            rptTypeService.UpdateReportType(newType);

                            oBlobParams = (SAPbobsCOM.BlobParams)DI.oCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
                            oBlobParams.Table = "RDOC";
                            oBlobParams.Field = "Template";

                            oKeySegment = oBlobParams.BlobTableKeySegments.Add();
                            oKeySegment.Name = "DocCode";
                            oKeySegment.Value = newReportParam.LayoutCode;

                            oFile = new FileStream(as_reppath, System.IO.FileMode.Open);
                            fileSize = (int)oFile.Length;
                            buf = new byte[fileSize];

                            oFile.Read(buf, 0, fileSize);
                            oFile.Dispose();

                            oBlob = (SAPbobsCOM.Blob)DI.oCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);
                            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
                            DI.oCompany.GetCompanyService().SetBlob(oBlobParams, oBlob);
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                        oRecordSet = null;

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
                        oRecordSet1 = null;
                        GC.Collect();
                    }
                    else
                    {

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                        oRecordSet = null;
                        GC.Collect();
                    }
                }
                catch (Exception e)
                {
                    GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                    GC.Collect();
                    return false;
                }

            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);

                GC.Collect();
                return false;
            }
            GC.Collect();
            return true;
        }
        public static bool initreports()
        {
            SAPbobsCOM.ReportTypesService rptTypeService;
            SAPbobsCOM.ReportType newType;
            SAPbobsCOM.ReportTypeParams newTypeParam;
            SAPbobsCOM.ReportLayoutsService rptService;
            SAPbobsCOM.ReportLayoutParams newReportParam;
            SAPbobsCOM.Recordset oRecordSet;
            SAPbobsCOM.Recordset oRecordSet1;

            string typecode = "", layout = "";

            oRecordSet1 = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet1.DoQuery("select a.CODE from RTYP a where a.ADD_NAME like 'FTOOCW%'");
                if (oRecordSet1.RecordCount > 0)
                {
                    oRecordSet1.MoveFirst();
                    for (int r = 0; r < oRecordSet1.RecordCount; r++)
                    {
                        typecode = oRecordSet1.Fields.Item("CODE").Value.ToString();

                        rptTypeService = (SAPbobsCOM.ReportTypesService)DI.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                        newTypeParam = (SAPbobsCOM.ReportTypeParams)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportTypeParams);
                        newTypeParam.TypeCode = typecode;

                        newType = rptTypeService.GetReportType(newTypeParam);
                        newType.DefaultReportLayout = null;
                        rptTypeService.UpdateReportType(newType);

                        oRecordSet = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        oRecordSet.DoQuery("select a.DocCode from RDOC a inner join RTYP b on a.TypeCode = b.Code where a.TypeCode = '" + typecode + "'");
                        if (oRecordSet.RecordCount > 0)
                        {
                            oRecordSet.MoveFirst();
                            for (int i = 0; i < oRecordSet.RecordCount; i++)
                            {
                                layout = oRecordSet.Fields.Item("DocCode").Value.ToString();

                                rptService = (SAPbobsCOM.ReportLayoutsService)DI.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                                newReportParam = (SAPbobsCOM.ReportLayoutParams)rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams);
                                newReportParam.LayoutCode = layout;

                                rptService.DeleteReportLayout(newReportParam);

                                oRecordSet.MoveNext();
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                            oRecordSet = null;
                            GC.Collect();
                        }
                        else
                        {

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                            oRecordSet = null;
                            GC.Collect();
                        }

                        rptTypeService.DeleteReportType(newType);

                        oRecordSet1.MoveNext();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
                    oRecordSet1 = null;
                    GC.Collect();
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
                    oRecordSet1 = null;
                    GC.Collect();
                }
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
                oRecordSet1 = null;
                GC.Collect();
                return false;
            }
            return true;
        }
        public static bool execstoredproc(string path)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                string script = sr.ReadToEnd();

                //Server serverM = new Server(new ServerConnection(mySqlConnection));
                //serverM.ConnectionContext.ExecuteNonQuery(script);
                createStorepProcedure(script);
            }

            return true;
        }
        public static bool createStorepProcedure(string as_sql)
        {
            //string ls_sql;
            try
            {

                foreach (var batch in as_sql.Split(new string[] { "\nGO", "\ngo" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        new SqlCommand(batch, mySqlConnection).ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw;
                    }
                }
                //GlobalFunction.fileappend("Created Stored Procedure - " + as_sql);
                return true;
            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);
                return false;
            }
        }
        public static bool savetofile(string query, string filename, string ftype)
        {
            try
            {
                SAPbobsCOM.Recordset oRecordset;

                string ls_path = "", ls_filename;

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery("select isnull(a.AttachPath, '') as AttachPath from OADP a");
                if (oRecordset.RecordCount > 0)
                {
                    oRecordset.MoveFirst();
                    ls_path = oRecordset.Fields.Item("AttachPath").Value.ToString();
                }
                GC.Collect();

                if (string.IsNullOrEmpty(ls_path))
                {
                    UI.SBO_Application.StatusBar.SetText("The attachment folder has not been defined. Check Company Details in Administration >> System Initialization >> General Settings menu.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                {
                    if (!Directory.Exists(ls_path))
                    {
                        Directory.CreateDirectory(ls_path);
                    }
                }

                ls_filename = ls_path + filename + "." + ftype;
                
                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oRecordset.DoQuery(query);

                if (oRecordset.RecordCount != 0)
                {
                    UI.changeStatus("Saving file...");
                    oRecordset.MoveFirst();
                    using (StreamWriter sw = new StreamWriter(ls_filename))
                    {
                        for (int ctr = 1; ctr <= oRecordset.RecordCount; ctr++)
                        {
                            sw.WriteLine(oRecordset.Fields.Item(0).Value.ToString());
                            oRecordset.MoveNext();
                        }
                    }
                }
                else
                {
                    UI.SBO_Application.StatusBar.SetText("No data.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();

                UI.SBO_Application.StatusBar.SetText("File saved successfuly: " + ls_path, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                UI.hideStatus();

                return true;
            }
            catch (Exception e)
            {
                UI.hideStatus();
                GlobalFunction.f_error(e);
                return false;
            }
        }
        public static string getreportcode(string as_repname)
        {
            SAPbobsCOM.Recordset oRecordset;
            string reportcode = "";

            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT TOP 1 \"CODE\" AS \"TypeCode\" from RTYP WHERE \"NAME\" LIKE '" + as_repname + "%' ORDER BY \"CODE\" DESC");
            if (oRecordset.RecordCount > 0)
            {
                reportcode = oRecordset.Fields.Item("TypeCode").Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return reportcode;
        }
        public static bool checkreportcode(string as_repname)
        {
            SAPbobsCOM.Recordset oRecordset;

            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordset.DoQuery("select top 1 \"CODE\" as \"TypeCode\" from RTYP where \"ADD_NAME\" in ('Check Writing Add On') and \"CODE\" = '" + as_repname + "' order by \"CODE\" desc");
            if (oRecordset.RecordCount > 0)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
                return true;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return false;
        }
        public static bool createUDT(string as_tablename, string as_tabledescription, SAPbobsCOM.BoUTBTableType aole_tabletype)
        {

            RetVal = 0;
            SAPbobsCOM.UserTablesMD UserTablesMD;

            try
            {
                UserTablesMD = (SAPbobsCOM.UserTablesMD)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (UserTablesMD.GetByKey(as_tablename) == false)
                {
                    UI.changeStatus("Creating Table " + as_tablename + " - " + as_tabledescription + "...");

                    UserTablesMD.TableName = as_tablename;
                    UserTablesMD.TableDescription = as_tabledescription;
                    UserTablesMD.TableType = aole_tabletype;
                    RetVal = UserTablesMD.Add();
                    if (RetVal != 0)
                    {
                        ErrNumber = DI.oCompany.GetLastErrorCode();
                        ErrMsg = DI.oCompany.GetLastErrorDescription();
                        UI.SBO_Application.MessageBox("Add Table Failed~nTable Name: " + as_tablename + "~nTable Description: " + as_tabledescription + "~nTable Type: " + aole_tabletype.ToString() + "~nError No : " + ErrNumber.ToString() + "~nError Desciption : " + ErrMsg, 1, "Ok", "", "");

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                        UserTablesMD = null;
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                        UserTablesMD = null;
                        GC.Collect();
                        globalvar.gb_installedUDO = true;
                        return true;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserTablesMD);
                UserTablesMD = null;
                GC.Collect();
                return true;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                UserTablesMD = null;
                GC.Collect();
                return false;
            }
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, SAPbobsCOM.BoFieldTypes aole_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            long al_type;
            al_type = 0;
            UI.changeStatus("Creating Field " + as_tablename + "." + as_name + " - " + as_description + "...");

            switch (aole_type)
            {
                case SAPbobsCOM.BoFieldTypes.db_Alpha:
                    al_type = 0;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Date:
                    al_type = 3;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Float:
                    al_type = 4;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Memo:
                    al_type = 1;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Numeric:
                    al_type = 2;
                    break;
                default:
                    al_type = 0;
                    break;
            }

            return createUDF(as_tablename, as_name, as_description, al_type, ai_size, as_default, as_options, as_reltable);
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, SAPbobsCOM.BoFldSubTypes aole_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            long al_type;
            al_type = 0;
            UI.changeStatus("Creating Field " + as_tablename + "." + as_name + " - " + as_description + "...");
            switch (aole_type)
            {
                case SAPbobsCOM.BoFldSubTypes.st_Address:
                    al_type = 63;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Image:
                    al_type = 73;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Link:
                    al_type = 66;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Measurement:
                    al_type = 77;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_None:
                    al_type = 0;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Percentage:
                    al_type = 37;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Phone:
                    al_type = 35;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Price:
                    al_type = 80;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Quantity:
                    al_type = 81;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Rate:
                    al_type = 82;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Sum:
                    al_type = 83;
                    break;
                case SAPbobsCOM.BoFldSubTypes.st_Time:
                    al_type = 84;
                    break;
                default:
                    break;
            }
            return createUDF(as_tablename, as_name, as_description, al_type, ai_size, as_default, as_options, as_reltable);
        }
        public static bool createUDF(string as_tablename, string as_name, string as_description, long al_type, int ai_size, string as_default, string as_options, string as_reltable)
        {
            RetVal = 0;
            UI.changeStatus("Creating Field " + as_tablename + "." + as_name + " - " + as_description + "...");
            SAPbobsCOM.UserFieldsMD UserFieldsMD;
            int li_index;
            string ls_data;
            try
            {
                UserFieldsMD = (SAPbobsCOM.UserFieldsMD)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                UserFieldsMD.Description = as_description;
                UserFieldsMD.Name = as_name;
                UserFieldsMD.TableName = as_tablename;
                if (!string.IsNullOrEmpty(as_reltable))
                    UserFieldsMD.LinkedTable = as_reltable;

                switch (al_type)
                {
                    case 0:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                        if (ai_size > 0)
                        {
                            UserFieldsMD.EditSize = ai_size;
                        }
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 1:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                        break;
                    case 2:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric;
                        if (ai_size > 0)
                        {
                            UserFieldsMD.EditSize = ai_size;
                        }
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 3:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 4:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 77:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 37:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 80:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 81:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 82:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 83:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 84:
                        UserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date;
                        UserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time;
                        if (as_default != "")
                        {
                            UserFieldsMD.DefaultValue = as_default;
                        }
                        break;
                    case 35:
                        //SAPbobsCOM.BoFldSubTypes.st_Phone;
                        break;
                    case 63:
                        //SAPbobsCOM.BoFldSubTypes.st_Address;
                        break;
                    case 73:
                        //SAPbobsCOM.BoFldSubTypes.st_Image;
                        break;
                    case 66:
                        //SAPbobsCOM.BoFldSubTypes.st_Link;
                        break;
                    default:
                        break;

                }
                li_index = 0;
                while (as_options != "")
                {
                    ls_data = nv_string.of_getToken(ref as_options, ",");
                    if (ls_data != "")
                    {
                        li_index = li_index + 1;
                        if (li_index > 0)
                        {
                            UserFieldsMD.ValidValues.Add();
                            UserFieldsMD.ValidValues.SetCurrentLine(li_index);
                            //oUserObjectMD.FindColumns.Add();
                            //oUserObjectMD.FindColumns.SetCurrentLine(li_index);
                        }
                        UserFieldsMD.ValidValues.Value = nv_string.of_getToken(ref ls_data, "-");
                        UserFieldsMD.ValidValues.Description = nv_string.of_getToken(ref ls_data, "-");
                        //oUserObjectMD.FindColumns.ColumnAlias = ls_data;
                    }
                }
                RetVal = UserFieldsMD.Add();
                if (RetVal != 0)
                {
                    ErrNumber = DI.oCompany.GetLastErrorCode();
                    ErrMsg = DI.oCompany.GetLastErrorDescription();
                    UI.SBO_Application.MessageBox("Add Field Failed~nTable Name: " + as_tablename + "~nField Name: " + as_name + "~nField Description: " + as_description + "~nError No : " + ErrNumber.ToString() + "~nError Desciption : " + ErrMsg, 1, "Ok", "", "");

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD);
                    UserFieldsMD = null;
                    GC.Collect();
                    return false;
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(UserFieldsMD);
                    UserFieldsMD = null;
                    GC.Collect();
                    return true;
                }

                //UserFieldsMD = null;
                //return true;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                UserFieldsMD = null;
                GC.Collect();
                return false;
            }

        }
        public static bool isUDFexists(string as_tablename, string as_fieldname)
        {
            SAPbobsCOM.Recordset oRecordSet;
            oRecordSet = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet.DoQuery("Select \"AliasID\" from CUFD where \"TableID\" = '" + as_tablename + "' and \"AliasID\" = '" + as_fieldname + "'");
                if (oRecordSet.RecordCount == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oRecordSet = null;
                    GC.Collect();
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();

                return true;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
                return false;
            }
        }
        public static bool createUDO(string as_code, string as_name, SAPbobsCOM.BoUDOObjType al_objecttype, string as_tablename, string as_childtables, string as_findcolumns, Boolean ab_manageseries)
        {
            try
            {
                if (al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                {
                    return createUDO(as_code, as_name, al_objecttype, as_tablename, as_childtables, as_findcolumns, ab_manageseries, true, true, false, true);
                }
                else
                {
                    return createUDO(as_code, as_name, al_objecttype, as_tablename, as_childtables, as_findcolumns, ab_manageseries, false, false, true, true);
                }

            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                return false;
            }
            //return true;
        }
        public static bool createUDO(string as_code, string as_name, SAPbobsCOM.BoUDOObjType al_objecttype, string as_tablename, string as_childtables, string as_findcolumns, Boolean ab_manageseries, Boolean ab_cancel, Boolean ab_close, Boolean ab_delete, Boolean ab_logs)
        {
            try
            {
                SAPbobsCOM.UserObjectsMD oUserObjectMD;
                SAPbobsCOM.Recordset oRecordset;
                long ErrNumber;
                int li_index;
                string ErrMsg, ls_data;

                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                if (oUserObjectMD.GetByKey(as_code) == false)
                {
                    //CanCancel
                    if (ab_cancel)
                    {
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    //CanClose
                    if (ab_close)
                    {
                        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    //CanDelete
                    if (ab_delete)
                    {
                        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    if (ab_logs)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + as_tablename;
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

                    li_index = 0;
                    while (as_findcolumns != "")
                    {
                        ls_data = nv_string.of_getToken(ref as_findcolumns, ",");
                        if (ls_data != "")
                        {
                            li_index = li_index + 1;
                            if (li_index > 0)
                            {
                                oUserObjectMD.FindColumns.Add();
                                oUserObjectMD.FindColumns.SetCurrentLine(li_index);
                            }
                            oUserObjectMD.FindColumns.ColumnAlias = ls_data;
                        }
                    }

                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.LogTableName = "";
                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;

                    li_index = 0;
                    while (as_childtables != "")
                    {
                        ls_data = nv_string.of_getToken(ref as_childtables, ",");
                        if (ls_data != "")
                        {
                            li_index = li_index + 1;
                            if (li_index > 0)
                            {
                                oUserObjectMD.ChildTables.Add();
                                oUserObjectMD.ChildTables.SetCurrentLine(li_index);
                            }
                            oUserObjectMD.ChildTables.TableName = ls_data;
                        }
                    }


                    oUserObjectMD.ExtensionName = "";

                    if (ab_manageseries && al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                    {
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                    }

                    oUserObjectMD.Code = as_code;
                    oUserObjectMD.Name = as_name;
                    oUserObjectMD.ObjectType = al_objecttype;
                    oUserObjectMD.TableName = as_tablename;

                    if (oUserObjectMD.Add() != 0)
                    {
                        ErrNumber = DI.oCompany.GetLastErrorCode();
                        ErrMsg = DI.oCompany.GetLastErrorDescription();

                        UI.SBO_Application.MessageBox("Add UDO Failed" + System.Environment.NewLine + "Table Name: " + oUserObjectMD.TableName + System.Environment.NewLine + "UDO Name: " + oUserObjectMD.Code + System.Environment.NewLine + "UDO Description: " + oUserObjectMD.Name + System.Environment.NewLine + "Error No : " + ErrNumber.ToString() + System.Environment.NewLine + "Error Desciption : " + ErrMsg, 1, "Ok", "", "");
                        oUserObjectMD = null;
                        GC.Collect();
                        return false;
                    }
                    else
                    {
                        globalvar.gb_installedUDO = true;
                        if (ab_manageseries && al_objecttype == SAPbobsCOM.BoUDOObjType.boud_Document)
                        {
                            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery("Update NNM1 set Indicator='Default' where ObjectCode='" + as_code + "'");
                            oRecordset = null;
                            GC.Collect();
                        }
                    }
                }
                else
                {
                    oUserObjectMD = null;
                    GC.Collect();
                }

                GC.Collect();
                return true;

            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);
                return false;
            }

        }
        public static bool executeQuery(string as_string)
        {
            try
            {
                SAPbobsCOM.Recordset executeQuery;
                executeQuery = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                executeQuery.DoQuery(as_string);
                executeQuery = null;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                UI.SBO_Application.StatusBar.SetText(e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            return true;
        }
        public static void logoff()
        {
            RegistryKey objRegistryKey;
            string ls_sbolocation;

            objRegistryKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\InstallShield_{3B7CBDC4-20D1-4E0F-8E36-ADFED7E767E5}");
            ls_sbolocation = (string)objRegistryKey.GetValue("InstallLocation");
            if (File.Exists(ls_sbolocation + "SAP Business One.exe") == true)
            {
                System.Diagnostics.Process.Start(ls_sbolocation + "SAP Business One.exe");
                UI.SBO_Application.ActivateMenuItem("526");

            }
        }
        public static bool createStorepProcedure(string as_name, string as_sql)
        {
            string ls_sql;
            try
            {
                ls_sql = "IF  EXISTS (SELECT * FROM dbo.sysobjects " +
                         "WHERE [name] = '" + as_name + "') " +
                         "DROP PROCEDURE [dbo].[" + as_name + "] ";

                mySqlCommand = mySqlConnection.CreateCommand();
                mySqlCommand.CommandText = ls_sql;
                mySqlCommand.ExecuteNonQuery();
                mySqlCommand.Dispose();

                mySqlCommand = mySqlConnection.CreateCommand();
                mySqlCommand.CommandText = as_sql;
                mySqlCommand.ExecuteNonQuery();
                mySqlCommand.Dispose();
                GlobalFunction.fileappend("Created Stored Procedure - " + as_sql);
                return true;
            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);
                return false;
            }
        }
    }
}
