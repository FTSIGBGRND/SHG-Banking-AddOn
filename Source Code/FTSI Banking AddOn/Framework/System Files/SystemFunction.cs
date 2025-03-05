using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace AddOn
{
    public partial class SystemFunction : UserControl
    {
        public SystemFunction()
        {
            InitializeComponent();
        }
        public static void SetApplication(string sConnectionString)
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            //string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            //if (Environment.GetCommandLineArgs().Length > 1) sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //else sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            try
            {
                SboGuiApi.Connect(sConnectionString);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                System.Environment.Exit(0);
            }

            UI.SBO_Application = SboGuiApi.GetApplication(-1);
            UI.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            UI.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            UI.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            UI.SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
            UI.SBO_Application.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(SBO_Application_LayoutKeyEvent);
            UI.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);


        }

       
        public static int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;

            string sCookie = null;
            string sConnectionContext = null;

            // // First initialize the Company object

            DI.oCompany = new SAPbobsCOM.Company();

            // // Acquire the connection context cookie from the DI API.
            sCookie = DI.oCompany.GetContextCookie();

            // // Retrieve the connection context string from the UI API using the
            // // acquired cookie.

            try
            {
                sConnectionContext = UI.SBO_Application.Company.GetConnectionContext(sCookie);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                System.Environment.Exit(0);
            }
            // before setting the SBO Login Context make sure the company is not
            // connected

            if (DI.oCompany.Connected == true)
            {
                DI.oCompany.Disconnect();
            }

            // Set the connection context information to the DI API.

            setConnectionContextReturn = DI.oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }
        public static int ConnectToCompany()
        {
            int connectToCompanyReturn = 0;

            // // Establish the connection to the company database.
            connectToCompanyReturn = DI.oCompany.Connect();
            return connectToCompanyReturn;
        }
        public static bool checklicense(string as_addon)
        {
            //StreamReader SR;
            //string ls_text, ls_sbolocation, ls_hkey = "", ls_expiry, ls_key, ls_addon;
            //RegistryKey objRegistryKey;
            //SAPbobsCOM.Recordset oRecordset;

            //oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //oRecordset.DoQuery("Select U_HKEY,U_EXPIRY,U_ADDON,U_KEY from [@FTLIC]");
            //if (oRecordset.RecordCount > 0)
            //{
            //    ls_hkey = oRecordset.Fields.Item("U_HKEY").Value.ToString();
            //    ls_expiry = Convert.ToDateTime(oRecordset.Fields.Item("U_EXPIRY").Value).ToString();
            //    ls_key = oRecordset.Fields.Item("U_KEY").Value.ToString();
            //    ls_addon = oRecordset.Fields.Item("U_ADDON").Value.ToString();
            //    if (ls_addon != as_addon)
            //    {
            //        return false;
            //    }
            //    if (nv_string.encrypt(ls_hkey + as_addon + ls_expiry) != ls_key)
            //    {
            //        return false;
            //    }
            //    else
            //    {
            //        if (Convert.ToDateTime(ls_expiry) <= DateTime.Now)
            //        {
            //            return false;
            //        }
            //        else
            //        {
            //            return true;
            //        }
            //    }
            //}
            //else
            //{
            //    objRegistryKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\sap");
            //    ls_sbolocation = (string)objRegistryKey.GetValue("SAP Business One ServerTools");
            //    objRegistryKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\InstallShield_" + ls_sbolocation);
            //    ls_sbolocation = (string)objRegistryKey.GetValue("InstallLocation");
            //    SR = File.OpenText(ls_sbolocation + "\\License\\B1LicenseFile.txt");
            //    //MessageBox.Show(SR.ReadToEnd());
            //    ls_text = SR.ReadLine();
            //    if (string.IsNullOrEmpty(ls_text))
            //    {
            //        ls_text = "";
            //    }
            //    while (true)
            //    {   
            //        if (ls_text.Length > 0)
            //        {
            //            if (ls_text.Substring(0, 13) == "HARDWARE-KEY=")
            //            {
            //                ls_hkey = ls_text.Substring(13).Trim();
            //                break;
            //            }
            //        }
            //        ls_text = SR.ReadLine();
            //        if (string.IsNullOrEmpty(ls_text))
            //        {
            //            ls_text = "";
            //        }
            //    }
            //    SR.Close();
            //    ls_expiry = DateTime.Now.AddDays(30).ToString();
            //    ls_key = nv_string.encrypt(ls_hkey + as_addon + ls_expiry);
            //    if (!DI.oCompany.InTransaction)
            //        DI.oCompany.StartTransaction();
            //    //MessageBox.Show("INSERT INTO [@FT_LIC] (Code,Name,DocEntry,U_HKEY,U_ADDON,U_EXPIRY,U_KEY)values('1','1','" + ls_hkey + "','" + as_addon + "','" + ls_expiry + "','" + ls_key + "')");
            //    if (!DI.executeQuery("INSERT INTO [@FTLIC] (Code,Name,DocEntry,U_HKEY,U_ADDON,U_EXPIRY,U_KEY)values('1','1',1,'" + ls_hkey + "','" + as_addon + "','" + ls_expiry + "','" + ls_key + "')"))
            //    {
            //        string errmsg;
            //        errmsg = DI.oCompany.GetLastErrorDescription();

            //        UI.SBO_Application.StatusBar.SetText(errmsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //        if (DI.oCompany.InTransaction)
            //            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            //        return false;
            //    }
            //    if (DI.oCompany.InTransaction)
            //        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            //    UI.SBO_Application.MessageBox("You have 30 days trial period.", 1, "Ok", "", "");
            return true;
            //}
        }
        public static Boolean CreateStoredProc()
        {
            bool lb_return = true;
            //if (DI.oCompany.Server == "(local)")
            //{
                try
                {
                    createConn(DI.oCompany.CompanyDB);
                    lb_return = addon.onCreateStoredProcedure();
                    if (DI.mySqlConnection != null)
                    {
                        if (DI.mySqlConnection.State == ConnectionState.Open)
                        {
                            DI.mySqlConnection.Close();
                        }
                        DI.mySqlConnection.Dispose();
                    }
                }
                catch (Exception e)
                {
                    UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                    return false;
                }
            //}
            return lb_return;
        }
        public static SqlConnection createConn(string database)
        {
            // Here you define your server. Values can not be NULL.        //Database Server Name.
            //string myDSN = "SQLSERVER";

            //Local Server Name.
            string mySN = DI.oCompany.Server;
            //string myUserId = DI.oCompany.DbUserName;
            //string myPassword = "B1Admin";//DI.oCompany.DbPassword;
            //Define the type of security, 'TRUE' or 'FALSE'.
            //string mySecType = "TRUE";

            //Here you have your connection string you can edit it here.
            // Server = myServerAddress; Database = myDataBase; User ID = myUsername; Password = myPassword; Trusted_Connection = False;
            //string DI.mySqlConnectionString = ("Server=" + mySN + ";Database=" + database + ";User ID=" + myUserId + ";Password=" + myPassword + ";Trusted_Connection=False;");
            string mySqlConnectionString = ("Data Source=" + mySN + ";Initial Catalog=" + database + ";Integrated Security=SSPI;");
            //If you wish to use SQL security, well just make your own connection string...
            // I make sure I have declare what DI.mySqlConnection stand for.
            if (DI.mySqlConnection == null) { DI.mySqlConnection = new SqlConnection(); };

            // Since i will be reusing the connection I will try this it the connection dose not exist.
            if (DI.mySqlConnection.ConnectionString == string.Empty || DI.mySqlConnection.ConnectionString == null)
            {
                // I use a try catch stament cuz I use 2 set of arguments to connect to the database
                try
                {
                    //First I try with a pool of 5-40 and a connection time out of 4 seconds. then I open the connection.
                    DI.mySqlConnection.ConnectionString = "Min Pool Size=5;Max Pool Size=40;Connect Timeout=4;" + mySqlConnectionString + ";";
                    DI.mySqlConnection.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    //If it did not work i try not using the pool and I give it a 45 seconds timeout.
                    try
                    {
                        if (DI.mySqlConnection.State != ConnectionState.Closed)
                        {
                            DI.mySqlConnection.Close();
                        }
                        DI.mySqlConnection.ConnectionString = "Pooling=false;Connect Timeout=45;" + mySqlConnectionString + ";";
                        DI.mySqlConnection.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                return DI.mySqlConnection;
            }
            //Here if the connection exsist and is open i try this.
            if (DI.mySqlConnection.State != ConnectionState.Open)
            {
                try
                {
                    DI.mySqlConnection.ConnectionString = "Min Pool Size=5;Max Pool Size=40;Connect Timeout=4;" + mySqlConnectionString + ";";
                    DI.mySqlConnection.Open();
                }
                catch (Exception)
                {
                    if (DI.mySqlConnection.State != ConnectionState.Closed)
                    {
                        DI.mySqlConnection.Close();
                    }
                    DI.mySqlConnection.ConnectionString = "Pooling=false;Connect Timeout=45;" + mySqlConnectionString + ";";
                    DI.mySqlConnection.Open();
                }
            }
            return DI.mySqlConnection;
        }
        public static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Windows.Forms.Application.Exit();
                    GC.Collect();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    System.Windows.Forms.Application.Exit();
                    GC.Collect();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    System.Windows.Forms.Application.Exit();
                    GC.Collect();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    GC.Collect();
                    break;
            }
        }
        public static void SBO_Application_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {

            if (DI.checkreportcode(eventInfo.ReportCode))
            {
                eventInfo.LayoutKey = globalvar.DocKey;
            }
            BubbleEvent = true;
        }
        public static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent eventType, out bool BubbleEvent)
        {
            //globalvar.CurrentForm = UI.SBO_Application.Forms.ActiveForm;
            BubbleEvent = true;
            UI.ProccessMenuEvent(globalvar.CurrentForm.UniqueID, ref eventType, ref BubbleEvent);

            int formIndex;
            switch (eventType.MenuUID)
            {
                case "FTLIC":
                    if (eventType.BeforeAction)
                    {
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_lic();
                        globalvar.sboform[formIndex].createForm(formIndex);
                    }
                    break;
                default:
                    addon.onMenuEvent(globalvar.CurrentForm.UniqueID, ref eventType, ref BubbleEvent);
                    break;
            }
        }
        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                globalvar.CurrentForm = UI.SBO_Application.Forms.Item(FormUID);

                if (!UI.ProccessItemEvent(FormUID, ref pVal, ref BubbleEvent))
                {
                    addon.onItemEvent(FormUID, ref pVal, ref BubbleEvent);

                }

            }
            catch (Exception e)
            {
                string errmsg;
                int errcode;
                errcode = DI.oCompany.GetLastErrorCode();
                errmsg = DI.oCompany.GetLastErrorDescription();
                GlobalFunction.fileappend("Errmsg -" + e.Message + ". Errorcode - " + errcode.ToString() + "  Msg - " + errmsg);
                BubbleEvent = false;
            }
        }
        public static void SBO_Application_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            addon.onStatusBarEvent(Text, MessageType);
        }
        public static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            addon.onFormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
        }
    }
}
