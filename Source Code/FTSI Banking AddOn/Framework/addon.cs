using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using AddOn.Setup;
using AddOn.Outgoing_Check_Wrting;
using AddOn.Bank_Integration;



//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : addon.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class addon : UserControl
    {
        /***************Please dont delete or change this code...******************************************/
        
        #region System Code
        public addon()
        {
            InitializeComponent();
        }
        public addon(string sConnectionString)
        {
            InitializeComponent();
            
            SystemFunction.SetApplication(sConnectionString);

            if (UI.SBO_Application != null)
            {      
                UI.displayStatus();
                UI.changeStatus("Connecting to DI API.");
                //UI.SBO_Application.MessageBox("Add UDO Failed" + System.Environment.NewLine + "Table Name: " + System.Environment.NewLine + "UDO Name: " + System.Environment.NewLine + "UDO Description: "  + System.Environment.NewLine + "Error No : " + System.Environment.NewLine + "Error Desciption : ", 1, "Ok", "", "");
                if (!(SystemFunction.SetConnectionContext() == 0))
                {
                    UI.SBO_Application.MessageBox("Failed setting a connection to DI API", 1, "OK", "", "");
                    UI.hideStatus();
                    System.Environment.Exit(0); //  Terminating the Add-On Application
                }

                //DI.oCompany = (SAPbobsCOM.Company)UI.SBO_Application.Company.GetDICompany();

                if (!(SystemFunction.ConnectToCompany() == 0))
                {

                    //UI.SBO_Application.MessageBox("Failed connecting to the company's Data Base", 1, "Ok", "", "");
                    UI.SBO_Application.MessageBox("Failed connecting to the company's Data Base. \nError Code: " + DI.oCompany.GetLastErrorCode().ToString() + "\nError Description: " + DI.oCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                    UI.hideStatus();
                    System.Environment.Exit(0);
                }
                else
                {

                    globalvar.addondescription = "";

                    globalvar.userid = DI.oCompany.UserName;

                    onConnectToSBO(ref globalvar.addondescription);

                    UI.changeStatus("Connected to DI API.");
                    UI.changeStatus("Checking Add-on UDT's, Create if necessary");

                    DI.oCompany.StartTransaction();
                    /*
                     *  System Table
                     *  Please dont delete
                    */
                    #region @FTNKEY
                    if (DI.createUDT("FTNKEY", "Next Key", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                    {

                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("Failed in Creating User Define Tables", 1, "Ok", "", "");
                        System.Environment.Exit(0);
                    }
                    if (DI.createUDT("FTLIC", "LICENSE", SAPbobsCOM.BoUTBTableType.bott_MasterData) == false)
                    {
                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("Failed in Creating User Define Tables", 1, "Ok", "", "");
                        System.Environment.Exit(0);
                    }
                    #endregion
                    /*
                     *  End System Table
                    */
                    if (onInitTables() == false)
                    {
                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("Failed in Creating User Define Tables. Please see the Error Log File.", 1, "Ok", "", "");
                        System.Environment.Exit(0);
                    }
                    else
                    {
                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                    }


                    DI.oCompany.StartTransaction();
                    GC.Collect();
                    UI.changeStatus("Checking Add-on UDF's, Create if necessary");
                    /*
                     *  System Fields
                     *  Please dont delete
                    */
                    #region System Fields
                    if (DI.isUDFexists("@FTNKEY", "Nkey") == false)
                    {
                        if (DI.createUDF("@FTNKEY", "Nkey", "Next Key", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", "", "") == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Define Fields", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                    }
                    if (DI.isUDFexists("@FTLIC", "HKEY") == false)
                    {
                        if (DI.createUDF("@FTLIC", "HKEY", "Hardware Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Define Fields", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                    }
                    if (DI.isUDFexists("@FTLIC", "AddOn") == false)
                    {
                        if (DI.createUDF("@FTLIC", "AddOn", "AddOn", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "", "") == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Define Fields", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                    }
                    if (DI.isUDFexists("@FTLIC", "Expiry") == false)
                    {
                        if (DI.createUDF("@FTLIC", "Expiry", "Expiry", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Define Fields", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                    }
                    if (DI.isUDFexists("@FTLIC", "Key") == false)
                    {
                        if (DI.createUDF("@FTLIC", "Key", "Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 210, "", "", "") == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Define Fields", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                    }
                    #endregion
                    /*
                     *  End System Fields
                    */
                    if (onInitFields() == false)
                    {
                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("Failed in Creating User Define Fields. Please see the Error Log File.", 1, "Ok", "", "");
                        System.Environment.Exit(0);
                    }
                    else
                    {

                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                    }
                    GC.Collect();

                    DI.oCompany.StartTransaction();

                    globalvar.gb_installedUDO = false;
                    UI.changeStatus("Checking Add-on UDO's, Create if necessary");
                    //if (DI.createUDO("FTLIC", "", SAPbobsCOM.BoUDOObjType.boud_MasterData, "FTLIC", "", "Code", false, false, false, true) == false)
                    //{
                    //    if (DI.oCompany.InTransaction)
                    //    {
                    //        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    //    }
                    //    UI.hideStatus();
                    //    UI.SBO_Application.MessageBox("Failed in Creating User Define Objects", 1, "Ok", "", "");
                    //    System.Environment.Exit(0);
                    //}
                    if (onRegisterUDO() == false)
                    {
                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("Failed in Creating User Define Objects. Please see the Error Log File.", 1, "Ok", "", "");
                        System.Environment.Exit(0);
                    }
                    else
                    {

                        if (DI.oCompany.InTransaction)
                        {
                            DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                    }
                    //UI.changeStatus("Checking License.");
                    if (!SystemFunction.checklicense(globalvar.addondescription))
                    {
                        UI.hideStatus();
                        UI.SBO_Application.MessageBox("No License For this server. \n Please get a License File from fasttrack.", 1, "Ok", "", "");
                        /*
                         *  System Menu
                         *  Please dont delete
                        */
                        //43524

                        SAPbouiCOM.Menus m_menus;
                        m_menus = UI.SBO_Application.Menus;
                        if (!m_menus.Exists("FTLIC"))
                        {
                            m_menus = m_menus.Item("43524").SubMenus;
                            m_menus.Add("FTLIC", "FT - License Admin", SAPbouiCOM.BoMenuType.mt_STRING, 3);
                            m_menus = null;
                        }
                        /*
                         *  System Menu
                         *  Please dont delete
                        */
                        return;
                    }
                    else
                    {
                        GC.Collect();
                        //UI.changeStatus("Checking Add-on Stored Procedure, Create if necessary");
                        //if (CreateStoredProc() == false)
                        //{
                        //    UI.hideStatus();
                        //    UI.SBO_Application.MessageBox("Failed in Creating Stored Procedure", 1, "Ok", "", "");
                        //    System.Environment.Exit(0);
                        //}
                        if (globalvar.gb_installedUDO == true)
                        {
                            UI.hideStatus();
                            DI.logoff();
                            return;
                        }
                        UI.changeStatus("Updating Add-on " + globalvar.addondescription + " menus...");
                        /*
                         *  System Menu
                         *  Please dont delete
                        */
                        //43524

                        SAPbouiCOM.Menus m_menus;
                        m_menus = UI.SBO_Application.Menus;
                        if (!m_menus.Exists("FTLIC"))
                        {
                            m_menus = m_menus.Item("43524").SubMenus;
                            m_menus.Add("FTLIC", "FT - License Admin", SAPbouiCOM.BoMenuType.mt_STRING, 3);
                            m_menus = null;
                        }
                        /*
                         *  System Menu
                         *  Please dont delete
                        */
                        if (onInitMenus() == false)
                        {
                            if (DI.oCompany.InTransaction)
                            {
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Creating User Menus", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                        UI.changeStatus("Setting Filters");
                        if (onInitFilters() == false)
                        {
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Setting Filters", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                        UI.changeStatus("Executing stored procedure/s...");
                        if (onCreateStoredProcedure() == false)
                        {
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Executing stored procedure/s", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                        UI.changeStatus("Uploading Report Layout/s...");
                        if (onInitReports() == false)
                        {
                            UI.hideStatus();
                            UI.SBO_Application.MessageBox("Failed in Uploading Report/s", 1, "Ok", "", "");
                            System.Environment.Exit(0);
                        }
                        UI.hideStatus();
                    }
                    UI.SBO_Application.StatusBar.SetText("Fasttrack " + globalvar.addondescription + " Successfully Connected!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

               
            }
            else
            {
                UI.hideStatus();
                System.Windows.Forms.MessageBox.Show("Failed connecting to UI API");
            }

        }
        #endregion

        /**************************************************************************************************/

        public static Boolean onConnectToSBO(ref string addondescription)
        {
            addondescription = "Banking AddOn";
            return true;
        }
        public static void onItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int formIndex;

            if (pVal.Before_Action == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                switch (pVal.FormTypeEx)
                {
                    //case "139":
                    //    formIndex = UI.generateFormIndex();
                    //    globalvar.sboform[formIndex] = new nv_sysform139();
                    //    globalvar.sboform[formIndex].attachForm(formIndex, FormUID, ref BubbleEvent);
                    //    break;
                }
            }

        }
        public static Boolean onInitFields()
        {

            #region "SAP Table"

            if (DI.isUDFexists("OUSR", "RGenBankF") == false)
                if (DI.createUDF("OUSR", "RGenBankF", "Regenerate Bank File", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "Y- Yes, N - No", "") == false)
                    return false;

            if (DI.isUDFexists("DSC1", "RGLAcctCode") == false)
                if (DI.createUDF("DSC1", "RGLAcctCode", "ReClass G/L Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("DSC1", "RGLAcctName") == false)
                if (DI.createUDF("DSC1", "RGLAcctName", "ReClass G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("ODSC", "BankTemp") == false)
                if (DI.createUDF("ODSC", "BankTemp", "Bank Template", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "", "BDO - Banco De Oro, UB - Union Bank, MBTC - Metro Bank Trust Corp, CB - China Bank, " +
                                                                                                                "BPI - Bank of the Philippine Island, PNB - Philippine National Bank, RB - Robinson Bank," +
                                                                                                                "MB - MayBank, UCPB - United Coconut Planters Banks", "") == false)
                    return false;

            if (DI.isUDFexists("OWHT", "WTCode") == false)
                if (DI.createUDF("OWHT", "WTCode", "Withholding Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("OFPR", "Quarter") == false)
                if (DI.createUDF("OFPR", "Quarter", "Quarter", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "1 - 1st Quarter, 2 - 2nd Quarter, 3 - 3rd Quarter, 4 - 4th Quarter", "") == false)
                    return false;

            if (DI.isUDFexists("VPM1", "CheckNo") == false)
                if (DI.createUDF("VPM1", "CheckNo", "Check Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;


            #endregion

            #region "Setup"

            if (DI.isUDFexists("@FTOBAS", "GenApp") == false)
                if (DI.createUDF("@FTOBAS", "GenApp", "Generate Upon Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            #endregion

            #region "Outgoing - Check Writing"


            if (DI.isUDFexists("@FTOOCW", "TransType") == false)            
                if (DI.createUDF("@FTOOCW", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "S", "S - Vendor, A - Account,  C - Customer", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "CardCode") == false)            
                if (DI.createUDF("@FTOOCW", "CardCode", "CardCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)                
                   return false;
                
            if (DI.isUDFexists("@FTOOCW", "CardName") == false)            
                if (DI.createUDF("@FTOOCW", "CardName", "CardName", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOOCW", "PayName") == false)
                if (DI.createUDF("@FTOOCW", "PayName", "PayName", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;
                       
            if (DI.isUDFexists("@FTOOCW", "Bank") == false)          
                if (DI.createUDF("@FTOOCW", "Bank", "Bank", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)               
                    return false;
                 
            if (DI.isUDFexists("@FTOOCW", "Branch") == false)            
                if (DI.createUDF("@FTOOCW", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "Account") == false)          
                if (DI.createUDF("@FTOOCW", "Account", "Account No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)                
                    return false;
         
            if (DI.isUDFexists("@FTOOCW", "BGLAcctCode") == false)       
                if (DI.createUDF("@FTOOCW", "BGLAcctCode", "Bank G/L Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "BGLAcctName") == false)         
                if (DI.createUDF("@FTOOCW", "BGLAcctName", "Bank G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)               
                    return false;

            if (DI.isUDFexists("@FTOOCW", "RGLAcctCode") == false)
                if (DI.createUDF("@FTOOCW", "RGLAcctCode", "ReClass G/L Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "RGLAcctName") == false)
                if (DI.createUDF("@FTOOCW", "RGLAcctName", "ReClass G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "RBranch") == false)
                if (DI.createUDF("@FTOOCW", "RBranch", "Releasing Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "PntCntr") == false)
                if (DI.createUDF("@FTOOCW", "PntCntr", "Printing Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "CrsChk") == false)
                if (DI.createUDF("@FTOOCW", "CrsChk", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "N - No, Y - Yes", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "CheckNo") == false)
                if (DI.createUDF("@FTOOCW", "CheckNo", "Check No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "Status") == false)            
                if (DI.createUDF("@FTOOCW", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", "O - Open, A - Approved, X - Cancelled, D - Rejected, R - Released, C - Cleared, P - Posted", "") == false)          
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "DocDate") == false)       
                if (DI.createUDF("@FTOOCW", "DocDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)             
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "DueDate") == false)            
                if (DI.createUDF("@FTOOCW", "DueDate", "Due Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "TaxDate") == false)         
                if (DI.createUDF("@FTOOCW", "TaxDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)          
                    return false;

            if (DI.isUDFexists("@FTOOCW", "OVPMDocEnt") == false)
                if (DI.createUDF("@FTOOCW", "OVPMDocEnt", "OVPM DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "OVPMDocNum") == false)          
                if (DI.createUDF("@FTOOCW", "OVPMDocNum", "OVPM DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOOCW", "TransId") == false)
                if (DI.createUDF("@FTOOCW", "TransId", "Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "CanDate") == false)            
                if (DI.createUDF("@FTOOCW", "CanDate", "Canceled Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOOCW", "RelDate") == false)            
                if (DI.createUDF("@FTOOCW", "RelDate", "Release Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOOCW", "ClrDate") == false)
                if (DI.createUDF("@FTOOCW", "ClrDate", "Cleared Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "Comments") == false)          
                if (DI.createUDF("@FTOOCW", "Comments", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)      
                    return false;

            if (DI.isUDFexists("@FTOOCW", "BnkRmks") == false)
                if (DI.createUDF("@FTOOCW", "BnkRmks", "Bank Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "AppRmks") == false)
                if (DI.createUDF("@FTOOCW", "AppRmks", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "PrprdBy") == false)       
                if (DI.createUDF("@FTOOCW", "PrprdBy", "Prepared By", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)         
                    return false;
                
            if (DI.isUDFexists("@FTOOCW", "ApprvdBy") == false)            
                if (DI.createUDF("@FTOOCW", "ApprvdBy", "Approve By", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)                
                    return false;
                            
            if (DI.isUDFexists("@FTOOCW", "TotalDue") == false)            
                if (DI.createUDF("@FTOOCW", "TotalDue", "Total Amount Due", SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOOCW", "GenBankF") == false)
                if (DI.createUDF("@FTOOCW", "GenBankF", "Generated Bank File", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "EmailNotif") == false)
                if (DI.createUDF("@FTOOCW", "EmailNotif", "Email Notification", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "N - No, Y - Yes", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOOCW", "RetFleNme") == false)
                if (DI.createUDF("@FTOOCW", "RetFleNme", "Return FileName", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "Select") == false)
                if (DI.createUDF("@FTOCW1", "Select", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)             
                    return false;               

            if (DI.isUDFexists("@FTOCW1", "BaseType") == false)
                if (DI.createUDF("@FTOCW1", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "BaseEntry") == false)            
                if (DI.createUDF("@FTOCW1", "BaseEntry", "Base Document Enrty", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOCW1", "BaseNum") == false)
                if (DI.createUDF("@FTOCW1", "BaseNum", "Base Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "DocLine") == false)
                if (DI.createUDF("@FTOCW1", "DocLine", "Document Line", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "InstId") == false)
                if (DI.createUDF("@FTOCW1", "InstId ", "Installment ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "DocDate") == false)            
                if (DI.createUDF("@FTOCW1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)                
                    return false;

            if (DI.isUDFexists("@FTOCW1", "DueDate") == false)
                if (DI.createUDF("@FTOCW1", "DueDate", "Due Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTOCW1", "OverDue") == false)            
                if (DI.createUDF("@FTOCW1", "OverDue", "Overdue Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", "", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOCW1", "BalDue") == false)            
                if (DI.createUDF("@FTOCW1", "BalDue ", "Balance Due", SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", "", "") == false)                
                    return false;
                
            if (DI.isUDFexists("@FTOCW1", "TotPay") == false)
                if (DI.createUDF("@FTOCW1", "TotPay ", "Total Payment", SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", "", "") == false)        
                    return false;

            #endregion

            #region "Releasing Branch"

            if (DI.isUDFexists("@FTBARB", "Address") == false)
                if (DI.createUDF("@FTBARB", "Address", "Address", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (DI.isUDFexists("@FTBARB", "Type") == false)
                if (DI.createUDF("@FTBARB", "Type", "Releasing Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "RET - Head Office, C/R - Releasing Site", "") == false)
                    return false;

            #endregion

            return true;
        }
        public static Boolean onInitMenus()
        {

            SAPbouiCOM.Menus m_menus;
            SAPbouiCOM.Menus[] m_menus2;
            SAPbouiCOM.MenuItem innerMenu;
            int[] i;
            int a;

            m_menus = UI.SBO_Application.Menus;
            m_menus2 = new SAPbouiCOM.Menus[20];

            i = new int[20];
            for (a = 0; a < m_menus.Count; a++)
            {

                //Modules
                if (m_menus.Item(a).UID == "43520")
                {
                    m_menus2[1] = m_menus.Item(a).SubMenus;
                    for (i[1] = 0; i[1] < m_menus2[1].Count; i[1]++)
                    {

                        /******* New Menu - VivaMax AddOn ******************************************/
                        if (!m_menus.Exists("FTBANK"))
                            innerMenu = m_menus2[1].Add("FTBANK", "FTSI Banking AddOn", SAPbouiCOM.BoMenuType.mt_POPUP, 17);

                        for (i[1] = 0; i[1] < m_menus2[1].Count; i[1]++)
                        {
                            if (m_menus2[1].Item(i[1]).UID == "FTBANK")
                            {
                                m_menus2[2] = m_menus2[1].Item(i[1]).SubMenus;

                                if (!m_menus.Exists("FTSET"))
                                    innerMenu = m_menus2[2].Add("FTSET", "Setup", SAPbouiCOM.BoMenuType.mt_POPUP, 0);

                                //if (!m_menus.Exists("FTINCP"))
                                //    innerMenu = m_menus2[2].Add("FTINCP", "Incoming Payments", SAPbouiCOM.BoMenuType.mt_POPUP, 1);

                                if (!m_menus.Exists("FTOUTP"))
                                    innerMenu = m_menus2[2].Add("FTOUTP", "Outgoing Payments", SAPbouiCOM.BoMenuType.mt_POPUP, 2);

                                for (i[2] = 0; i[2] < m_menus2[2].Count; i[2]++)
                                {
                                    if (m_menus2[2].Item(i[2]).UID == "FTSET")
                                    {
                                        m_menus2[3] = m_menus2[2].Item(i[2]).SubMenus;

                                        if (!m_menus2[3].Exists("FTSET1"))
                                            innerMenu = m_menus2[3].Add("FTSET1", "Banking AddOn Setup", SAPbouiCOM.BoMenuType.mt_STRING, 0);

                                    }

                                    if (m_menus2[2].Item(i[2]).UID == "FTOUTP")
                                    {
                                        m_menus2[3] = m_menus2[2].Item(i[2]).SubMenus;

                                        if (!m_menus2[3].Exists("FTOCWA"))
                                            innerMenu = m_menus2[3].Add("FTOCWA", "Check Warehousing", SAPbouiCOM.BoMenuType.mt_POPUP, 0);

                                        for (i[3] = 0; i[3] < m_menus2[3].Count; i[3]++)
                                        {
                                            if (m_menus2[3].Item(i[3]).UID == "FTOCWA")
                                            {
                                                m_menus2[4] = m_menus2[3].Item(i[3]).SubMenus;

                                                if (!m_menus2[4].Exists("FTOCWA1"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA1", "Check Writing", SAPbouiCOM.BoMenuType.mt_STRING, 0);

                                                if (!m_menus2[4].Exists("FTOCWA2"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA2", "Check Approval", SAPbouiCOM.BoMenuType.mt_STRING, 1);

                                                if (!m_menus2[4].Exists("FTOCWA3"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA3", "Check Posting", SAPbouiCOM.BoMenuType.mt_STRING, 2);

                                                if (!m_menus2[4].Exists("FTOCWA4"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA4", "Check Releasing", SAPbouiCOM.BoMenuType.mt_STRING, 3);

                                                if (!m_menus2[4].Exists("FTOCWA5"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA5", "Check Clearing", SAPbouiCOM.BoMenuType.mt_STRING, 4);

                                                if (!m_menus2[4].Exists("FTOCWA6"))
                                                    innerMenu = m_menus2[4].Add("FTOCWA6", "E-Mail Notification", SAPbouiCOM.BoMenuType.mt_STRING, 5);

                                            }
                                        }

                                        if (!m_menus2[3].Exists("FTOBIS"))
                                            innerMenu = m_menus2[3].Add("FTOBIS", "Bank Integration Service", SAPbouiCOM.BoMenuType.mt_POPUP, 1);

                                        for (i[3] = 0; i[3] < m_menus2[3].Count; i[3]++)
                                        {
                                            if (m_menus2[3].Item(i[3]).UID == "FTOBIS")
                                            {
                                                m_menus2[4] = m_menus2[3].Item(i[3]).SubMenus;

                                                if (!m_menus2[4].Exists("FTOBIS1"))
                                                    innerMenu = m_menus2[4].Add("FTOBIS1", "Bank File Generator", SAPbouiCOM.BoMenuType.mt_STRING, 0);

                                                if (!m_menus2[4].Exists("FTOBIS2"))
                                                    innerMenu = m_menus2[4].Add("FTOBIS2", "Return File Upload", SAPbouiCOM.BoMenuType.mt_STRING, 1);

                                            }
                                        }
                                    }

                                    //if (m_menus2[2].Item(i[2]).UID == "FTINCP")
                                    //{
                                    //    m_menus2[3] = m_menus2[2].Item(i[2]).SubMenus;

                                    //    if (!m_menus2[3].Exists("FTICWA"))
                                    //        innerMenu = m_menus2[3].Add("FTICWA", "Check Warehousing", SAPbouiCOM.BoMenuType.mt_POPUP, 0);

                                    //    for (i[3] = 0; i[3] < m_menus2[3].Count; i[3]++)
                                    //    {
                                    //        if (m_menus2[3].Item(i[3]).UID == "FTICWA")
                                    //        {
                                    //            m_menus2[4] = m_menus2[3].Item(i[3]).SubMenus;

                                    //            if (!m_menus2[4].Exists("FTICWA1"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA1", "Check Receiving", SAPbouiCOM.BoMenuType.mt_STRING, 0);

                                    //            if (!m_menus2[4].Exists("FTICWA2"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA2", "Check Approval", SAPbouiCOM.BoMenuType.mt_STRING, 1);

                                    //            if (!m_menus2[4].Exists("FTICWA3"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA3", "Canceled Check", SAPbouiCOM.BoMenuType.mt_STRING, 2);

                                    //            if (!m_menus2[4].Exists("FTICWA4"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA4", "Bounce Check", SAPbouiCOM.BoMenuType.mt_STRING, 3);

                                    //            if (!m_menus2[4].Exists("FTICWA5"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA5", "Deposit Check", SAPbouiCOM.BoMenuType.mt_STRING, 4);

                                    //            if (!m_menus2[4].Exists("FTICWA6"))
                                    //                innerMenu = m_menus2[4].Add("FTICWA6", "Cleared Check", SAPbouiCOM.BoMenuType.mt_STRING, 5);

                                    //        }
                                    //    }
                                    //}
                                }
                            }
                        }

                        /***********************************************************************/

                    }
                }
            }

            GC.Collect();
            return true;
        }
        public static Boolean onInitTables()
        {

            if (DI.createUDT("FTOOCW", "Check Writing Header", SAPbobsCOM.BoUTBTableType.bott_Document) == false)
                return false;

            if (DI.createUDT("FTOCW1", "Check Writing Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines) == false)
                return false;

            if (DI.createUDT("FTBARB", "Banking AddOn Releasing Branch", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (DI.createUDT("FTBAPC", "Banking AddOn Printing Center", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (DI.createUDT("FTOBAS", "Banking AddOn Setup", SAPbobsCOM.BoUTBTableType.bott_MasterData) == false)
                return false;

            return true;
        }
        public static Boolean onInitReports()
        {
            string filename = "";

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Reports") == false)
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Reports");
            }

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Reports"))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\Reports");
                FileInfo[] Files = DirInfo.GetFiles("*.rpt");
                if (Files.Length > 0)
                {
                    if (!DI.initreports())
                    {
                        return false;
                    }

                    if (!DI.uploadReportType())
                    {
                        return false;
                    }
                    foreach (FileInfo file in Files)
                    {
                        filename = Path.GetFileNameWithoutExtension(file.Name);
                        if (!DI.uploadReportLayout(filename, DirInfo.FullName + "\\" + file.ToString()))
                        {
                            return false;
                        }
                        else
                        {
                            file.Delete();
                        }
                    }
                }
            }
            return true;
        }
        public static void onMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent eventType, ref bool BubbleEvent)
        {
            int formIndex;
            if (eventType.BeforeAction)
            {

                switch (eventType.MenuUID)
                {
                    case "FTSET1":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Setup();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA1":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckWriting();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA2":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckApproval();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA3":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckPosting();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA4":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckReleasing();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA5":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckClearing();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOCWA6":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_EmailNotification();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOBIS1":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_BankIntegration_BankFileGenerator();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                    case "FTOBIS2":
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_BankIntegration_BankFileUpload();
                        globalvar.sboform[formIndex].createForm(formIndex);
                        break;

                }

            }
        }
        public static Boolean onRegisterUDO()
        {

            if (DI.createUDO("FTOOCW", "Outgoing Check Writing", SAPbobsCOM.BoUDOObjType.boud_Document, "FTOOCW", "FTOCW1", "DocEntry", false, false, false, false, true) == false)
                return false;

            if (DI.createUDO("FTOBAS", "Banking AddOn Setup", SAPbobsCOM.BoUDOObjType.boud_MasterData, "FTOBAS", "", "Code", false, false, false, false, true) == false)
                return false;

            return true;
        }
        public static Boolean onInitFilters()
        {
            return true;
        }
        public static void onStatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
        }
        public static Boolean onCreateStoredProcedure()
        {
            //string filepath = "";

            //if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Queries") == false)
            //{
            //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Queries");
            //}

            //if (Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Queries"))
            //{
            //    DirectoryInfo DirInfo = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\Queries");
            //    FileInfo[] Files = DirInfo.GetFiles("*.sql");

            //    if (Files.Length > 0)
            //    {
            //        if (SystemFunction.CreateStoredProc())
            //        {
            //            foreach (FileInfo file in Files)
            //            {
            //                filepath = DirInfo.FullName + "\\" + file.Name;
            //                if (!DI.execstoredproc(filepath))
            //                {
            //                    return false;
            //                }
            //                else
            //                {
            //                    file.Delete();
            //                }
            //            }
            //        }
            //    }
            //}
            return true;
        }
        public static void onFormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            string strDocEntry, strDocNum, strOFPCDocNum, strABMCDocNum, strDrftNum, strTransType, strQuery;

            SAPbobsCOM.Recordset oRecordset, oRSFCP;
            SAPbobsCOM.Documents oDocuments;

            if (!BusinessObjectInfo.BeforeAction)
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.Type == "13")
                    {   

                        SAPbouiCOM.Form oForm;
                        oForm = UI.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);

                        oRSFCP = null;
                        oRSFCP = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRSFCP.DoQuery(string.Format("SELECT DISTINCT OINV.\"DocEntry\", OINV.\"DocNum\", INV1.\"U_OFPCDocNum\" " +
                                                         "FROM OINV INNER JOIN INV1 ON OINV.\"DocEntry\" = INV1.\"DocEntry\" " +
                                                         "WHERE OINV.\"DocEntry\" = (SELECT MAX(\"DocEntry\") FROM OINV) AND ISNULL(INV1.\"U_OFPCDocNum\", '') != '' "));

                        if (oRSFCP.RecordCount > 0)
                        {
                            oRSFCP.MoveFirst();

                            strDocNum = oRSFCP.Fields.Item("DocNum").Value.ToString();
                            strDocEntry = oRSFCP.Fields.Item("DocEntry").Value.ToString();

                            while (!(oRSFCP.EoF))
                            {
                                strOFPCDocNum = oRSFCP.Fields.Item("U_OFPCDocNum").Value.ToString();

                                strQuery = string.Format("UPDATE \"@FTOFPC\" SET \"U_Status\" = 'P', \"U_OINVDocNum\" = '{0}', \"U_OINVDocEnt\" = '{1}' WHERE \"DocNum\" = '{2}' ", strDocNum, strDocEntry, strOFPCDocNum);
                                if (!(DI.executeQuery(strQuery)))
                                {
                                    UI.SBO_Application.StatusBar.SetText("Error updating Freight Charges Computation.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    BubbleEvent = false;
                                }


                                oRecordset = null;
                                oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oRecordset.DoQuery(string.Format("SELECT \"U_ODLNDocEnt\" FROM \"@FTOFPC\" WHERE \"DocNum\" = '{0}' ", strOFPCDocNum));

                                if (oRecordset.RecordCount > 0)
                                {
                                    oDocuments = null;
                                    oDocuments = (SAPbobsCOM.Documents)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);

                                    if (oDocuments.GetByKey(Convert.ToInt32(oRecordset.Fields.Item("U_ODLNDocEnt").Value.ToString())))
                                        oDocuments.Close();

                                }

                                oRSFCP.MoveNext();

                            }
                        }

                    }                
                    else if (BusinessObjectInfo.Type == "18")
                    {
                        SAPbouiCOM.Form oForm;
                        oForm = UI.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);

                        strTransType = oForm.DataSources.DBDataSources.Item(0).GetValue("U_TransType", 0);

                        if (strTransType == "ISC")
                            strABMCDocNum = oForm.DataSources.DBDataSources.Item(0).GetValue("U_OISCDocNum", 0);
                        else
                            strABMCDocNum = oForm.DataSources.DBDataSources.Item(0).GetValue("U_OSMCDocNum", 0);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(string.Format("SELECT \"DocEntry\", \"DocNum\" FROM OPCH WHERE \"DocEntry\" = (SELECT MAX(\"DocEntry\") FROM OPCH) "));

                        if (oRecordset.RecordCount > 0)
                        {
                            strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
                            strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();

                            if (strTransType == "ISC")
                                strQuery = string.Format("UPDATE \"@FTOISC\" SET \"U_Status\" = 'P', \"U_OPCHDocNum\" = '{0}', \"U_OPCHDocEnt\" = '{1}' WHERE \"DocNum\" = '{2}' ", strDocNum, strDocEntry, strABMCDocNum);
                            else
                                strQuery = string.Format("UPDATE \"@FTOSMC\" SET \"U_Status\" = 'P', \"U_OPCHDocNum\" = '{0}', \"U_OPCHDocEnt\" = '{1}' WHERE \"DocNum\" = '{2}' ", strDocNum, strDocEntry, strABMCDocNum);

                            if (!(DI.executeQuery(strQuery)))
                            {
                                UI.SBO_Application.StatusBar.SetText("Error updating Insurance and Storage/Surrender to Mill Computation.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                if (DI.oCompany.InTransaction)
                                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }
        }

    }
}