using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SAPbouiCOM;
using SAPbobsCOM;

namespace AddOn.Outgoing_Check_Wrting
{
    public partial class userform_Outgoing_CheckApproval : AddOn.Form
    {
        bool blSelect = false;
        private static string strUserCode;

        private static string strDBType;

        private static System.Data.DataTable oDTBank;
        private static System.Data.DataTable oDTHeader;
        private static System.Data.DataTable oDTDetails;

        private static DataRow[] oDRHeader, oDRDetails;
        public userform_Outgoing_CheckApproval()
        {
            InitializeComponent();
        }
        public override void choosefromlist(string FormUID, ref ChooseFromListEvent pVal, ref bool BubbleEvent)
        {
            base.choosefromlist(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.DataTable oDatatable;

            string strCardCode, strCardName;

            if (!pVal.BeforeAction)
            {
                if (pVal.SelectedObjects != null)
                {
                    oDatatable = pVal.SelectedObjects;

                    switch (pVal.ItemUID)
                    {
                        case "CardCode":

                            oForm.Freeze(true);

                            strCardName = oDatatable.GetValue("CardName", 0).ToString();
                            strCardCode = oDatatable.GetValue("CardCode", 0).ToString();

                            oForm.DataSources.UserDataSources.Item("CardCode").ValueEx = strCardCode;
                            oForm.DataSources.UserDataSources.Item("CardName").ValueEx = strCardName;

                            oForm.Update();
                            oForm.Freeze(false);

                            break;
                    }
                }
            }
            else
            {
                
                switch (pVal.ItemUID)
                {
                    case "CardCode":

                        oForm.Freeze(true);

                        uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;
                }
            }


            GC.Collect();
        }
        public override void doubleclick(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            base.click(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.Grid oGrid;

            if (!pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "grd1":

                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
                        switch (pVal.ColUID)
                        {
                            case "Selected":

                                oForm.Freeze(true);

                                if (blSelect == true)
                                    blSelect = false;
                                else
                                    blSelect = true;

                                if (pVal.Row == -1)
                                {
                                    for (int li_row = 0; li_row <= oGrid.Rows.Count - 1; li_row++)
                                    {
                                        if (blSelect == true)
                                            oGrid.DataTable.Columns.Item("Selected").Cells.Item(li_row).Value = "Y";
                                        else
                                            oGrid.DataTable.Columns.Item("Selected").Cells.Item(li_row).Value = "N";

                                    }
                                }

                                oForm.Update();
                                oForm.Freeze(false);

                                break;
                        }

                        break; 
                }
            }
        }
        public override void itempressed(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);

            string strFDate, strTDate;

            DateTime dteFDoc, dteTDoc;

            if (!pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "btnList":

                        oForm.Freeze(true);

                        strFDate = getItemString("FDocDate");
                        if (string.IsNullOrEmpty(strFDate))
                            strFDate = "01/01/1900";

                        strTDate = getItemString("TDocDate");
                        if (string.IsNullOrEmpty(strTDate))
                            strTDate = "12/31/2099";

                        dteFDoc = Convert.ToDateTime(strFDate);
                        dteTDoc = Convert.ToDateTime(strTDate);

                        uf_LoadCheck(getItemString("CardCode"), dteFDoc, dteTDoc);

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "btnApp":

                        oForm.Freeze(true);

                        uf_Approved("A", getItemString("Remarks"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "btnRej":

                        oForm.Freeze(true);

                        if (string.IsNullOrEmpty(getItemString("Remarks")))
                        {
                            UI.SBO_Application.StatusBar.SetText("Remarks for Rejected Check Writing is Required.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                        else
                            uf_Approved("D", getItemString("Remarks"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                }

            }

            GC.Collect();
        }
        public override void matrixlinkpressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.matrixlinkpressed(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.EditTextColumn oLink;
            int formIndex;
            string ls_docno;
            if (pVal.BeforeAction)
            {
                if (pVal.ItemUID == "grd1")
                {

                    oGrid = (SAPbouiCOM.Grid)getItemSpecific("grd1");
                    if (pVal.ColUID == "DocEntry")
                    {
                        oLink = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                        ls_docno = oLink.GetText(pVal.Row).Trim();
                        formIndex = UI.generateFormIndex();
                        globalvar.sboform[formIndex] = new userform_Outgoing_CheckWriting(ls_docno);
                        globalvar.sboform[formIndex].createForm(formIndex);
                        BubbleEvent = false;
                    }
                }
            }

            GC.Collect();
        }
        public override void onGetCreationParams(ref BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
            base.onGetCreationParams(ref io_BorderStyle, ref is_FormType, ref is_ObjectType, ref xmlPath);

            is_FormType = "100000002";
            io_BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

            GC.Collect();
        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);

            string strUserId;

            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.EditText oEditText;

            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            SAPbouiCOM.Condition oCon;
            SAPbouiCOM.Conditions oCons;

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)UI.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.UniqueID = "cflBP";
            oCFLCreationParams.ObjectType = "2";
            oCFL = oCFLs.Add(oCFLCreationParams);

            oForm.Freeze(true);

            strDBType = DI.oCompany.DbServerType.ToString();

            oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 250);
            oForm.DataSources.UserDataSources.Add("FDocDate", SAPbouiCOM.BoDataType.dt_DATE, 0);
            oForm.DataSources.UserDataSources.Add("TDocDate", SAPbouiCOM.BoDataType.dt_DATE, 0);
            oForm.DataSources.UserDataSources.Add("Remarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 250);

            oForm.Title = "Check Approval";
            oForm.Width = 1200;
            oForm.Height = 570;

            oItem = createEditText(150, 15, 150, 14, "CardCode", true, "", "CardCode");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.ChooseFromListUID = "cflBP";
            oEditText.ChooseFromListAlias = "CardCode";
            oItem = createLinkButton("lnkBP", oItem, SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
            oItem = createEditText(301, 15, 200, 14, "CardName", true, "", "CardName");
            oItem.Enabled = false;
            oItem = createStaticText(6, 15, 100, 14, "stCustomer", "CardCode", "CardCode");

            oItem = createEditText(150, 30, 150, 14, "FDocDate", true, "", "FDocDate");
            oItem = createStaticText(6, 30, 100, 14, "stFDocDate", "Posting Date From", "FDocDate");

            oItem = createEditText(150, 45, 150, 14, "TDocDate", true, "", "TDocDate");
            oItem = createStaticText(6, 45, 100, 14, "stTDocDate", "Posting Date To", "TDocDate");

            oItem = createButton(150, 65, 150, 19, "btnList", "&List");

            oItem = oForm.Items.Add("grd1", SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Enabled = true;
            oItem.Left = 6;
            oItem.Top = 100;
            oItem.Width = 1180;
            oItem.Height = 300;

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("CHKAPP");

            uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

            oItem = createExtEditText(150, 415, 250, 70, "Remarks", true, "", "Remarks");
            oItem = createStaticText(6, 410, 120, 14, "stRemarks", "Remarks", "Remarks");

            strUserId = DI.oCompany.UserSignature.ToString();

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)AddOn.DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"USER_CODE\" FROM OUSR WHERE \"USERID\" = '" + strUserId + "' ");

            strUserCode = oRecordset.Fields.Item("USER_CODE").Value.ToString();

            oItem = createButton(6, 505, 100, 19, "btnApp", "&Approved");
            oItem = createButton(110, 505, 100, 19, "btnRej", "&Rejected");
            oItem = createButton(300, 505, 100, 19, "2", "");

            GC.Collect();

        }
        public override void validate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.validate(FormUID, ref pVal, ref BubbleEvent);

            string strCardCode;

            if (!pVal.BeforeAction && pVal.ItemChanged)
            {
                switch (pVal.ItemUID)
                {
                    case "CardCode":

                        oForm.Freeze(true);

                        strCardCode = getItemString("CardCode");

                        if (string.IsNullOrEmpty(strCardCode))
                        {
                            oForm.DataSources.UserDataSources.Item("CardName").ValueEx = "";
                            uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));
                        }


                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "FDocDate":
                    case "TDocDate":

                        oForm.Freeze(true);

                        uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;
                }
            }

            GC.Collect();
        }
        private void uf_LoadCheck(string strCardCode, DateTime dteFDoc, DateTime dteTDoc)
        {
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.CheckBoxColumn oCCheckBox;
            SAPbouiCOM.EditTextColumn oLink;

            string strQuery;

            oForm.Freeze(true);

            if (string.IsNullOrEmpty(strCardCode))
            {
                if (strDBType == "dst_HANADB")
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKAPPROVAL\" (to_date('{0}', 'MM/DD/YYYY'), to_date('{1}', 'MM/DD/YYYY')) ", dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKAPPROVAL '{0}', '{1}' ", dteFDoc, dteTDoc);
            }
            else
            {
                if (strDBType == "dst_HANADB")
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKAPPROVAL_BP\" ('{0}', to_date('{1}', 'MM/DD/YYYY'), to_date('{2}', 'MM/DD/YYYY')) ", strCardCode, dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKAPPROVAL_BP '{0}', '{1}', '{2}' ", strCardCode, dteFDoc, dteTDoc);
            }

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            oForm.DataSources.DataTables.Item("CHKAPP").ExecuteQuery(strQuery);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("CHKAPP");

            oGrid.Columns.Item("Selected").TitleObject.Caption = "";
            oGrid.Columns.Item("Selected").Editable = true;
            oGrid.Columns.Item("Selected").Width = 20;
            oGrid.Columns.Item("Selected").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            oCCheckBox = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Selected");

            oGrid.Columns.Item("DocEntry").TitleObject.Caption = "";
            oGrid.Columns.Item("DocEntry").Width = 20;
            oGrid.Columns.Item("DocEntry").Editable = false;
            oLink = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
            oLink.LinkedObjectType = "18";

            oGrid.Columns.Item("DocNum").Width = 110;
            oGrid.Columns.Item("DocNum").Editable = false;
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Document No.";

            oGrid.Columns.Item("U_DocDate").Width = 110;
            oGrid.Columns.Item("U_DocDate").Editable = false;
            oGrid.Columns.Item("U_DocDate").TitleObject.Caption = "Document Date";

            oGrid.Columns.Item("U_DueDate").Width = 110;
            oGrid.Columns.Item("U_DueDate").Editable = false;
            oGrid.Columns.Item("U_DueDate").TitleObject.Caption = "Due Date";

            //oGrid.Columns.Item("U_CardCode").Width = 120;
            //oGrid.Columns.Item("U_CardCode").Editable = false;
            //oGrid.Columns.Item("U_CardCode").TitleObject.Caption = "Customer Code";

            oGrid.Columns.Item("U_CardName").Width = 180;
            oGrid.Columns.Item("U_CardName").Editable = false;
            oGrid.Columns.Item("U_CardName").TitleObject.Caption = "Customer Name";

            //oGrid.Columns.Item("U_PayName").Width = 200;
            //oGrid.Columns.Item("U_PayName").Editable = false;
            //oGrid.Columns.Item("U_PayName").TitleObject.Caption = "Pay To Name";

            oGrid.Columns.Item("U_Bank").Width = 150;
            oGrid.Columns.Item("U_Bank").Editable = false;
            oGrid.Columns.Item("U_Bank").TitleObject.Caption = "Bank";

            oGrid.Columns.Item("U_BankTemp").Width = 0;
            oGrid.Columns.Item("U_BankTemp").Editable = false;
            oGrid.Columns.Item("U_BankTemp").Visible = false;
            oGrid.Columns.Item("U_BankTemp").TitleObject.Caption = "Bank Template";

            oGrid.Columns.Item("U_Branch").Width = 162;
            oGrid.Columns.Item("U_Branch").Editable = false;
            oGrid.Columns.Item("U_Branch").TitleObject.Caption = "Branch";

            oGrid.Columns.Item("U_Account").Width = 120;
            oGrid.Columns.Item("U_Account").Editable = false;
            oGrid.Columns.Item("U_Account").TitleObject.Caption = "Account";

            oGrid.Columns.Item("U_TotalDue").Width = 150;
            oGrid.Columns.Item("U_TotalDue").RightJustified = true;
            oGrid.Columns.Item("U_TotalDue").Editable = false;
            oGrid.Columns.Item("U_TotalDue").TitleObject.Caption = "Total Due";

            oForm.Freeze(false);
            GC.Collect();
        }
        private void uf_Approved(string strStatus, string strAppRmks)
        {
            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Grid oGrid;

            string  strDocEntry, strDocNum, strBankTemp, strAccount,
                    strQuery = "";

            bool blProcess = false;

            oForm.Freeze(true);

            if (!DI.oCompany.InTransaction)
                DI.oCompany.StartTransaction();

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            if (oGrid.Rows.Count == 0)
            {
                UI.SBO_Application.StatusBar.SetText("Check for Approval is missing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return;
            }
            else
            {
                uf_initDataTable();

                for (int ctr = 0; ctr < oGrid.Rows.Count; ctr++)
                {
                    if (oGrid.DataTable.Columns.Item("Selected").Cells.Item(ctr).Value.ToString() == "Y")
                    {
                        strDocEntry = oGrid.DataTable.Columns.Item("DocEntry").Cells.Item(ctr).Value.ToString();
                        strDocNum = oGrid.DataTable.Columns.Item("DocNum").Cells.Item(ctr).Value.ToString();
                        strBankTemp = oGrid.DataTable.Columns.Item("U_BankTemp").Cells.Item(ctr).Value.ToString();
                        strAccount = oGrid.DataTable.Columns.Item("U_Account").Cells.Item(ctr).Value.ToString();

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery("SELECT \"U_GenApp\" FROM \"@FTOBAS\" ");

                        if (oRecordset.Fields.Item("U_GenApp").Value.ToString() == "Y")
                        {
                            strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = '{0}', \"U_GenBankF\" = 'Y', \"U_AppRmks\" = '{1}', \"U_ApprvdBy\" = '{2}' WHERE \"DocEntry\" = '{3}' ", strStatus, strAppRmks, strUserCode, strDocEntry);
                            if (!DI.executeQuery(strQuery))
                            {
                                UI.SBO_Application.StatusBar.SetText("Error updating selected document(s). Failed to approve/reject.", SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                blProcess = false;
                                break;
                            }

                            if (!(string.IsNullOrEmpty(strBankTemp)))
                            {
                                oDTBank.Rows.Add(strBankTemp, strAccount, strDocEntry);
                                blProcess = true;
                            }
                            else
                            {
                                UI.SBO_Application.StatusBar.SetText("Selected document(s) failed to generate bank file. Please setup Bank Template.", SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                blProcess = false;
                                break;
                            }
                        }
                        else
                        {
                            strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = '{0}', \"U_AppRmks\" = '{1}', \"U_ApprvdBy\" = '{2}' WHERE \"DocEntry\" = '{3}' ", strStatus, strAppRmks, strUserCode, strDocEntry);
                            if (!DI.executeQuery(strQuery))
                            {
                                UI.SBO_Application.StatusBar.SetText("Error updating selected document(s). Failed to approve/reject.", SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                blProcess = false;
                                break;
                            }

                            blProcess = true;
                        }
                    }
                }
            }

            if (blProcess == true)
            {
                if (uf_generateFile())
                {
                    if (DI.oCompany.InTransaction)
                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    UI.SBO_Application.StatusBar.SetText("Selected document(s) successfully approved/rejected.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    setItemString("FDocDate", "");
                    setItemString("TDocDate", "");
                    setItemString("Remarks", "");

                    uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));
                }
                else
                {
                    if (DI.oCompany.InTransaction)
                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    UI.SBO_Application.StatusBar.SetText("Error creating Bank File.", SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            else
            {
                if (DI.oCompany.InTransaction)
                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }

            blProcess = false;

            oForm.Freeze(false);

            GC.Collect();

        }
        private static bool uf_generateFile()
        {

            System.Data.DataTable oDTBankFile;

            string strBankTemp, strAccount;

            if (oDTBank.Rows.Count > 0)
            {
                oDTBankFile = oDTBank.DefaultView.ToTable(true, "Bank", "Account");

                if (oDTBank.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow <= oDTBankFile.Rows.Count - 1; intRow++)
                    {
                        strBankTemp = oDTBankFile.Rows[intRow]["Bank"].ToString();
                        strAccount = oDTBankFile.Rows[intRow]["Account"].ToString();

                        if (strBankTemp == "UB")
                            uf_generateFileUB(strBankTemp, strAccount);
                        else
                            return false;
                    }
                }

            }

            return true;

        }
        private static bool uf_generateFileUB(string strBankTemp, string strAccount)
        {
            SAPbobsCOM.Recordset oRecordset;

            System.Data.DataTable oDTBankFile;

            string strDocEntry, strType, strData, strCompany, strFilePath, strFileName;
            string strQuery;

            double dblChkAmt;

            int intDTHRow = -1, intDRRow, intCtr, intChkCtr;

            try
            {
                strQuery = string.Format("Bank = '{0}' AND Account = '{1}' ", strBankTemp, strAccount);
                oDTBankFile = oDTBank.Select(strQuery).CopyToDataTable().DefaultView.ToTable();

                if (oDTBankFile.Rows.Count > 0)
                {
                    oDTHeader.Clear();

                    strQuery = ("SELECT \"AliasName\" FROM OADM");

                    oRecordset = null;
                    oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery(strQuery);

                    strCompany = oRecordset.Fields.Item("AliasName").Value.ToString();
                    strCompany = oRecordset.Fields.Item("AliasName").Value.ToString();
                    if (string.IsNullOrEmpty(strCompany))
                    {
                        UI.SBO_Application.StatusBar.SetText("Please Setup Company at Company Details > Alias Name.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                    {
                        for (int intRow = 0; intRow <= oDTBankFile.Rows.Count - 1; intRow++)
                        {
                            intChkCtr = intRow + 1;

                            strDocEntry = oDTBankFile.Rows[intRow]["DocEntry"].ToString();

                            if (strDBType == "dst_HANADB")
                                strQuery = string.Format("CALL \"FTSI_BANKINGADDON_EXPORT_OUTGOING_BANKINTEGRATION_LINEDETAILS_UB\" ('{0}', '{1}' ) ", strDocEntry, strAccount);
                            else
                                strQuery = string.Format("EXEC FTSI_BANKINGADDON_EXPORT_OUTGOING_BANKINTEGRATION_LINEDETAILS_UB '{0}', '{1}' ", strDocEntry, strAccount);

                            oRecordset = null;
                            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery(strQuery);

                            if (oRecordset.RecordCount > 0)
                            {
                                oRecordset.MoveFirst();

                                while (!(oRecordset.EoF))
                                {
                                    dblChkAmt = System.Convert.ToDouble(oRecordset.Fields.Item("U_TotalDue").Value.ToString());
                                    strType = oRecordset.Fields.Item("Type").Value.ToString();

                                    strData = oRecordset.Fields.Item("LineDetails").Value.ToString();

                                    oDRHeader = oDTHeader.Select("Bank = '" + strBankTemp + "' AND Account = '" + strAccount + "' ");
                                    if (oDRHeader.Length > 0)
                                    {
                                        intDRRow = Convert.ToInt32(oDRHeader[0]["Row"]);

                                        oDTHeader.Rows[intDRRow]["Amount"] = dblChkAmt + Convert.ToDouble(oDTHeader.Rows[intDRRow]["Amount"]);
                                        oDTHeader.Rows[intDRRow]["Counter"] = intChkCtr;
                                    }
                                    else
                                    {
                                        intDTHRow++;
                                        oDTHeader.Rows.Add(intDTHRow, strBankTemp, strAccount, intChkCtr, dblChkAmt);

                                    }

                                    oDTDetails.Rows.Add(intDTHRow, strType, strData);

                                    oRecordset.MoveNext();
                                }
                            }
                        }

                        if (oDTHeader.Rows.Count > 0)
                        {
                            strFilePath = string.Format(@"C:\Fasttrack Banking AddOn\{0}\", strCompany);
                            if (!Directory.Exists(strFilePath))
                                Directory.CreateDirectory(strFilePath);

                            for (int intRowH = 0; intRowH <= oDTHeader.Rows.Count - 1; intRowH++)
                            {
                                intDTHRow = Convert.ToInt32(oDTHeader.Rows[intRowH]["Row"].ToString());

                                strBankTemp = oDTHeader.Rows[intRowH]["Bank"].ToString();
                                strAccount = oDTHeader.Rows[intRowH]["Account"].ToString().Replace("-", "");

                                intCtr = Convert.ToInt32(oDTHeader.Rows[intRowH]["Counter"].ToString());
                                dblChkAmt = Convert.ToDouble(oDTHeader.Rows[intRowH]["Amount"].ToString());

                                strData = string.Format("FIN|{0}|{1}|{2}|0", intCtr, dblChkAmt, DateTime.Now.ToString("MMddyyyyHHmmss"));

                                oDTDetails.Rows.Add(intDTHRow, "0", strData);

                                strFileName = string.Format(@"{0}{1}_{2}_{3}_{4}.txt", strFilePath, strCompany, strBankTemp, strAccount, DateTime.Now.ToString("yyyyMMddHHmm"));

                                if (File.Exists(strFileName))
                                    File.Delete(strFileName);

                                FileInfo csvGenFile = new FileInfo(strFileName);
                                if (!(File.Exists(strFileName)))
                                {
                                    StreamWriter swGenFile = csvGenFile.CreateText();
                                    swGenFile.Close();
                                }

                                oDRDetails = oDTDetails.Select("Row = '" + intDTHRow.ToString() + "' ", "Type, Type ASC");
                                if (oDRDetails.Length > 0)
                                {
                                    for (int intRowD = 0; intRowD <= oDRDetails.Length - 1; intRowD++)
                                    {
                                        StreamWriter swGenFile = csvGenFile.AppendText();
                                        swGenFile.WriteLine(oDRDetails[intRowD]["Data"].ToString());
                                        swGenFile.Close();
                                    }
                                }
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                GlobalFunction.fileappend(ex.Message.ToString());
                return false;
            }
        }
        private void uf_initDataTable()
        {
            oDTBank = new System.Data.DataTable("BankFile");
            oDTBank.Columns.Add("Bank", typeof(System.String));
            oDTBank.Columns.Add("Account", typeof(System.String));
            oDTBank.Columns.Add("DocEntry", typeof(System.String));
            oDTBank.Clear();

            oDTHeader = new System.Data.DataTable("Bank File Header");
            oDTHeader.Columns.Add("Row", typeof(System.Int32));
            oDTHeader.Columns.Add("Bank", typeof(System.String));
            oDTHeader.Columns.Add("Account", typeof(System.String));
            oDTHeader.Columns.Add("Counter", typeof(System.Int32));
            oDTHeader.Columns.Add("Amount", typeof(System.Double));
            oDTHeader.Clear();

            oDTDetails = new System.Data.DataTable("Bank File Details");
            oDTDetails.Columns.Add("Row", typeof(System.Int32));
            oDTDetails.Columns.Add("Type", typeof(System.String));
            oDTDetails.Columns.Add("Data", typeof(System.String));
            oDTDetails.Clear();

        }
    }
}
