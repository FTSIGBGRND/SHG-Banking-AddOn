using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;

namespace AddOn.Outgoing_Check_Wrting
{
    public partial class userform_Outgoing_CheckClearing : AddOn.Form
    {
        private static string strDBType;
        public userform_Outgoing_CheckClearing()
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
        public override void itempressed(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);

            string strClrDte;

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

                    case "btnClr":

                        oForm.Freeze(true);

                        strClrDte =getItemString("ClrDate");

                        if (!(string.IsNullOrEmpty(strClrDte)))
                            uf_ClearedCheck(Convert.ToDateTime(strClrDte));
                        else
                        {
                            UI.SBO_Application.StatusBar.SetText("Clearing Date is missing!", SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            oForm.Freeze(false);
                            return;
                        }

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

            is_FormType = "100000005";
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
            oForm.DataSources.UserDataSources.Add("ClrDate", SAPbouiCOM.BoDataType.dt_DATE, 0);
            oForm.DataSources.UserDataSources.Add("Remarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 250);

            oForm.Title = "Check Clearing";
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

            oItem = createEditText(150, 60, 150, 14, "ClrDate", true, "", "ClrDate");
            oItem = createStaticText(6, 60, 100, 14, "stClrDate", "Cleared Date", "ClrDate");

            oItem = createButton(150, 80, 150, 19, "btnList", "&List");

            oItem = oForm.Items.Add("grd1", SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Enabled = true;
            oItem.Left = 6;
            oItem.Top = 120;
            oItem.Width = 1180;
            oItem.Height = 360;

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("CHKCLR");

            uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

            oItem = createButton(6, 505, 100, 19, "btnClr", "&Cleared");
            oItem = createButton(110, 505, 100, 19, "2", "");

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
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKCLEARING\" (to_date('{0}', 'MM/DD/YYYY'), to_date('{1}', 'MM/DD/YYYY')) ", dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKCLEARING '{0}', '{1}' ", dteFDoc, dteTDoc);
            }
            else
            {
                if (strDBType == "dst_HANADB")
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKCLEARING_BP\" ('{0}', to_date('{1}', 'MM/DD/YYYY'), to_date('{2}', 'MM/DD/YYYY')) ", strCardCode, dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKCLEARING_BP '{0}', '{1}', '{2}' ", strCardCode, dteFDoc, dteTDoc);
            }

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            oForm.DataSources.DataTables.Item("CHKCLR").ExecuteQuery(strQuery);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("CHKCLR");

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
        private void uf_ClearedCheck(DateTime dteClr)
        {
            SAPbouiCOM.Grid oGrid;

            SAPbobsCOM.Recordset oRecordset;

            SAPbobsCOM.JournalEntries oJournalEntries;

            string strDocEntry, strDocNum, strQuery = "";

            string strErrMsg, strPostDocEnt;
            int intErrCode;
            string strFDate, strTDate;

            DateTime dteFDoc, dteTDoc;

            oForm.Freeze(true);

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            if (oGrid.Rows.Count == 0)
            {
                UI.SBO_Application.StatusBar.SetText("Check for Clearing is missing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return;
            }
            else
            {
                for (int ctr = 0; ctr < oGrid.Rows.Count; ctr++)
                {
                    if (oGrid.DataTable.Columns.Item("Selected").Cells.Item(ctr).Value.ToString() == "Y")
                    {
                        strDocEntry = oGrid.DataTable.Columns.Item("DocEntry").Cells.Item(ctr).Value.ToString();
                        strDocNum = oGrid.DataTable.Columns.Item("DocNum").Cells.Item(ctr).Value.ToString();

                        strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", OOCW.\"U_DocDate\", OOCW.\"U_DueDate\", OOCW.\"U_TaxDate\", " +
                                                 "       OOCW.\"U_CardCode\", OOCW.\"U_CardName\", OOCW.\"U_PayName\", OOCW.\"U_Comments\", " +
                                                 "       OOCW.\"U_Bank\", OOCW.\"U_Branch\", OOCW.\"U_Account\", ACT1.\"AcctCode\" AS \"BAcctCode\", OOCW.\"U_BGLAcctCode\", " +
                                                 "       OOCW.\"U_BGLAcctName\",  ACT2.\"AcctCode\" AS \"RAcctCode\", OOCW.\"U_RGLAcctCode\", OOCW.\"U_RGLAcctName\", " +
                                                 "       OOCW.\"U_CheckNo\", OOCW.\"U_TotalDue\" " +
                                                 "FROM \"@FTOOCW\" OOCW LEFT JOIN OACT ACT1 ON OOCW.\"U_BGLAcctCode\" = ACT1.\"FormatCode\" " +
                                                 "                      LEFT JOIN OACT ACT2 ON OOCW.\"U_RGLAcctCode\" = ACT2.\"FormatCode\" " +
                                                 "WHERE OOCW.\"DocEntry\" = '{0}' ", strDocEntry);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            if (!DI.oCompany.InTransaction)
                                DI.oCompany.StartTransaction();

                            oJournalEntries = null;
                            oJournalEntries = (SAPbobsCOM.JournalEntries)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                            oJournalEntries.ReferenceDate = dteClr;
                            oJournalEntries.DueDate = dteClr;
                            oJournalEntries.TaxDate = dteClr;

                            oJournalEntries.Memo = string.Format("Check Writing # {0} ReClass", strDocNum);

                            oJournalEntries.Lines.AccountCode = oRecordset.Fields.Item("RAcctCode").Value.ToString();
                            oJournalEntries.Lines.Debit = Convert.ToDouble(oRecordset.Fields.Item("U_TotalDue").Value.ToString());
                            oJournalEntries.Lines.Credit = 0;

                            oJournalEntries.Lines.Add();
                            oJournalEntries.Lines.AccountCode = oRecordset.Fields.Item("BAcctCode").Value.ToString();
                            oJournalEntries.Lines.Credit = Convert.ToDouble(oRecordset.Fields.Item("U_TotalDue").Value.ToString());
                            oJournalEntries.Lines.Debit = 0;

                            if (oJournalEntries.Add() != 0)
                            {
                                if (DI.oCompany.InTransaction)
                                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                intErrCode = DI.oCompany.GetLastErrorCode();
                                strErrMsg = DI.oCompany.GetLastErrorDescription();

                                UI.SBO_Application.StatusBar.SetText(string.Format("Error Posting Journal Entry ReClass for Check Writing # {0}. {1} - {2}.", strDocNum, intErrCode, strErrMsg), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                strPostDocEnt = DI.oCompany.GetNewObjectKey();

                                if (strDBType == "dst_HANADB")
                                    strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'C', \"U_ClrDate\" = to_date('{0}', 'MM/DD/YYYY'), \"U_TransId\" = '{1}' " +
                                                             "WHERE \"DocEntry\" = '{2}' ", dteClr.ToString("MM/dd/yyyy"), strPostDocEnt, strDocEntry);
                                else
                                    strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'C', \"U_ClrDate\" = '{0}', \"U_TransId\" = '{1}' " +
                                                             "WHERE \"DocEntry\" = '{2}' ", dteClr, strPostDocEnt, strDocEntry);

                                if (!DI.executeQuery(strQuery))
                                {
                                    UI.SBO_Application.StatusBar.SetText(string.Format("Error Updating Base Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                else
                                {
                                    if (DI.oCompany.InTransaction == true)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                    strFDate = getItemString("FDocDate");
                                    if (string.IsNullOrEmpty(strFDate))
                                        strFDate = "01/01/1900";

                                    strTDate = getItemString("TDocDate");
                                    if (string.IsNullOrEmpty(strTDate))
                                        strTDate = "12/31/2099";

                                    dteFDoc = Convert.ToDateTime(strFDate);
                                    dteTDoc = Convert.ToDateTime(strTDate);

                                    uf_LoadCheck(getItemString("CardCode"), dteFDoc, dteTDoc);
                                }
                            }
                        }
                    }
                }
            }

            oForm.Freeze(false);

            GC.Collect();
        }

    }
}
