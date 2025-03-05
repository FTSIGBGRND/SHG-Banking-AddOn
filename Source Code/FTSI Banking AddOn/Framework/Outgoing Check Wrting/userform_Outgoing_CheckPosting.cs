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
    public partial class userform_Outgoing_CheckPosting : AddOn.Form
    {
        bool blSelect = false;

        private static string strDBType;
        public userform_Outgoing_CheckPosting()
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

                        break; ;
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

                    case "btnPost":

                        oForm.Freeze(true);

                        uf_PostCheck();

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

            is_FormType = "100000003";
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

            oForm.Title = "Check Posting";
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
            oItem.Height = 380;

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("CHKPOST");

            uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

            oItem = createButton(6, 505, 100, 19, "btnPost", "&Post");
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
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKPOSTING\" (to_date('{0}', 'MM/DD/YYYY'), to_date('{1}', 'MM/DD/YYYY')) ", dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKPOSTING '{0}', '{1}' ", dteFDoc, dteTDoc);
            }
            else
            {
                if (strDBType == "dst_HANADB")
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKPOSTING_BP\" ({0}, to_date('{1}', 'MM/DD/YYYY'), to_date('{2}', 'MM/DD/YYYY')) ", strCardCode, dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKPOSTING_BP '{0}', '{1}', '{2}' ", strCardCode, dteFDoc, dteTDoc);
            }

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            oForm.DataSources.DataTables.Item("CHKPOST").ExecuteQuery(strQuery);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("CHKPOST");

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
        private void uf_PostCheck()
        {
            SAPbouiCOM.Grid oGrid;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRSInvoice;
            SAPbobsCOM.Recordset oRSPayment;

            SAPbobsCOM.Payments oPayments;

            string strDocEntry, strDocNum, strBaseType, strQuery = "";
            string strCheckNo;
            string strFDate, strTDate;
            
            string strErrMsg, strPostDocEnt, strPostDocNum;
            int intErrCode;

            DateTime dteFDoc, dteTDoc;

            oForm.Freeze(true);

            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            if (oGrid.Rows.Count == 0)
            {
                UI.SBO_Application.StatusBar.SetText("Check for Posting is missing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                                                 "       DSC1.\"Country\", OOCW.\"U_Bank\", OOCW.\"U_Branch\", OOCW.\"U_Account\", OACT.\"AcctCode\", " +
                                                 "       OOCW.\"U_RGLAcctCode\", OOCW.\"U_RGLAcctName\", OOCW.\"U_CheckNo\", OOCW.\"U_TotalDue\" " +
                                                 "FROM \"@FTOOCW\" OOCW INNER JOIN OACT ON OOCW.\"U_RGLAcctCode\" = OACT.\"FormatCode\" " +
                                                 "                      INNER JOIN DSC1 ON OOCW.\"U_Bank\" = DSC1.\"BankCode\" " +
                                                 "WHERE OOCW.\"DocEntry\" = '{0}' ", strDocEntry);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            if (!DI.oCompany.InTransaction)
                                DI.oCompany.StartTransaction();

                            oPayments = null;
                            oPayments = (SAPbobsCOM.Payments)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

                            oPayments.DocDate = Convert.ToDateTime(oRecordset.Fields.Item("U_DocDate").Value.ToString());
                            oPayments.DueDate = Convert.ToDateTime(oRecordset.Fields.Item("U_DueDate").Value.ToString());
                            oPayments.TaxDate = Convert.ToDateTime(oRecordset.Fields.Item("U_TaxDate").Value.ToString());

                            oPayments.CardCode = oRecordset.Fields.Item("U_CardCode").Value.ToString();
                            oPayments.Address = oRecordset.Fields.Item("U_PayName").Value.ToString();
                            oPayments.Remarks = oRecordset.Fields.Item("U_Comments").Value.ToString();


                            strQuery = string.Format("SELECT OCW1.\"U_BaseType\", OCW1.\"U_BaseEntry\", OCW1.\"U_BaseNum\", OCW1.\"U_DocLine\", OCW1.\"U_InstId\", OCW1.\"U_TotPay\" " +
                                                     "FROM \"@FTOCW1\" OCW1 " +
                                                     "WHERE OCW1.\"DocEntry\" = '{0}' ", strDocEntry);

                            oRSInvoice = null;
                            oRSInvoice = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRSInvoice.DoQuery(strQuery);

                            if (oRSInvoice.RecordCount > 0)
                            {
                                oRSInvoice.MoveFirst();

                                for (int intInv = 0; intInv <= oRSInvoice.RecordCount - 1; intInv++)
                                {
                                    strBaseType = oRSInvoice.Fields.Item("U_BaseType").Value.ToString();

                                    if (intInv > 0)
                                        oPayments.Invoices.Add();

                                    oPayments.Invoices.DocEntry = Convert.ToInt32(oRSInvoice.Fields.Item("U_BaseEntry").Value.ToString());

                                    if (strBaseType == "18")
                                    {
                                        oPayments.Invoices.InstallmentId = Convert.ToInt32(oRSInvoice.Fields.Item("U_InstId").Value.ToString());
                                        oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_PurchaseInvoice;
                                    }
                                    else if (strBaseType == "19")
                                    {
                                        oPayments.Invoices.InstallmentId = Convert.ToInt32(oRSInvoice.Fields.Item("U_InstId").Value.ToString());
                                        oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_PurchaseCreditNote;
                                    }
                                    else if (strBaseType == "30")
                                    {
                                        oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_JournalEntry;
                                        oPayments.Invoices.DocLine = System.Convert.ToInt32(oRSInvoice.Fields.Item("U_DocLine").Value.ToString());

                                    }
                                    else if (strBaseType == "204")
                                        oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_PurchaseDownPayment;

                                    oPayments.Invoices.SumApplied = System.Convert.ToDouble(oRSInvoice.Fields.Item("U_TotPay").Value.ToString());

                                    oRSInvoice.MoveNext();
                                }
                            }

                            strCheckNo = oRecordset.Fields.Item("U_CheckNo").Value.ToString();

                            oPayments.Checks.UserFields.Fields.Item("U_CheckNo").Value = strCheckNo;

                            oPayments.Checks.CountryCode = oRecordset.Fields.Item("Country").Value.ToString();
                            oPayments.Checks.BankCode = oRecordset.Fields.Item("U_Bank").Value.ToString();
                            oPayments.Checks.Branch = oRecordset.Fields.Item("U_Branch").Value.ToString();
                            oPayments.Checks.AccounttNum = oRecordset.Fields.Item("U_Account").Value.ToString();
                            oPayments.Checks.CheckAccount = oRecordset.Fields.Item("AcctCode").Value.ToString();

                            if (!(string.IsNullOrEmpty(strCheckNo)) && strCheckNo != "0")
                            {
                                if (strCheckNo.Length > 6)
                                    strCheckNo = strCheckNo.Substring(strCheckNo.Length - 6);

                                oPayments.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES;
                                oPayments.Checks.CheckNumber = Convert.ToInt32(strCheckNo);
                            }

                            oPayments.Checks.CheckSum = Convert.ToDouble(oRecordset.Fields.Item("U_TotalDue").Value.ToString());
                            oPayments.Checks.DueDate = Convert.ToDateTime(oRecordset.Fields.Item("U_DueDate").Value.ToString());

                            if (oPayments.Add() != 0)
                            {
                                if (DI.oCompany.InTransaction)
                                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                intErrCode =  DI.oCompany.GetLastErrorCode();
                                strErrMsg = DI.oCompany.GetLastErrorDescription();

                                UI.SBO_Application.StatusBar.SetText(string.Format("Error Posting Outgoing Payment for Check Writing # {0}. {1} - {2}.", strDocNum, intErrCode, strErrMsg), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {

                                strPostDocEnt = DI.oCompany.GetNewObjectKey();

                                oRSPayment = null;
                                oRSPayment = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oRSPayment.DoQuery("SELECT \"DocNum\" FROM OVPM WHERE \"DocEntry\" = '" + strPostDocEnt + "' ");

                                strPostDocNum = oRSPayment.Fields.Item("DocNum").Value.ToString();

                                strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'P', \"U_OVPMDocEnt\" = '{0}', \"U_OVPMDocNum\" = '{1}' " +
                                                         "WHERE \"DocEntry\" = '{2}' ", strPostDocEnt, strPostDocNum, strDocEntry);
                                if (!DI.executeQuery(strQuery))
                                {
                                    UI.SBO_Application.StatusBar.SetText(string.Format("Error Updating Base Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                else
                                {
                                    if (DI.oCompany.InTransaction)
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
