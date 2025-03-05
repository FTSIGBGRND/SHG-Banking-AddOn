using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Runtime.InteropServices.ComTypes;

namespace AddOn.Outgoing_Check_Wrting
{
    public partial class userform_Outgoing_EmailNotification : AddOn.Form
    {
        bool blSelect = false;
        private static string strDBType;
        public userform_Outgoing_EmailNotification()
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

                    case "btnNot":

                        uf_SendEmailNotif();

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

            is_FormType = "100000008";
            io_BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

            GC.Collect();
        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);

            string strUserId;
            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.CheckBox oCheckBox;

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

            oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 250);
            oForm.DataSources.UserDataSources.Add("FDocDate", SAPbouiCOM.BoDataType.dt_DATE, 0);
            oForm.DataSources.UserDataSources.Add("TDocDate", SAPbouiCOM.BoDataType.dt_DATE, 0);
            oForm.DataSources.UserDataSources.Add("Remarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 250);

            oForm.Title = "Check Writing E-Mail Notification";
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

            strDBType = DI.oCompany.DbServerType.ToString();

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("CHKNTF");

            uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));

            oItem = createButton(6, 505, 180, 19, "btnNot", "&Send E-Mail Notification");
            oItem = createButton(200, 505, 100, 19, "2", "");

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
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_EMAILNOTIF\" (to_date('{0}', 'MM/DD/YYYY'), to_date('{1}', 'MM/DD/YYYY')) ", dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_EMAILNOTIF '{0}', '{1}' ", dteFDoc, dteTDoc);
            }
            else
            {
                if (strDBType == "dst_HANADB")
                    strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_EMAILNOTIF_BP\" ('{0}', to_date('{1}', 'MM/DD/YYYY'), to_date('{2}', 'MM/DD/YYYY')) ", strCardCode, dteFDoc.ToString("MM/dd/yyyy"), dteTDoc.ToString("MM/dd/yyyy"));
                else
                    strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_EMAILNOTIF_BP '{0}', '{1}', '{2}' ", strCardCode, dteFDoc, dteTDoc);
            }

           
            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            oForm.DataSources.DataTables.Item("CHKNTF").ExecuteQuery(strQuery);
            oGrid.DataTable = oForm.DataSources.DataTables.Item("CHKNTF");

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
        private void uf_SendEmailNotif()
        {

            string strQuery, strDocEntry, strDocNum, strInvNo;
            string strSubject, strMailTo, strMailCC, strMailBody;

            SAPbobsCOM.Recordset oRecordset, oRSInv;
            
            SAPbouiCOM.Grid oGrid;

            oForm.Freeze(true);


            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd1").Specific;
            if (oGrid.Rows.Count == 0)
            {
                UI.SBO_Application.StatusBar.SetText("Check Writing for E-Mail Notification is missing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return;
            }
            else
            {
                UI.displayStatus();
                UI.changeStatus("Sending E-Mail Notifications. Please Wait...");

                for (int ctr = 0; ctr < oGrid.Rows.Count; ctr++)
                {
                    if (oGrid.DataTable.Columns.Item("Selected").Cells.Item(ctr).Value.ToString() == "Y")
                    {
                        strMailTo = "";
                        strInvNo = "";
                        strSubject = "";
                        strMailTo = "";
                        strMailCC = "";

                        strDocEntry = oGrid.DataTable.Columns.Item("DocEntry").Cells.Item(ctr).Value.ToString();
                        strDocNum = oGrid.DataTable.Columns.Item("DocNum").Cells.Item(ctr).Value.ToString();

                        strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", " +
                                                 "       OOCW.\"U_CardName\", OCRD.\"E_Mail\", OOCW.\"U_DueDate\", ODSC.\"BankName\", OOCW.\"U_CheckNo\", " +
                                                 "       OADM.\"AliasName\", OADM.\"CompnyName\", " +
                                                 "       OOCW.\"U_TotalDue\", BARB.\"Name\", BARB.\"U_Address\" " + 
                                                 "FROM \"@FTOOCW\" OOCW LEFT JOIN ODSC ON OOCW.\"U_Bank\" = ODSC.\"BankCode\" " +
                                                 "                      LEFT JOIN OCRD ON OOCW.\"U_CardCode\" = OCRD.\"CardCode\" " +
                                                 "                      LEFT JOIN \"@FTBARB\" BARB ON OOCW.\"U_RBranch\" = BARB.\"Code\", OADM " +
                                                 "WHERE OOCW.\"DocEntry\" = '{0}' ", strDocEntry);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {

                            strQuery = string.Format("SELECT OCW1.\"U_BaseType\", OCW1.\"U_BaseEntry\", OCW1.\"U_BaseNum\", " +
                                                     "       CASE WHEN OCW1.\"U_BaseType\" = '18' THEN IFNULL(OPCH.\"NumAtCard\", '') " +
                                                     "            WHEN OCW1.\"U_BaseType\" = '19' THEN IFNULL(ORPC.\"NumAtCard\", '') " +
                                                     "            WHEN OCW1.\"U_BaseType\" = '204' THEN IFNULL(ORPC.\"NumAtCard\", '') " +
                                                     "            WHEN OCW1.\"U_BaseType\" = '30' THEN IFNULL(CAST(OJDT.\"TransId\" AS NVARCHAR(30)), '') END AS \"NumAtCard\" " +
                                                     "FROM \"@FTOCW1\" AS \"OCW1\" LEFT JOIN (SELECT OPCH.\"NumAtCard\", OPCH.\"DocEntry\" " +
                                                     "                                        FROM OPCH) AS OPCH ON OCW1.\"U_BaseEntry\" = OPCH.\"DocEntry\" AND OCW1.\"U_BaseType\" = '18' " +
                                                     "                             LEFT JOIN (SELECT ORPC.\"NumAtCard\", ORPC.\"DocEntry\" " +
                                                     "                                        FROM ORPC) AS ORPC ON OCW1.\"U_BaseEntry\" = ORPC.\"DocEntry\" AND OCW1.\"U_BaseType\" = '19' " +
                                                     "                             LEFT JOIN (SELECT ODPO.\"NumAtCard\", ODPO.\"DocEntry\" " +
                                                     "                                        FROM ODPO) AS ODPO ON OCW1.\"U_BaseEntry\" = ODPO.\"DocEntry\" AND OCW1.\"U_BaseType\" = '204' " +
                                                     "                             LEFT JOIN (SELECT OJDT.\"TransId\" " +
                                                     "                                        FROM OJDT) AS OJDT ON OCW1.\"U_BaseEntry\" = OJDT.\"TransId\" AND OCW1.\"U_BaseType\" = '30' " +
                                                     "WHERE OCW1.\"DocEntry\" = '{0}' ", strDocEntry);

                            oRSInv = null;
                            oRSInv = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRSInv.DoQuery(strQuery);

                            if (oRSInv.RecordCount > 0)
                            {
                                while (!(oRSInv.EoF))
                                {
                                    strMailTo = oRecordset.Fields.Item("E_Mail").Value.ToString();

                                    if (string.IsNullOrEmpty(strInvNo))
                                        strInvNo = oRSInv.Fields.Item("NumAtCard").Value.ToString();
                                    else
                                        strInvNo = strInvNo + "; " + oRSInv.Fields.Item("NumAtCard").Value.ToString();

                                    oRSInv.MoveNext();
                                }
                            }

                            strMailBody = uf_EmailBody(oRecordset.Fields.Item("CompnyName").Value.ToString(),
                                                       oRecordset.Fields.Item("U_CardName").Value.ToString(),
                                                       Convert.ToDateTime(oRecordset.Fields.Item("U_DueDate").Value.ToString()),
                                                       oRecordset.Fields.Item("BankName").Value.ToString(),
                                                       oRecordset.Fields.Item("U_CheckNo").Value.ToString(),
                                                       oRecordset.Fields.Item("AliasName").Value.ToString(),
                                                       oRecordset.Fields.Item("U_TotalDue").Value.ToString(),
                                                       strInvNo,
                                                       oRecordset.Fields.Item("Name").Value.ToString(),
                                                       oRecordset.Fields.Item("U_Address").Value.ToString());

                            if (!(string.IsNullOrEmpty(strMailTo)))
                            {
                                if (EmailSender.sendSMTPEmail(strSubject, strMailTo, strMailCC, strMailBody))
                                {
                                    strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_EmailNotif\" = 'Y' WHERE \"DocEntry\" = '{0}' ", strDocEntry);

                                    if (!DI.executeQuery(strQuery))
                                        UI.SBO_Application.StatusBar.SetText(string.Format("Error Updating Base Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    else
                                        UI.SBO_Application.StatusBar.SetText(string.Format("Successfully sent e-mail notification for Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                else
                                    UI.SBO_Application.StatusBar.SetText(string.Format("Failed sending e-mail notification for Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                UI.SBO_Application.StatusBar.SetText(string.Format("Failed sending e-mail notification for Check Writing # {0}. Business Partner E-Mail is missing.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                    }
                }

                UI.hideStatus();

                uf_LoadCheck("", Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/1900"));
            }

            oForm.Freeze(false);

            GC.Collect();

        }
        private static string uf_EmailBody(string strCompName, string strCardName, DateTime dteDue, string strBankName, string strCheckNo, string strAliasName, string strTotalDue, string strInvNo, 
                                           string strRBranch, string strAddress)
        {
            string strEmailBody = "";

            strEmailBody = string.Format("{0} \r\n\r\n " +
                                         "7/F RFM Corporate Center, Pioneer Corp. \r\n\r\n " +
                                         "Sheridan St. Brgy. Buayang Bato, Mandaluyong City \r\n\r\n " +
                                         "TIN: 008-007-902-000 \r\n\r\n\r\n " +
                                         "PAYMENT ADVICE NOTIFICATION \r\n\r\n\r\n " +
                                         "Date: {1} \r\n\r\n\r\n " +
                                         "Please be advised that you have a check for collection this coming Friday with the ff. Details below. \r\n\r\n " +
                                         "Supplier: {2} \r\n " +
                                         "Check Date: {3} \r\n " +
                                         "Bank: {4} \r\n " +
                                         "Check No: {5} \r\n " +
                                         "Brand: {6} \r\n " +
                                         "Net Amount: {7} \r\n " +
                                         "Invoice No./s: {8} \r\n\r\n " +
                                         "COLLECTION INSTRUCTION \r\n\r\n\r\n " +
                                         "Releasing Site: {9} \r\n\r\n " +
                                         "Releasing Address: {10} \r\n\r\n " +
                                         "Operating Hours: 1:00 PM to 4:00 PM Every Friday Only(If friday falls on a holiday, releasing of checks will be done on the next banking day. \r\n\r\n " +
                                         "Contact Person / s: Jordana Lumabang / Gezelle Matimtim \r\n\r\n " +
                                         "Contact No: 09215394963 / 09214685579 \r\n\r\n\r\n " +
                                         "COLLECTION AND RELEASING POLICY \r\n\r\n\r\n " +
                                         "For Supplier with Terms: \r\n\r\n " +
                                         "*******No Official Receipts / Collection Receipts, no receiving of check payment ******* \r\n\r\n " +
                                         "For APDP / COD:  \r\n\r\n " +
                                         "*****No Sales Invoice, No Official Receipts / Collection Receipts, No Receiving of check payment ****** \r\n\r\n\r\n\r\n " +
                                         "DISCLAIMER: This is a system generated message. If you have any questions and clarifications about the payment details. " +
                                         "Please send an email to treasury@shg.com.ph. ",
                                         strCompName, DateTime.Today.ToString("MM/dd/yyyy"), strCardName, dteDue.ToString("MM/dd/yyyy"), strBankName, strCheckNo, 
                                         strAliasName, strTotalDue, strInvNo, strRBranch, strAddress);

            return strEmailBody;
        }
    }
}
