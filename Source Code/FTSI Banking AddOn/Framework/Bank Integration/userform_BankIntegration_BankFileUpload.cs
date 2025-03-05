using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;
using System.IO;

namespace AddOn.Bank_Integration
{
    public partial class userform_BankIntegration_BankFileUpload : AddOn.Form
    {
        public userform_BankIntegration_BankFileUpload()
        {
            InitializeComponent();
        }
        public override void itempressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);
            string strFilePath;


            if (pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "btnBrowse":

                        GlobalFunction.showOpenFileDialog(oForm.Items.Item("FilePath"), "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|Excel Files (*.xlsx)|*.xlsx");
                        GC.Collect();

                        break;

                    case "btnUpload":

                        oForm.Freeze(true);

                        strFilePath = getItemString("FilePath");

                        if (string.IsNullOrEmpty(strFilePath))
                        {
                            UI.SBO_Application.StatusBar.SetText("Please select file to be uploaded.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            oForm.Freeze(false);
                            return;
                        }
                        else
                        {
                            UI.displayStatus();
                            UI.changeStatus("Processing Bank Return File Upload. Please Wait...");

                            uf_uploadRetFile(strFilePath);

                            UI.hideStatus();
                        }

                        oForm.Freeze(false);

                        break;
                }
            }
        }
        public override void onGetCreationParams(ref BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
            base.onGetCreationParams(ref io_BorderStyle, ref is_FormType, ref is_ObjectType, ref xmlPath);

            is_FormType = "100000007";
            io_BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

            GC.Collect();

        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);

            SAPbouiCOM.Item oItem;

            oForm.Freeze(true);

            oForm.DataSources.UserDataSources.Add("FilePath", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 10000);

            oForm.Title = "Bank File Upload";
            oForm.Width = 550;
            oForm.Height = 200;

            oItem = createEditText(107, 45, 388, 14, "FilePath", true, "", "FilePath");
            oItem.Enabled = true;
            oItem = createStaticText(6, 45, 100, 14, "stFilePath", "File Path", "FilePath");

            oItem = createButton(500, 45, 20, 14, "btnBrowse", "...");

            oItem = createButton(6, 130, 100, 19, "btnUpload", "&Upload");
            oItem = createButton(107, 130, 100, 19, "2", "");

            oForm.Freeze(false);

            GC.Collect();

        }
        private void uf_uploadRetFile(string strFilePath)
        {
            string strFileName, strCompany, strBankTemp, strType;
            string[] strAFileName;

            SAPbobsCOM.Recordset oRecordset, oRSRet;

            oForm.Freeze(true);

            try
            {
               
                strFileName = Path.GetFileNameWithoutExtension(strFilePath);
                strAFileName = strFileName.Split(Convert.ToChar("_"));

                if (strAFileName.Length != 4)
                {
                    UI.SBO_Application.StatusBar.SetText("Invalid File Selected. Please Validate Filename.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return;
                }
                else
                {
                    strCompany = strAFileName[1];

                    oRecordset = null;
                    oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery("SELECT \"AliasName\" FROM OADM");

                    if (strCompany == oRecordset.Fields.Item("AliasName").Value.ToString())
                    {
                        strBankTemp = strAFileName[2];

                        oRSRet = null;
                        oRSRet = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRSRet.DoQuery(string.Format("SELECT \"U_RetFleNme\" FROM \"@FTOOCW\" WHERE \"U_RetFleNme\" = '{0}' ", strFileName));

                        if (oRSRet.RecordCount > 0)
                        {
                            UI.SBO_Application.StatusBar.SetText("Selected File already uploaded.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        else
                        {
                            if (strBankTemp == "UB")
                                uf_uploadRetFileUB(strFilePath);

                        }
                    }
                    else
                    {
                        UI.SBO_Application.StatusBar.SetText("Invalid File Selected. Please Validate Company Setup  .", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalFunction.fileappend("Error uploading Return File - " + ex.Message.ToString());

                GC.Collect();
                oForm.Freeze(false);
            }

            GC.Collect();
            oForm.Freeze(false);

        }
        private void uf_uploadRetFileUB(string strFilePath)
        {
            string strStatus, strDocNum, strCheckNo, strRlsBrnch, strClrDate;

            bool blWithErr = false;

            try
            {
                if (GlobalFunction.importXLSX(strFilePath, "YES", "Sheet1"))
                {
                    if (globalvar.oDTImpData.Rows.Count > 0)
                    {                        
                        if (!DI.oCompany.InTransaction)
                            DI.oCompany.StartTransaction();

                        for (int intIRow = 0; intIRow <= globalvar.oDTImpData.Rows.Count - 1; intIRow++)
                        {
                            strDocNum = globalvar.oDTImpData.Rows[intIRow][0].ToString();  
                            strCheckNo = globalvar.oDTImpData.Rows[intIRow][4].ToString();
                            strRlsBrnch = globalvar.oDTImpData.Rows[intIRow][8].ToString();
                            strClrDate = globalvar.oDTImpData.Rows[intIRow][9].ToString();
                            strStatus = globalvar.oDTImpData.Rows[intIRow][10].ToString();

                            if (strStatus == "For Delivery")
                            {
                                if (!(uf_postOutgoingPayment(strDocNum, strCheckNo)))
                                {
                                    blWithErr = true;

                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                    break;
                                }
                            }
                            else if (strStatus == "Claimed")
                            {
                                if (!(uf_postReleasedCheck(strDocNum)))
                                {
                                    blWithErr = true;

                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                    break;
                                }
                            }
                            else if (strStatus == "Negotiated")
                            {

                                if (!(uf_postJournalEntry(strDocNum, Convert.ToDateTime(strClrDate))))
                                {
                                    blWithErr = true;

                                    if (DI.oCompany.InTransaction)
                                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                    break;
                                }
                            }
                        }

                        if (blWithErr == false)
                        {
                            if (DI.oCompany.InTransaction)
                                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                            UI.SBO_Application.StatusBar.SetText(string.Format("Bank Return File Successfully uploaded."), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                            setItemString("FilePath", "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GlobalFunction.fileappend("Error uploading Return File - " + ex.Message.ToString());

                GC.Collect();
                oForm.Freeze(false);
            }
        }
        private bool uf_postOutgoingPayment(string strDocNum, string strCheckNo)
        {

            string strQuery, strDocEntry, strBaseType, strOCheck = "";

            string strErrMsg, strPostDocEnt, strPostDocNum;
            int intErrCode;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRSInvoice;
            SAPbobsCOM.Recordset oRSPayment;

            SAPbobsCOM.Payments oPayments;


            strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", OOCW.\"U_DocDate\", OOCW.\"U_DueDate\", OOCW.\"U_TaxDate\", " +
                         "       OOCW.\"U_CardCode\", OOCW.\"U_CardName\", OOCW.\"U_PayName\", OOCW.\"U_Comments\", " +
                         "       DSC1.\"Country\", OOCW.\"U_Bank\", OOCW.\"U_Branch\", OOCW.\"U_Account\", OACT.\"AcctCode\", " +
                         "       OOCW.\"U_RGLAcctCode\", OOCW.\"U_RGLAcctName\", OOCW.\"U_CheckNo\", OOCW.\"U_TotalDue\" " +
                         "FROM \"@FTOOCW\" OOCW INNER JOIN OACT ON OOCW.\"U_RGLAcctCode\" = OACT.\"FormatCode\" " +
                         "                      INNER JOIN DSC1 ON OOCW.\"U_Bank\" = DSC1.\"BankCode\" " +
                         "WHERE OOCW.\"DocNum\" = '{0}' AND OOCW.\"U_Status\" =  'A' ", strDocNum);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
            {
                strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();

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

                oPayments.Checks.UserFields.Fields.Item("U_CheckNo").Value = strCheckNo;

                oPayments.Checks.CountryCode = oRecordset.Fields.Item("Country").Value.ToString();
                oPayments.Checks.BankCode = oRecordset.Fields.Item("U_Bank").Value.ToString();
                oPayments.Checks.Branch = oRecordset.Fields.Item("U_Branch").Value.ToString();
                oPayments.Checks.AccounttNum = oRecordset.Fields.Item("U_Account").Value.ToString();
                oPayments.Checks.CheckAccount = oRecordset.Fields.Item("AcctCode").Value.ToString();

                if (!(string.IsNullOrEmpty(strCheckNo)) && strCheckNo != "0")
                {
                    strOCheck = strCheckNo;

                    if (strCheckNo.Length > 6)
                        strCheckNo = strCheckNo.Substring(strCheckNo.Length - 4, 4);

                    oPayments.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES;
                    oPayments.Checks.CheckNumber = Convert.ToInt32(strCheckNo);
                }

                oPayments.Checks.CheckSum = Convert.ToDouble(oRecordset.Fields.Item("U_TotalDue").Value.ToString());
                oPayments.Checks.DueDate = Convert.ToDateTime(oRecordset.Fields.Item("U_DueDate").Value.ToString());

                if (oPayments.Add() != 0)
                {
                    intErrCode = DI.oCompany.GetLastErrorCode();
                    strErrMsg = DI.oCompany.GetLastErrorDescription();

                    UI.SBO_Application.StatusBar.SetText(string.Format("Error Posting Outgoing Payment for Check Writing # {0}. {1} - {2}.", strDocNum, intErrCode, strErrMsg), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                    return false;
                }
                else
                {
                    strPostDocEnt = DI.oCompany.GetNewObjectKey();

                    oRSPayment = null;
                    oRSPayment = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRSPayment.DoQuery("SELECT \"DocNum\" FROM OVPM WHERE \"DocEntry\" = '" + strPostDocEnt + "' ");

                    strPostDocNum = oRSPayment.Fields.Item("DocNum").Value.ToString();

                    strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_CheckNo\" = '{0}', \"U_Status\" = 'P', \"U_OVPMDocEnt\" = '{1}', \"U_OVPMDocNum\" = '{2}' " +
                                             "WHERE \"DocEntry\" = '{3}' ", strOCheck, strPostDocEnt, strPostDocNum, strDocEntry);

                    if (!DI.executeQuery(strQuery))
                    {
                        UI.SBO_Application.StatusBar.SetText(string.Format("Error Updating Base Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        return false;
                    }

                    return true;
                }
            }
            else
            {
                UI.SBO_Application.StatusBar.SetText(string.Format("Check Writing # {0} not found or already posted.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
        }
        private bool uf_postJournalEntry(string strDocNum, DateTime dteClr)
        {
            string strQuery, strDocEntry;
            string strErrMsg, strPostDocEnt;

            string strDBType = DI.oCompany.DbServerType.ToString();

            int intErrCode;

            SAPbobsCOM.Recordset oRecordset;

            SAPbobsCOM.JournalEntries oJournalEntries;

            strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", OOCW.\"U_DocDate\", OOCW.\"U_DueDate\", OOCW.\"U_TaxDate\", " +
                                     "       OOCW.\"U_CardCode\", OOCW.\"U_CardName\", OOCW.\"U_PayName\", OOCW.\"U_Comments\", " +
                                     "       OOCW.\"U_Bank\", OOCW.\"U_Branch\", OOCW.\"U_Account\", ACT1.\"AcctCode\" AS \"BAcctCode\", OOCW.\"U_BGLAcctCode\", " +
                                     "       OOCW.\"U_BGLAcctName\",  ACT2.\"AcctCode\" AS \"RAcctCode\", OOCW.\"U_RGLAcctCode\", OOCW.\"U_RGLAcctName\", " +
                                     "       OOCW.\"U_CheckNo\", OOCW.\"U_TotalDue\" " +
                                     "FROM \"@FTOOCW\" OOCW LEFT JOIN OACT ACT1 ON OOCW.\"U_BGLAcctCode\" = ACT1.\"FormatCode\" " +
                                     "                      LEFT JOIN OACT ACT2 ON OOCW.\"U_RGLAcctCode\" = ACT2.\"FormatCode\" " +
                                     "WHERE OOCW.\"DocNum\" = '{0}' AND OOCW.\"U_Status\" =  'R' ", strDocNum);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
            {
                strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();

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
                    intErrCode = DI.oCompany.GetLastErrorCode();
                    strErrMsg = DI.oCompany.GetLastErrorDescription();

                    UI.SBO_Application.StatusBar.SetText(string.Format("Error Posting Journal Entry ReClass for Check Writing # {0}. {1} - {2}.", strDocNum, intErrCode, strErrMsg), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                    return false;
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

                        return false;
                    }
                }
            }
            else
            {
                UI.SBO_Application.StatusBar.SetText(string.Format("Check Writing # {0} not found or already cleared.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                return false;
            }

            return true;
        }
        private bool uf_postReleasedCheck(string strDocNum)
        {
            string strQuery, strDocEntry, strRType;

            string strDBType = DI.oCompany.DbServerType.ToString();

            SAPbobsCOM.Recordset oRecordset;

            strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", BARB.\"U_Type\" " +
                                     "FROM \"@FTOOCW\" OOCW LEFT JOIN \"@FTBARB\" BARB ON OOCW.\"U_RBranch\" = BARB.\"Code\" " +
                                     "WHERE OOCW.\"DocEntry\" = '{0}' AND OOCW.\"U_Status\" =  'P' ", strDocNum);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
            {
                strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();
                strRType = oRecordset.Fields.Item("U_Type").Value.ToString();

                if (strDBType == "dst_HANADB")
                {
                    if (strRType != "RET")
                        strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'R', \"U_RelDate\" = to_date('{0}', 'MM/DD/YYYY') WHERE \"DocEntry\" = '{1}' ", DateTime.Today.ToString("MM/dd/yyyy"), strDocEntry);
                    else
                        strQuery = "";
                }
                else
                {
                    if (strRType != "RET")
                        strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'R', \"U_RelDate\" = '{0}' WHERE \"DocEntry\" = '{1}' ", DateTime.Today, strDocEntry);
                    else
                        strQuery = "";
                }

                if (!string.IsNullOrEmpty(strQuery))
                {
                    if (!DI.executeQuery(strQuery))
                    {
                        UI.SBO_Application.StatusBar.SetText(string.Format("Error Updating Base Check Writing # {0}.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                        return false;
                    }
                }               
            }
            else
            {
                UI.SBO_Application.StatusBar.SetText(string.Format("Check Writing # {0} not found or already released.", strDocNum), SAPbouiCOM.BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                return false;
            }

            return true;
        }
    }
}
