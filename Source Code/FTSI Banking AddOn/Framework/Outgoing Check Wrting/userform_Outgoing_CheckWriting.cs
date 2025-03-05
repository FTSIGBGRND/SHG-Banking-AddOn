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
    public partial class userform_Outgoing_CheckWriting : AddOn.Form
    {
        
        private static double dblTotPay = 0;
        private static string is_docno;

        private static string strDBType;
        public userform_Outgoing_CheckWriting()
        {
            InitializeComponent();
        }
        public userform_Outgoing_CheckWriting(string as_docno)
        {
            InitializeComponent();
            is_docno = as_docno;

        }
        public override void comboselect(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            base.comboselect(FormUID, ref pVal, ref BubbleEvent);

            string strTransType, strBank, strAccount;

            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.ComboBox oCmbAccount;

            if (!pVal.BeforeAction && pVal.ItemChanged)
            {
                switch (pVal.ItemUID)
                {
                    case "Account":

                        oForm.Freeze(true);

                        strAccount = getItemSelectedValue(pVal.ItemUID, "").ToString();

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(string.Format("SELECT DSC1.\"Branch\", DSC1.\"GLAccount\", OACT.\"FormatCode\", OACT.\"AcctName\", " +
                                                         "       DSC1.\"U_RGLAcctCode\", DSC1.\"U_RGLAcctName\"  " +
                                                         "FROM DSC1 INNER JOIN OACT ON DSC1.\"GLAccount\" = OACT.\"AcctCode\" " +
                                                         "WHERE \"Account\" = '{0}' ", strAccount));

                        if (oRecordset.RecordCount > 0)
                        {
                            setItemString("Branch", oRecordset.Fields.Item("Branch").Value.ToString());
                            setItemString("BGLAcctCod", oRecordset.Fields.Item("FormatCode").Value.ToString());
                            setItemString("BGLAcctNam", oRecordset.Fields.Item("AcctName").Value.ToString());
                            setItemString("RGLAcctCod", oRecordset.Fields.Item("U_RGLAcctCode").Value.ToString());
                            setItemString("RGLAcctNam", oRecordset.Fields.Item("U_RGLAcctName").Value.ToString());
                        }
                        else
                        {
                            setItemString("Branch", "");
                            setItemString("BGLAcctCod", "");
                            setItemString("BGLAcctNam", "");
                            setItemString("RGLAcctCod", "");
                            setItemString("RGLAcctNam", "");
                        }

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "Bank":

                        oForm.Freeze(true);

                        strBank = getItemSelectedValue(pVal.ItemUID, "").ToString();

                        oCmbAccount = (SAPbouiCOM.ComboBox)oForm.Items.Item("Account").Specific;
                        oCmbAccount.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                        if (oCmbAccount.ValidValues.Count > 1)
                        {
                            while (oCmbAccount.ValidValues.Count > 1)
                            {
                                oCmbAccount.ValidValues.Remove(1, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(string.Format("SELECT \"Account\" FROM DSC1 WHERE \"BankCode\" = '{0}' ", strBank));

                        if (oRecordset.RecordCount > 0)
                        {
                            oRecordset.MoveFirst();

                            while (!(oRecordset.EoF))
                            {
                                oCmbAccount.ValidValues.Add(oRecordset.Fields.Item("Account").Value.ToString(), oRecordset.Fields.Item("Account").Value.ToString());
                                oRecordset.MoveNext();
                            }
                        }

                        setItemSelectedValue("Account", "");
                        setItemString("Branch", "");
                        setItemString("BGLAcctCod", "");
                        setItemString("BGLAcctNam", "");
                        setItemString("RGLAcctCod", "");
                        setItemString("RGLAcctNam", "");

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "TransType":

                        oForm.Freeze(true);

                        oEditText = (SAPbouiCOM.EditText)getItemSpecific("CardCode");

                        strTransType = getItemSelectedValue("TransType", "");
                        if (strTransType == "S")
                        {
                            oEditText.ChooseFromListUID = "cflBPS";
                            oEditText.ChooseFromListAlias = "CardCode";

                            setItemEnabled("PayName", true);
                            setItemEnabled("CardCode", true);

                        }
                        else if (strTransType == "C")
                        {
                            oEditText.ChooseFromListUID = "cflBPC";
                            oEditText.ChooseFromListAlias = "CardCode";

                            setItemEnabled("PayName", true);
                            setItemEnabled("CardCode", true);
                        }
                        else
                        {
                            setItemEnabled("PayName", false);
                            setItemEnabled("CardCode", false);
                        }

                        oForm.Update();
                        oForm.Freeze(false);

                        break;
                }
            }

            GC.Collect();
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

                            if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                strCardCode = oDatatable.GetValue("CardCode", 0).ToString();
                                strCardName = oDatatable.GetValue("CardName", 0).ToString();

                                oForm.DataSources.DBDataSources.Item("@FTOOCW").SetValue("U_CardCode", 0, strCardCode);
                                oForm.DataSources.DBDataSources.Item("@FTOOCW").SetValue("U_CardName", 0, strCardName);
                                oForm.DataSources.DBDataSources.Item("@FTOOCW").SetValue("U_PayName", 0, strCardName);

                                uf_ListInvoices(strCardCode);
                            }

                            GC.Collect();
                            oForm.Freeze(false);

                            break;
                    }
                }
            }

            GC.Collect();
        }
        public override void itempressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.CheckBox oCheckBox;

            if (!pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "btnEmail":

                        oForm.Freeze(true);

                        uf_SendEmailNotif(getItemString("DocNum"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "btnCan":

                        oForm.Freeze(true);

                        uf_Cancel(getItemString("DocNum"), getItemString("OVPMDocEnt"));

                        oForm.Update();
                        oForm.Freeze(false);

                        break;

                    case "grd1":

                        switch (pVal.ColUID)
                        {
                            case "Select":

                                oForm.Freeze(true);
                                
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;
                                oCheckBox = (SAPbouiCOM.CheckBox)getColumnSpecific("grd1", pVal.ColUID, pVal.Row);

                                if (pVal.Row > 0 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    uf_Compute();
                                

                                oForm.Update();
                                oForm.Freeze(false);

                                break;

                        }

                        break;
                }
            }
        }
        public override void keydown(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.keydown(FormUID, ref pVal, ref BubbleEvent);

            if (!pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "DocNum":
                    case "CardCode":
                    case "CardName":
                    case "DocDate":
                    case "DueDate":
                    case "TaxDate":
                    case "CheckNo":
                    case "Status":

                        if (oForm.Mode == BoFormMode.fm_FIND_MODE)
                        {
                            if (pVal.CharPressed == 13)
                                itemclick("1");
                        }
                             
                        break;
                }
            }

        }
        public override void onGetCreationParams(ref SAPbouiCOM.BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
            base.onGetCreationParams(ref io_BorderStyle, ref is_FormType, ref is_ObjectType, ref xmlPath);

            is_ObjectType = "FTOOCW";
            is_FormType = "100000001";
            io_BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);

            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.ComboBox oComboBox;
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.Column oColumn;

            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            SAPbouiCOM.Condition oCon;
            SAPbouiCOM.Conditions oCons;

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)UI.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.UniqueID = "cflBPS";
            oCFLCreationParams.ObjectType = "2";
            oCFL = oCFLs.Add(oCFLCreationParams);
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "S";
            oCFL.SetConditions(oCons);

            oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)UI.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.UniqueID = "cflBPC";
            oCFLCreationParams.ObjectType = "2";
            oCFL = oCFLs.Add(oCFLCreationParams);
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "C";
            oCFL.SetConditions(oCons);

            oForm.Freeze(true);

            oForm.Title = "Outgoing - Check Writing";
            oForm.Width = 1020;
            oForm.Height = 590;

            oItem = createEditText(820, 15, 170, 14, "DocEntry", true, "@FTOOCW", "DocEntry");
            oItem.Enabled = false;
            oItem = createEditText(820, 15, 170, 14, "DocNum", true, "@FTOOCW", "DocNum");
            oItem.Enabled = false;
            oItem = createStaticText(670, 15, 120, 14, "stDocNum", "Document No", "DocNum");

            oItem = createEditText(820, 30, 170, 14, "DocDate", true, "@FTOOCW", "U_DocDate");
            oItem.Enabled = true;
            oItem = createStaticText(670, 30, 120, 14, "stDocDate", "Posting Date", "DocDate");

            oItem = createEditText(820, 45, 170, 14, "DueDate", true, "@FTOOCW", "U_DueDate");
            oItem.Enabled = true;
            oItem = createStaticText(670, 45, 120, 14, "stDueDate", "Due Date", "DueDate");

            oItem = createEditText(820, 60, 170, 14, "TaxDate", true, "@FTOOCW", "U_TaxDate");
            oItem.Enabled = true;
            oItem = createStaticText(670, 60, 120, 14, "stTaxDate", "Document Date", "TaxDate");

            oItem = createCombobox(820, 75, 170, 14, "Status", true, "@FTOOCW", "U_Status");
            oItem.DisplayDesc = true;
            oItem.Enabled = false;
            oItem = createStaticText(670, 75, 120, 14, "stStatus", "Status", "Status");

            oItem = createEditText(820, 105, 170, 14, "PrprdBy", true, "@FTOOCW", "U_PrprdBy");
            oItem.Enabled = false;
            oItem = createStaticText(670, 105, 120, 14, "stPrprdBy", "Prepared By", "PrprdBy");

            oItem = createEditText(820, 120, 170, 14, "ApprvdBy", true, "@FTOOCW", "U_ApprvdBy");
            oItem.Enabled = false;
            oItem = createStaticText(670, 120, 120, 14, "stApprvdBy", "Approved By", "ApprvdBy");

            oItem = createEditText(820, 135, 170, 14, "OVPMDocEnt", true, "@FTOOCW", "U_OVPMDocEnt");
            oItem.Enabled = false;
            oItem = createLinkButton("lnkPay", oItem, SAPbouiCOM.BoLinkedObject.lf_VendorPayment);
            oItem = createEditText(820, 135, 170, 14, "OVPMDocNum", true, "@FTOOCW", "U_OVPMDocNum");
            oItem.Enabled = false;
            oItem = createStaticText(670, 135, 120, 14, "stOutPay", "Outgoing Payment", "OVPMDocEnt");

            oItem = createEditText(820, 150, 170, 14, "RelDate", true, "@FTOOCW", "U_RelDate");
            oItem.Enabled = false;
            oItem = createStaticText(670, 150, 120, 14, "stRelDate", "Released Date", "RelDate");

            oItem = createEditText(820, 165, 170, 14, "ClrDate", true, "@FTOOCW", "U_ClrDate");
            oItem.Enabled = false;
            oItem = createStaticText(670, 165, 120, 14, "stClrDate", "Clearing Date", "ClrDate");

            oItem = createEditText(820, 180, 170, 14, "TransId", true, "@FTOOCW", "U_TransId");
            oItem.Enabled = false;
            oItem = createLinkButton("lnkJE", oItem, SAPbouiCOM.BoLinkedObject.lf_JournalPosting);
            oItem = createStaticText(670, 180, 120, 14, "stTransId", "JE ReClass", "TransId");

            oItem = createEditText(820, 195, 170, 14, "CanDate", true, "@FTOOCW", "U_CanDate");
            oItem.Enabled = false;
            oItem = createStaticText(670, 195, 120, 14, "stCanDate", "Canceled Date", "CanDate");

            oItem = createCombobox(200, 15, 170, 14, "TransType", true, "@FTOOCW", "U_TransType");
            oItem.DisplayDesc = true;
            oItem.Enabled = true;
            oItem = createStaticText(6, 15, 150, 14, "stTrnTyp", "Transaction Type", "TransType");

            oItem = createEditText(200, 30, 170, 14, "CardCode", true, "@FTOOCW", "U_CardCode");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.ChooseFromListUID = "cflBPS";
            oEditText.ChooseFromListAlias = "CardCode";
            oItem = createLinkButton("lnkBP", oItem, SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
            oItem.Enabled = true;
            oItem = createStaticText(6, 30, 150, 14, "stCardCode", "BP Code", "CardCode");

            oItem = createEditText(200, 45, 170, 14, "CardName", true, "@FTOOCW", "U_CardName");
            oItem.Enabled = false;
            oItem = createStaticText(6, 45, 150, 14, "stCardName", "BP Name", "CardName");

            oItem = createEditText(200, 60, 347, 14, "PayName", true, "@FTOOCW", "U_PayName");
            oItem.Enabled = true;
            oItem = createStaticText(6, 60, 150, 14, "stPayName", "Pay To Name", "PayName");

            oItem = createCombobox(200, 90, 170, 14, "Bank", true, "@FTOOCW", "U_Bank");

            oItem.DisplayDesc = true;
            oItem.Enabled = true;

            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"BankCode\", \"BankName\" FROM ODSC");

            if (oRecordset.RecordCount > 0)
            {
                oRecordset.MoveFirst();
                for (int i = 0; i <= oRecordset.RecordCount - 1; i++)
                {

                    oComboBox.ValidValues.Add(oRecordset.Fields.Item("BankCode").Value.ToString(), oRecordset.Fields.Item("BankName").Value.ToString());
                    oRecordset.MoveNext();
                }
            }

            oItem = createStaticText(6, 90, 150, 14, "stBank", "Bank", "Bank");

            oItem = createCombobox(200, 105, 170, 14, "Account", true, "@FTOOCW", "U_Account");
            oItem.DisplayDesc = true;
            oItem.Enabled = true;

            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oComboBox.ValidValues.Add("", "");

            oItem = createStaticText(6, 105, 150, 14, "stAccount", "Bank Account", "Account");

            oItem = createEditText(200, 120, 170, 14, "Branch", true, "@FTOOCW", "U_Branch");
            oItem.Enabled = false;
            oItem = createStaticText(6, 120, 150, 14, "stBranch", "Branch", "Branch");

            oItem = createEditText(200, 135, 170, 14, "BGLAcctCod", true, "@FTOOCW", "U_BGLAcctCode");
            oItem.Enabled = false;
            oItem = createLinkButton("lnkCA", oItem, SAPbouiCOM.BoLinkedObject.lf_GLAccounts);
            oItem = createEditText(373, 135, 170, 14, "BGLAcctNam", true, "@FTOOCW", "U_BGLAcctName");
            oItem.Enabled = false;
            oItem = createStaticText(6, 135, 150, 14, "stBGLCode", "Bank Account Code", "BGLAcctCod");

            oItem = createEditText(200, 150, 170, 14, "RGLAcctCod", true, "@FTOOCW", "U_RGLAcctCode");
            oItem.Enabled = false;
            oItem = createLinkButton("lnkRA", oItem, SAPbouiCOM.BoLinkedObject.lf_GLAccounts);
            oItem = createEditText(373, 150, 170, 14, "RGLAcctNam", true, "@FTOOCW", "U_RGLAcctName");
            oItem.Enabled = false;
            oItem = createStaticText(6, 150, 150, 14, "stRGLCode", "ReClass Account Code", "RGLAcctCod");

            oItem = createEditText(200, 165, 170, 14, "CheckNo", true, "@FTOOCW", "U_CheckNo");
            oItem.Enabled = true;
            oItem = createStaticText(6, 165, 150, 14, "stCheckNo", "Check No", "CheckNo");

            oItem = createCombobox(200, 180, 170, 14, "RBranch", true, "@FTOOCW", "U_RBranch");
            oItem.DisplayDesc = true;
            oItem.Enabled = true;

            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"Code\", \"Name\" FROM \"@FTBARB\" ");

            if (oRecordset.RecordCount > 0)
            {
                oRecordset.MoveFirst();
                for (int i = 0; i <= oRecordset.RecordCount - 1; i++)
                {

                    oComboBox.ValidValues.Add(oRecordset.Fields.Item("Code").Value.ToString(), oRecordset.Fields.Item("Name").Value.ToString());
                    oRecordset.MoveNext();
                }
            }
            oItem = createStaticText(6, 180, 150, 14, "stRBranch", "Releasing Branch", "RBranch");

            oItem = createCombobox(200, 195, 170, 14, "PntCntr", true, "@FTOOCW", "U_PntCntr");
            oItem.DisplayDesc = true;
            oItem.Enabled = true;

            oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"Code\", \"Name\" FROM \"@FTBAPC\" ");

            if (oRecordset.RecordCount > 0)
            {
                oRecordset.MoveFirst();
                for (int i = 0; i <= oRecordset.RecordCount - 1; i++)
                {

                    oComboBox.ValidValues.Add(oRecordset.Fields.Item("Code").Value.ToString(), oRecordset.Fields.Item("Name").Value.ToString());
                    oRecordset.MoveNext();
                }
            }
            oItem = createStaticText(6, 195, 150, 14, "stPntCntr", "Printing Center", "PntCntr");


            oItem = createCombobox(200, 210, 170, 14, "CrsChk", true, "@FTOOCW", "U_CrsChk");
            oItem.DisplayDesc = true;
            oItem.Enabled = true;
            oItem = createStaticText(6, 210, 150, 14, "stCrsChk", "Cross Check", "CrsChk");

            oItem = createMatrix(6, 240, 1002, 140, "grd1");
            oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            oColumn = createMatrixEditText("grd1", 20, "LineId", "#", true, "@FTOCW1", "LineId");
            oColumn.Editable = false;

            oColumn = createMatrixCheckBox("grd1", 30, "Select", "", "Y", "N", true, "@FTOCW1", "U_Select");
            oColumn.Editable = true;

            oColumn = createMatrixEditText("grd1", 120, "BaseType", "Base Type", true, "@FTOCW1", "U_BaseType");
            oColumn.Visible = false;

            oColumn = createMatrixEditText("grd1", 120, "DocLine", "Document Line", true, "@FTOCW1", "U_DocLine");
            oColumn.Visible = false;

            oColumn = creatematrixlinkedbutton("grd1", 20, "BaseEntry", "", SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoice, true, "@FTOCW1", "U_BaseEntry");
            oColumn.TitleObject.Sortable = true;
            oColumn.Editable = false;

            oColumn = createMatrixEditText("grd1", 127, "BaseNum", "Document No.", true, "@FTOCW1", "U_BaseNum");
            oColumn.Editable = false;

            oColumn = createMatrixEditText("grd1", 130, "InstId", "Installment", true, "@FTOCW1", "U_InstId");
            oColumn.Editable = false;

            oColumn = createMatrixEditText("grd1", 130, "DocDate", "Posting Date", true, "@FTOCW1", "U_DocDate");
            oColumn.Editable = false;

            oColumn = createMatrixEditText("grd1", 120, "DueDate", "Due Date", true, "@FTOCW1", "U_DueDate");
            oColumn.Editable = false;

            oColumn = createMatrixEditText("grd1", 110, "OverDue", "Over Due Days", true, "@FTOCW1", "U_OverDue");
            oColumn.Editable = false;
            oColumn.RightJustified = true;

            oColumn = createMatrixEditText("grd1", 130, "BalDue", "Balance Due", true, "@FTOCW1", "U_BalDue");
            oColumn.Editable = false;
            oColumn.RightJustified = true;

            oColumn = createMatrixEditText("grd1", 163, "TotPay", "Total Payment", true, "@FTOCW1", "U_TotPay");
            oColumn.Editable = true;
            oColumn.RightJustified = true;

            oItem = createEditText(820, 390, 170, 14, "TotalDue", true, "@FTOOCW", "U_TotalDue");
            oItem.Enabled = false;
            oItem.RightJustified = true;
            oItem = createStaticText(670, 390, 120, 14, "stTotDue", "Total Amount Due", "TotalDue");

            oItem = createExtEditText(150, 420, 170, 70, "Comments", true, "@FTOOCW", "U_Comments");
            oItem.Enabled = true;
            oItem = createStaticText(6, 420, 120, 14, "stComments", "Remarks", "Comments");

            oItem = createExtEditText(470, 420, 170, 70, "BnkRmks", true, "@FTOOCW", "U_BnkRmks");
            oItem.Enabled = true;
            oItem = createStaticText(340, 420, 120, 14, "stBnkRmks", "Bank Remarks", "BnkRmks");

            oItem = createExtEditText(820, 420, 170, 70, "AppRmks", true, "@FTOOCW", "U_AppRmks");
            oItem.Enabled = true;
            oItem = createStaticText(670, 420, 120, 14, "stAppRmks", "Approval Remarks", "AppRmks");

            oItem = createButton(6, 520, 80, 19, "1", "");
            oItem = createButton(90, 520, 80, 19, "2", "E&xit");

            oItem = createButton(646, 520, 170, 19, "btnCan", "&Canceled Check");
            oItem = createButton(820, 520, 170, 19, "btnEmail", "Send &E-Mail Notification");

            oForm.DataBrowser.BrowseBy = "DocEntry";

            if (string.IsNullOrEmpty(is_docno))
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            else
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

            dblTotPay = 0;

            strDBType = DI.oCompany.DbServerType.ToString();

            oForm.Freeze(false);
        }
        public override void matrixlinkpressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.matrixlinkpressed(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.LinkedButton oLinkButton;

            string strBaseType, strBaseEntry;

            if (pVal.BeforeAction)
            {
                if (pVal.ItemUID == "grd1")
                {
                    if (pVal.ColUID == "BaseEntry")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)getItemSpecific("grd1");

                        strBaseType = getColumnString("grd1", "BaseType", pVal.Row, "");
                        oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("BaseEntry");

                        oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;

                        switch (strBaseType)
                        {
                            case "30":
                                oLinkButton.LinkedObjectType = "30";
                                break;
                            case "18":
                                oLinkButton.LinkedObjectType = "18";
                                break;
                            case "19":
                                oLinkButton.LinkedObjectType = "19";
                                break;
                            case "204":
                                oLinkButton.LinkedObjectType = "204";
                                break;
                        }
                    }
                }
            }
        }
        public override void onadd(ref bool BubbleEvent)
        {
            base.onadd(ref BubbleEvent);

            oForm.Freeze(true);

            uf_validateGrid();

            if (!(uf_validateSave()))
            {

                BubbleEvent = false;
                oForm.Freeze(false);

                return;
            }

            oForm.Freeze(false);
        }
        public override void onaddmode()
        {
            base.onaddmode();

            SAPbouiCOM.EditText oEditText;

            SAPbobsCOM.Recordset oRecordset;

            string strTransType;

            oForm.Freeze(true);

            dblTotPay = 0;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"U_NAME\" FROM \"OUSR\" WHERE \"USERID\" = '" + DI.oCompany.UserSignature.ToString() + "' ");

            setItemString("PrprdBy", oRecordset.Fields.Item("U_NAME").Value.ToString());

            oEditText = (SAPbouiCOM.EditText)getItemSpecific("CardCode");

            setItemEnabled("TransType", true);

            strTransType = getItemSelectedValue("TransType", "");
            if (strTransType == "S")
            {
                oEditText.ChooseFromListUID = "cflBPS";
                oEditText.ChooseFromListAlias = "CardCode";

                setItemEnabled("PayName", true);
                setItemEnabled("CardCode", true);

            }
            else if (strTransType == "C")
            {
                oEditText.ChooseFromListUID = "cflBPC";
                oEditText.ChooseFromListAlias = "CardCode";

                setItemEnabled("PayName", true);
                setItemEnabled("CardCode", true);
            }
            else
            {
                setItemEnabled("PayName", false);
                setItemEnabled("CardCode", false);
            }


            setItemEnabled("Bank", true);
            setItemEnabled("Account", true);
            setItemEnabled("CheckNo", true);
            setItemEnabled("RBranch", true);
            setItemEnabled("PntCntr", true);
            setItemEnabled("CrsChk", true);

            setItemEnabled("Status", false);

            setItemEnabled("DocDate", true);
            setItemEnabled("DueDate", true);
            setItemEnabled("TaxDate", true);

            setItemEnabled("grd1", true);

            setItemEnabled("Comments", true);
            setItemEnabled("BnkRmks", true);
            setItemEnabled("AppRmks", false);

            setItemEnabled("btnCan", false);
            setItemEnabled("btnEmail", false);

            oForm.Update();
            oForm.Freeze(false);

        }
        public override void onfindmode()
        {
            base.onfindmode();

            SAPbouiCOM.Button oButton;

            oForm.Freeze(true);

            oButton = (SAPbouiCOM.Button)oForm.Items.Item("2").Specific;
            oButton.Caption = "E&xit";

            setItemEnabled("DocEntry", true);
            setItemEnabled("DocNum", true);
            setItemEnabled("DocDate", true);
            setItemEnabled("DueDate", true);
            setItemEnabled("TaxDate", true);
            setItemEnabled("CardCode", true);
            setItemEnabled("CardName", true);
            setItemEnabled("Status", true);

            setItemEnabled("TransType", false);
            setItemEnabled("PayName", false);
            setItemEnabled("Bank", false);
            setItemEnabled("Account", false);

            setItemEnabled("CheckNo", true);
            setItemEnabled("RBranch", false);
            setItemEnabled("PntCntr", false);
            setItemEnabled("CrsChk", false);

            setItemEnabled("grd1", false);

            setItemEnabled("Comments", false);
            setItemEnabled("BnkRmks", false);

            if (!string.IsNullOrEmpty(is_docno))
            {
                setItemString("DocEntry", is_docno);
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                is_docno = "";
                setItemEnabled("DocEntry", false);
            }

            oForm.Freeze(false);
        }
        public override void onokmode()
        {
            base.onokmode();

            SAPbouiCOM.Button oButton;
            SAPbobsCOM.Recordset oRecordset;

            string strStatus = "", strQuery, strDocNum;

            oForm.Freeze(true);

            strDocNum = getItemString("DocNum");

            strQuery = string.Format("SELECT \"U_Status\" FROM \"@FTOOCW\" WHERE \"DocNum\" = '{0}' ", strDocNum);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
                strStatus = oRecordset.Fields.Item("U_Status").Value.ToString();

            switch (strStatus)
            {
                case "O":

                    setItemEnabled("TransType", false);
                    setItemEnabled("CardCode", false);
                    setItemEnabled("PayName", true);

                    setItemEnabled("Bank", true);
                    setItemEnabled("Account", true);
                    setItemEnabled("CheckNo", true);
                    setItemEnabled("RBranch", true);
                    setItemEnabled("PntCntr", true);
                    setItemEnabled("CrsChk", true);

                    setItemEnabled("DocDate", true);
                    setItemEnabled("DueDate", true);
                    setItemEnabled("TaxDate", true);

                    setItemEnabled("Status", false);

                    setItemEnabled("grd1", true);

                    setItemEnabled("Comments", true);
                    setItemEnabled("BnkRmks", true);
                    setItemEnabled("AppRmks", false);

                    break;

                case "A":
                case "P":
                case "X":
                case "R":
                case "D":
                case "C":

                    setItemEnabled("TransType", false);
                    setItemEnabled("CardCode", false);
                    setItemEnabled("PayName", false);

                    setItemEnabled("Bank", false);
                    setItemEnabled("Account", false);
                    setItemEnabled("CheckNo", false);
                    setItemEnabled("RBranch", false);
                    setItemEnabled("PntCntr", false);
                    setItemEnabled("CrsChk", false);

                    setItemEnabled("DocDate", false);
                    setItemEnabled("DueDate", false);
                    setItemEnabled("TaxDate", false);

                    setItemEnabled("Status", false);

                    setItemEnabled("grd1", false);

                    setItemEnabled("Comments", true);
                    setItemEnabled("BnkRmks", true);
                    setItemEnabled("AppRmks", false);

                    break;

            }

            if (strStatus == "P")
                setItemEnabled("btnEmail", true);
            else
                setItemEnabled("btnEmail", false);

            if (strStatus == "O" || strStatus == "A" || strStatus == "R" || strStatus == "P")
                setItemEnabled("btnCan", true);
            else
                setItemEnabled("btnCan", false);

            if (strStatus == "C" || strStatus == "X")
            {
                setItemEnabled("BnkRmks", false);
                setItemEnabled("Comments", false);

            }

            setItemEnabled("DocEntry", false);
            setItemEnabled("DocNum", false);
            setItemEnabled("Status", false);
            setItemEnabled("CardName", false);

            oButton = (SAPbouiCOM.Button)oForm.Items.Item("2").Specific;
            oButton.Caption = "E&xit";

            oForm.Freeze(false);

        }
        public override void onupdate(ref bool BubbleEvent)
        {
            base.onupdate(ref BubbleEvent);

            oForm.Freeze(true);

            uf_validateGrid();

            if (!(uf_validateSave()))
            {

                BubbleEvent = false;
                oForm.Freeze(false);

                return;
            }

            oForm.Freeze(false);

        }
        public override void validate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.validate(FormUID, ref pVal, ref BubbleEvent);

            SAPbouiCOM.CheckBox oCheckBox;
            SAPbouiCOM.Matrix oMatrix;

            double dblTotPay, dblBalDue;

            if (!pVal.BeforeAction && pVal.ItemChanged)
            {
                switch (pVal.ItemUID)
                {
                    case "grd1":

                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;
                        
                        switch (pVal.ColUID)
                        {
                            case "TotPay":

                                oForm.Freeze(true);

                                dblTotPay = Convert.ToDouble(getColumnString("grd1", "TotPay", pVal.Row, ""));
                                dblBalDue = Convert.ToDouble(getColumnString("grd1", "BalDue", pVal.Row, ""));

                                if (dblTotPay > dblBalDue)
                                {
                                    UI.SBO_Application.StatusBar.SetText("Cannot made payment greater than balance due.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                                    setColumnString("grd1", "TotPay", pVal.Row, dblBalDue.ToString());
                                    
                                    oForm.Freeze(false);
                                    return;
                                }

                                uf_Compute();

                                oForm.Update();
                                oForm.Freeze(false);

                                break;

                        }
                        break;
                }
            }

            GC.Collect();
        }
        private void uf_ListInvoices(string strCardCode)
        {
            string strQuery;

            int intLineId;

            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Matrix oMatrix;

            dblTotPay = 0;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;

            oMatrix.Clear();
            oForm.DataSources.DBDataSources.Item("@FTOCW1").Clear();

            if (strDBType == "dst_HANADB")
                strQuery = string.Format("CALL \"FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKWRITING\" ('{0}') ", strCardCode);
            else
                strQuery = string.Format("EXEC FTSI_BANKINGADDON_IMPORT_OUTGOING_CHECKWRITING '{0}' ", strCardCode);


            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
            {
                intLineId = oMatrix.VisualRowCount + 1;

                UI.displayStatus();
                UI.changeStatus("Loading Open Invoices, Please Wait...");

                oRecordset.MoveFirst();

                while (!(oRecordset.EoF))
                {

                    oForm.DataSources.DBDataSources.Item("@FTOCW1").InsertRecord(0);
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").Offset = oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1;
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("LineId", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, intLineId.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_Select", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, "N");
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_BaseType", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("ObjType").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_BaseEntry", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("DocEntry").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_BaseNum", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("DocNum").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_DocLine", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("DocLine").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_InstId", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("InstId").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_DocDate", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, GlobalFunction.f_convert_date_sbodate(Convert.ToDateTime(oRecordset.Fields.Item("DocDate").Value.ToString())));
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_DueDate", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, GlobalFunction.f_convert_date_sbodate(Convert.ToDateTime(oRecordset.Fields.Item("DocDueDate").Value.ToString())));
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_OverDue", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("OverDue").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_BalDue", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("DocTotal").Value.ToString());
                    oForm.DataSources.DBDataSources.Item("@FTOCW1").SetValue("U_TotPay", oForm.DataSources.DBDataSources.Item("@FTOCW1").Size - 1, oRecordset.Fields.Item("DocTotal").Value.ToString());

                    intLineId++;

                    oMatrix.AddRow(1, -1);

                    oRecordset.MoveNext();
                }

                UI.hideStatus();

            }
        }
        private void uf_Compute()
        {
            SAPbouiCOM.Matrix oMatrix, oMatrix1;
            SAPbouiCOM.CheckBox oCheckBox;

            oForm.Freeze(true);

            double dblTotDue = 0;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;

            if (oMatrix.VisualRowCount > 0)
            {
                for (int ll_row = 1; ll_row <= oMatrix.VisualRowCount; ll_row++)
                {
                    oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Select").Cells.Item(ll_row).Specific;

                    if (oCheckBox.Checked == true)
                        dblTotDue = dblTotDue + Convert.ToDouble(getColumnString("grd1", "TotPay", ll_row, ""));

                }
            }

            dblTotPay = dblTotDue;

            setItemString("TotalDue", dblTotPay.ToString());

            oForm.Update();
            oForm.Freeze(false);
        }
        private bool uf_Cancel(string strDocNum, string strOVPMDocEnt)
        {
            string strQuery, strErrMsg;
            int intErrCode;

            SAPbobsCOM.Payments oPayments;

            if (!DI.oCompany.InTransaction)
                DI.oCompany.StartTransaction();

            if (!(string.IsNullOrEmpty(strOVPMDocEnt)))
            {

                oPayments = null;
                oPayments = (SAPbobsCOM.Payments)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

                if (oPayments.GetByKey(Convert.ToInt32(strOVPMDocEnt)))
                {
                    try
                    {
                        if (oPayments.Cancel() == 0)
                        {
                            strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'X', \"U_CanDate\" = to_date('{0}', 'MM/DD/YYYY') WHERE \"DocNum\" = '{1}' ", DateTime.Today.ToString("MM/dd/yyyy"), strDocNum);
                            if (!(DI.executeQuery(strQuery)))
                            {
                                UI.SBO_Application.StatusBar.SetText(string.Format("Error updating Check Writing Status."), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                                if (DI.oCompany.InTransaction)
                                    DI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                                return false;
                            }
                        }
                        else 
                        {
                            intErrCode = DI.oCompany.GetLastErrorCode();
                            strErrMsg = DI.oCompany.GetLastErrorDescription();

                            UI.SBO_Application.StatusBar.SetText(string.Format("Error processing cancellation of link Outgoing Payment. {0} - {1}.", intErrCode, strErrMsg), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        UI.SBO_Application.StatusBar.SetText(string.Format("Error processing cancellation of link Outgoing Payment."), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        
                        GlobalFunction.fileappend(ex.Message.ToString());

                        if (DI.oCompany.InTransaction)
                            DI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                        return false;
                    }
                }
            }
            else
            {
                strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_Status\" = 'X', \"U_CanDate\" = to_date('{0}', 'MM/DD/YYYY') WHERE \"DocNum\" = '{1}' ", DateTime.Today.ToString("MM/dd/yyyy"), strDocNum);
                if (!(DI.executeQuery(strQuery)))
                {
                    UI.SBO_Application.StatusBar.SetText(string.Format("Error updating Check Writing Status."), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                    if (DI.oCompany.InTransaction)
                        DI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);

                    return false;
                }
            }

            if (DI.oCompany.InTransaction)
                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

            setItemEnabled("DocNum", true);
            setItemString("DocNum", strDocNum);
            itemclick("1");
            setItemEnabled("DocNum", false);

            GC.Collect();

            return true;
        }
        private void uf_validateGrid()
        {
            SAPbouiCOM.Matrix oMatrix;
            SAPbouiCOM.CheckBox oCheckBox;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;
            for (int ll_row = oMatrix.VisualRowCount; ll_row > 0; ll_row--)
            {
                oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Select").Cells.Item(ll_row).Specific;
                if (oCheckBox.Checked == false)
                    oMatrix.DeleteRow(ll_row);

            }
            for (int ll_row = 1; ll_row <= oMatrix.VisualRowCount; ll_row++)
            {
                setColumnString("grd1", "LineId", ll_row, ll_row.ToString());
            }

            GC.Collect();
        }
        private bool uf_validateSave()
        {
           
            SAPbouiCOM.Matrix oMatrix;

            string strTransType;

            double dblTotDue;

            strTransType = getItemSelectedValue("TransType", "");

            if (strTransType != "A")
            {
                if (string.IsNullOrEmpty(getItemString("CardCode")))
                {
                    UI.SBO_Application.StatusBar.SetText("Business Partner is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Freeze(false);
                    return false;
                }
            }

            if (string.IsNullOrEmpty(getItemString("DocDate")))
            {
                UI.SBO_Application.StatusBar.SetText("Posting Date is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemString("DueDate")))
            {
                UI.SBO_Application.StatusBar.SetText("Due Date is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemString("TaxDate")))
            {
                UI.SBO_Application.StatusBar.SetText("Due Date is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemSelectedValue("Bank", "")))
            {
                UI.SBO_Application.StatusBar.SetText("Bank is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemSelectedValue("RBranch", "")))
            {
                UI.SBO_Application.StatusBar.SetText("Releasing Branch is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemString("BGLAcctCod")))
            {
                UI.SBO_Application.StatusBar.SetText("Bank GL Account is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            if (string.IsNullOrEmpty(getItemString("RGLAcctCod")))
            {
                UI.SBO_Application.StatusBar.SetText("ReClass GL Account is missing.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            dblTotDue = Convert.ToDouble(getItemString("TotalDue"));
            if (dblTotDue <= 0)
            {
                UI.SBO_Application.StatusBar.SetText("Cannot add Check Writing without Total Due.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grd1").Specific;
            if (oMatrix.VisualRowCount <= 0)
            {
                UI.SBO_Application.StatusBar.SetText("Cannot add Check Writing without line details.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
                return false;
            }

            return true;
        }
        private void uf_SendEmailNotif(string strDocNum)
        {

            string strQuery, strDocEntry, strInvNo;
            string strSubject, strMailTo, strMailCC, strMailBody;

            SAPbobsCOM.Recordset oRecordset, oRSInv;

            UI.displayStatus();
            UI.changeStatus("Sending E-Mail Notifications. Please Wait...");

            strMailTo = "";
            strInvNo = "";
            strSubject = "";
            strMailTo = "";
            strMailCC = "";

            strQuery = string.Format("SELECT OOCW.\"DocEntry\", OOCW.\"DocNum\", " +
                                        "       OOCW.\"U_CardName\", OCRD.\"E_Mail\", OOCW.\"U_DueDate\", ODSC.\"BankName\", OOCW.\"U_CheckNo\", " +
                                        "       OADM.\"AliasName\", OADM.\"CompnyName\", " +
                                        "       OOCW.\"U_TotalDue\", BARB.\"Name\", BARB.\"U_Address\" " +
                                        "FROM \"@FTOOCW\" OOCW LEFT JOIN ODSC ON OOCW.\"U_Bank\" = ODSC.\"BankCode\" " +
                                        "                      LEFT JOIN OCRD ON OOCW.\"U_CardCode\" = OCRD.\"CardCode\" " +
                                        "                      LEFT JOIN \"@FTBARB\" BARB ON OOCW.\"U_RBranch\" = BARB.\"Code\", OADM " +
                                        "WHERE OOCW.\"DocNum\" = '{0}' ", strDocNum);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
            {
                strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();

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
                        strQuery = string.Format("UPDATE \"@FTOOCW\" SET \"U_EmailNotif\" = 'Y' WHERE \"DocNum\" = '{0}' ", strDocNum);

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

                UI.hideStatus();

            }

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
