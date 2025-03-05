using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;


namespace AddOn.Setup
{
    public partial class userform_Setup : AddOn.Form
    {
        private static string is_docno;
        public userform_Setup()
        {
            InitializeComponent();
        }
        public override void onGetCreationParams(ref SAPbouiCOM.BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
            base.onGetCreationParams(ref io_BorderStyle, ref is_FormType, ref is_ObjectType, ref xmlPath);

            is_ObjectType = "FTOBAS";
            is_FormType = "100000007";
            io_BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);

            SAPbobsCOM.Recordset oRecordset;

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.CheckBox oCheckBox;

            oForm.Freeze(true);

            oForm.Title = "Banking AddOn Setup";
            oForm.Width = 650;
            oForm.Height = 350;

            oItem = createEditText(820, 15, 170, 14, "Code", true, "@FTOBAS", "Code");
            oItem.Enabled = false;

            oItem = createCheckBox(20, 30, 250, 15, "GenApp", "Y", "N", true, "@FTOBAS", "U_GenApp");
            oCheckBox = (SAPbouiCOM.CheckBox)oItem.Specific;
            oCheckBox.Caption = "Generate Bank File Upon Approval";

            oItem = createButton(6, 280, 80, 19, "1", "");

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"Code\" FROM \"@FTOBAS\" ");

            if (oRecordset.RecordCount > 0)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
            else
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

            oForm.Freeze(false);
        }
        public override void onaddsuccess()
        {
            base.onaddsuccess();

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        }
        public override void onaddmode()
        {
            base.onaddmode();

            oForm.Freeze(true);

            setItemEnabled("Code", false);

            setItemString("Code", "1");

            setItemEnabled("Code", true);

            oForm.Update();
            oForm.Freeze(false);

        }
        public override void onfindmode()
        {
            base.onfindmode();

            SAPbouiCOM.Button oButton;

            oForm.Freeze(true);

            setItemEnabled("Code", true);

            setItemString("Code", "1");

            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            setItemEnabled("Code", false);
            
            oForm.Freeze(false);
        }
    }
}
