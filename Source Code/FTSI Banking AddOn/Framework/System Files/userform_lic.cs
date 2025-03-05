using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace AddOn
{
    public partial class userform_lic : AddOn.Form
    {
        public userform_lic()
        {
            InitializeComponent();
        }
        public override void onGetCreationParams(ref SAPbouiCOM.BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
            base.onGetCreationParams(ref io_BorderStyle, ref is_FormType, ref is_ObjectType, ref xmlPath);
            is_FormType = "Upload License";
        }
        public override void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
            base.onFormCreate(ref ab_visible, ref ab_center);
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.ActiveX oActiveX;

            oForm.DataSources.UserDataSources.Add("Path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 214);

            oForm.Title = "Upload License";
            oForm.Width = 500;
            oForm.Height = 150;


            oItem = createButton(6, 90, 65, 18, "Upload", "Upload");
            oItem = createButton(72, 90, 65, 18, "2", "");
            oItem = createButton(443, 30, 20, 14, "Browse", "...");

            oItem = createEditText(90, 30, 350, 14, "Path", true, "", "Path");
            oItem.Enabled = false;
            oItem = createStaticText(6, 30, 80, 14, "stPath", "File Path", "Path");

            oItem = oForm.Items.Add("xdialog", SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X);
            oItem.Left = 1400;
            oItem.Top = 6;
            oActiveX = (SAPbouiCOM.ActiveX)oItem.Specific;
            oActiveX.ClassID = "FTSBOCommonDialogXControl.FTSBOCommonDialogX";
            oItem.Visible = false;
        }
        public override void itempressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);

            if (pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "Browse":
                        //browse();
                        break;
                    case "Upload":
                        UI.displayStatus();
                        UI.changeStatus("Uploading...");
                        upload(ref BubbleEvent);
                        UI.hideStatus();
                        break;
                }
            }
        }
        //private void browse()
        //{
        //    SAPbouiCOM.ActiveX oActivex;
        //    FTSBOCommonDialogXControl.FTSBOCommonDialogX oCommonDialogX;

        //    string ls_pathname;

        //    oActivex = (SAPbouiCOM.ActiveX)oForm.Items.Item("xdialog").Specific;
        //    oCommonDialogX = (FTSBOCommonDialogXControl.FTSBOCommonDialogX)oActivex.Object;
        //    oCommonDialogX.Filter = "(*.txt)|*.txt";
        //    oCommonDialogX.DialogTitle = "Select File";
        //    oCommonDialogX.FilterIndex = 0;

        //    if (oCommonDialogX.ShowOpen())
        //    {
        //        ls_pathname = oCommonDialogX.FileName;
        //        setItemString("Path", ls_pathname);
        //    }
        //}
        private void upload(ref bool BubbleEvent)
        {
            StreamReader SR;
            string ls_pathname, ls_addon, ls_hardwarekey, ls_expiry, ls_key;

            ls_pathname = getItemString("Path");
            SR = File.OpenText(ls_pathname);
            SR.ReadLine();
            ls_addon = SR.ReadLine();
            if (ls_addon.Substring(0, 7) == "Add-On=")
            {
                ls_addon = ls_addon.Substring(7).Trim();
            }
            if (globalvar.addondescription != ls_addon)
            {
                UI.SBO_Application.StatusBar.SetText("License Invalid.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                GC.Collect();
                return;
            }
            ls_hardwarekey = SR.ReadLine();
            if (ls_hardwarekey.Substring(0, 13) == "HARDWARE-KEY=")
            {
                ls_hardwarekey = ls_hardwarekey.Substring(13).Trim();
            }
            ls_expiry = SR.ReadLine();
            if (ls_expiry.Substring(0, 7) == "EXPIRY=")
            {
                ls_expiry = ls_expiry.Substring(7).Trim();
            }
            ls_key = SR.ReadLine();
            if (ls_key.Substring(0, 4) == "KEY=")
            {
                ls_key = ls_key.Substring(4).Trim();
            }
            SR.Close();
            queryopen(1, "SELECT * FROM [@FTLIC] WHERE U_HKEY='" + ls_hardwarekey + "' AND U_ADDON = '" + globalvar.addondescription + "'");
            if (queryrows(1) == 0)
            {
                UI.SBO_Application.StatusBar.SetText("License Invalid.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                GC.Collect();
                return;
            }
            if (nv_string.encrypt(ls_hardwarekey + ls_addon + ls_expiry) != ls_key)
            {
                UI.SBO_Application.StatusBar.SetText("License Invalid.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                GC.Collect();
                return;
            }
            if (!DI.oCompany.InTransaction)
                DI.oCompany.StartTransaction();
            //MessageBox.Show("UPDATE [@FT_LIC] SET U_HKEY = '" + ls_hardwarekey + "',U_ADDON = '" + ls_addon + "',U_EXPIRY = '" + ls_expiry + "',U_KEY = '" + ls_key + "' Where DocEntry = 1");
            if (!DI.executeQuery("UPDATE [@FTLIC] SET U_EXPIRY = '" + ls_expiry + "',U_KEY = '" + ls_key + "' Where DocEntry = 1"))
            {
                string errmsg;
                errmsg = DI.oCompany.GetLastErrorDescription();

                UI.SBO_Application.StatusBar.SetText("Failed to update License.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                if (DI.oCompany.InTransaction)
                    DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return;
            }
            if (DI.oCompany.InTransaction)
                DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            UI.SBO_Application.StatusBar.SetText("Operation successfully completed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            GC.Collect();
        }
    }
}

