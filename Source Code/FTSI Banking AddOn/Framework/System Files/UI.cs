using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;

//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : UI.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class UI : UserControl
    {
        public static SAPbouiCOM.Application SBO_Application;
        public static Form SBOCurrentForm;

        public static string generateFormUID()
        {
            string FormUID;

            SAPbouiCOM.Form oForm;

            FormUID = String.Concat("FT", globalvar.FormCount.ToString());
            try
            {
                oForm = UI.SBO_Application.Forms.Item(FormUID);
                globalvar.FormCount = globalvar.FormCount + 1;
                FormUID = generateFormUID();

            }
            catch
            {
                return FormUID;

            }

            return FormUID;

        }
        public static int generateFormIndex()
        {
            int x = 0;
            try
            {
                for (int i = 1; i < 100; i++)
                {
                    x = i;
                    globalvar.sboform[i].ToString();
                }
            }
            catch
            {
                return x;
            }

            return x;
        }
        public static void displayStatus()
        {
            SAPbouiCOM.Form oForm;
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.FormCreationParams creationPackage;
            SAPbouiCOM.StaticText oStaticText;
            try
            {

                creationPackage = (SAPbouiCOM.FormCreationParams)UI.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FTSTAT";
                creationPackage.FormType = "FTSTAT";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;
                oForm = UI.SBO_Application.Forms.AddEx(creationPackage);

                oForm.Left = 400;
                oForm.Width = 300;
                oForm.Top = 100;
                oForm.Height = 70;

                oItem = oForm.Items.Add("stStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 6;
                oItem.Width = 290;
                oItem.Top = 25;
                oItem.Height = 14;
                oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
                oStaticText.Caption = "Please Wait";

                oItem.Enabled = false;
                oForm.Visible = true;
                
            }
            catch (Exception e)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item("FTSTAT");
                    oForm.Visible = true;
                    oItem = oForm.Items.Item("stStatus");
                    oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
                    oStaticText.Caption = "Please Wait!";
                }
                catch
                {
                }

                //SBO_Application.MessageBox(e.Message);
            }
            creationPackage = null;
            GC.Collect();
        }
        public static void changeStatus(string as_status)
        {
            SAPbouiCOM.Form oForm;
            SAPbouiCOM.StaticText oStaticText;
            try
            {
                oForm = SBO_Application.Forms.Item("FTSTAT");
                oForm.Visible = true;
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("stStatus").Specific;
                oStaticText.Caption = as_status;

            }
            catch (Exception)
            {
            }
        }
        public static void hideStatus()
        {
            SAPbouiCOM.Form oForm;
            try
            {
                oForm = SBO_Application.Forms.Item("FTSTAT");
                oForm.Visible = false;
            }
            catch
            {
            }
        }
        public static bool ProccessMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent eventType, ref bool BubbleEvent)
        {
            DataRow[] oDataRow;
            oDataRow = globalvar.gdt_Form.Select("FormUID = '" + FormUID + "'");
            if (oDataRow.Length > 0)
            {
                try
                {
                    globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].menuevent(FormUID, ref eventType, ref BubbleEvent);
                }
                catch (Exception e)
                {
                    GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool ProccessItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            DataRow[] oDataRow;

            oDataRow = globalvar.gdt_Form.Select("FormUID = '" + FormUID + "'");
            if (oDataRow.Length > 0)
            {
                try
                {
                    Form.oForm = UI.SBO_Application.Forms.Item(FormUID);
                    globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].itemevent(FormUID, ref pVal, ref BubbleEvent);
                    
                    if (BubbleEvent == false)
                        return true;

                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].click(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].comboselect(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].datasourceload(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].doubleclick(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                            SBOCurrentForm = globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])];
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formactivate(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formclose(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                            SBOCurrentForm = null;
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formdeactivate(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formkeydown(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formload(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formmenuhilight(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].formresize(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:

                            DataRow Row;
                            int li_index = System.Convert.ToInt16(oDataRow[0]["Index"]);

                            globalvar.sboform[li_index].formunload(FormUID, ref pVal, ref BubbleEvent);
                            if (pVal.BeforeAction)
                            {
                                globalvar.sboform[li_index] = null;
                                Row = globalvar.gdt_Form.Rows.Find(li_index.ToString());
                                for (int i = 0; i < globalvar.gdt_Form.Rows.Count; i++)
                                {
                                    if (globalvar.gdt_Form.Rows[i].Equals(Row))
                                    {
                                        li_index = i;
                                        break;
                                    }
                                }
                                globalvar.gdt_Form.Rows[li_index].Delete();
                                BubbleEvent = false;
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].gotfocus(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].itempressed(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].keydown(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].lostfocus(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].matrixcollapsepressed(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].matrixlinkpressed(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].matrixload(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_MENU_CLICK:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].menuclick(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].validate(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            SAPbouiCOM.ChooseFromListEvent pVal2;
                            pVal2 = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].choosefromlist(FormUID, ref pVal2, ref BubbleEvent);
                            break;
                    }
                }
                catch (Exception e)
                {
                    GlobalFunction.fileappend("Errmsg -" + e.Message + ". Event - " + pVal.EventType.ToString());
                    //UI.SBO_Application.MessageBox(e.Message, 1, "OK", "", "");
                    if (DI.oCompany.InTransaction)
                        DI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    UI.hideStatus();
                }
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}