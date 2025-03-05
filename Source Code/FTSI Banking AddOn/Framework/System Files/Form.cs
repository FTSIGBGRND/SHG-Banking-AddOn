using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : Form.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class Form : UserControl
    {
        /// <summary>
        /// asdfadsf
        /// </summary>

        public Form()
        {
            InitializeComponent();
        }
        public static SAPbouiCOM.Form oForm;
        public static int il_CurrentMatrixRow;
        public static string is_CurrentMatrixUID = "";
        public static string is_CurrentMatrixColUID = "";
        public static bool ib_updatemode = false;
        public static bool ib_okmode = false;
        public static bool ib_adding = false;
        public static bool ib_updating = false;
        public static bool ib_udo = true;
        public static Form overlapForm = null;
        public static SAPbouiCOM.EditText targetEditText;
        public static Form targetsboform;
        public static bool overlap = false;
        public static string sItemQuery = "";
        public static SAPbouiCOM.Form FormModal;
        public static string currdocno = "";
        public static string closereason = "";
        public static DateTime closedate;
        public static string popupname = "";
        public static bool frompopup = false;
        

        public virtual void click(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void choosefromlist(string FormUID, ref SAPbouiCOM.ChooseFromListEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void comboselect(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void datasourceload(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void doubleclick(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formactivate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (frompopup)
                poponsave(ref BubbleEvent);
        }
        public virtual void formclose(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formdeactivate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formkeydown(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formload(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formmenuhilight(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formresize(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void formunload(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        
        public virtual void gotfocus(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.ItemUID != "")
                    {
                        if (oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            is_CurrentMatrixUID = pVal.ItemUID;
                            if (pVal.ColUID != "")
                                is_CurrentMatrixColUID = pVal.ColUID;
                            if (pVal.Row > 0)
                                il_CurrentMatrixRow = pVal.Row;
                            onsetmatrixmenu(is_CurrentMatrixUID, true);
                        }
                    }
                }
                //true;
            }
            catch (Exception e)
            {
                UI.SBO_Application.StatusBar.SetText(e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //false;
            }
        }
        
        public virtual void itemevent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            
            if (ib_updatemode == false && oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ib_updatemode = true;
                modechanged();
                return;
            }
            else if (ib_updatemode && oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ib_updatemode = false;
                modechanged();
                return;
            }
            if (ib_okmode == false && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                ib_okmode = true;
                modechanged();
                return;
            }
            else if (ib_okmode && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                ib_okmode = false;
                return;
            }
            if (pVal.ItemUID != "")
            {
                if (oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                {
                    if (pVal.Row > 0)
                        il_CurrentMatrixRow = pVal.Row;
                    if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_LOST_FOCUS &&
                       pVal.EventType != SAPbouiCOM.BoEventTypes.et_VALIDATE &&
                       pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
                        onmatrixeditchanging(pVal.ItemUID, pVal.ColUID, pVal.Row, (pVal.CharPressed == 9), ref BubbleEvent);
                }
            }
        }
        
        public virtual void itempressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.BeforeAction)
            {
                if (pVal.ItemUID == "1")
                {
                    try
                    {
                        switch (oForm.Mode)
                        {
                            case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                                ib_adding = true;
                                onadd(ref BubbleEvent);
                                if (BubbleEvent)
                                    onsave(1, ref BubbleEvent);
                                break;
                            case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                                ib_updating = true;
                                onupdate(ref BubbleEvent);
                                if (BubbleEvent)
                                    onsave(2, ref BubbleEvent);
                                break;
                            case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                                onfind(ref BubbleEvent);
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                        //System.Windows.Forms.MessageBox.Show(e.Message);
                        GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                    }
                }

                if ((pVal.ItemUID == "9999") || (pVal.ItemUID == "9998"))
                {
                    frompopup = true;
                }

            }
            else
            {
                if (pVal.ItemUID == "1")
                {
                    try
                    {
                        if (pVal.ActionSuccess)
                        {
                            if (ib_adding)
                            {
                                ib_adding = false;
                                onaddsuccess();
                                onsavesuccess(1);
                            }
                            else if (ib_updating)
                            {
                                ib_updating = false;
                                onupdatesuccess();
                                onsavesuccess(2);
                            }
                            modechanged();
                        }
                        else
                        {
                            if (ib_adding)
                            {
                                ib_adding = false;
                                onadderror();
                                onsaveerror(1);
                            }
                            else if (ib_updating)
                            {
                                ib_updating = false;
                                onupdateerror();
                                onsaveerror(2);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        //System.Windows.Forms.MessageBox.Show(e.Message);
                        GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                    }
                }
            }
        }
        public virtual void keydown(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        
        public virtual void lostfocus(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.ItemUID != "")
                    {
                        if (oForm.Items.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            onsetmatrixmenu(is_CurrentMatrixUID, false);
                            is_CurrentMatrixUID = "";
                            if (pVal.ColUID != "")
                                is_CurrentMatrixColUID = "";

                        }
                    }
                }
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                UI.SBO_Application.StatusBar.SetText(e.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            }
        }
        public virtual void matrixcollapsepressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void matrixlinkpressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void matrixload(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void menuclick(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
       
        public virtual void menuevent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            if (!pVal.BeforeAction)
            {
                switch (pVal.MenuUID)
                {
                    case "1281":    //find
                        modechanged();
                        break;
                    case "1282":	//add
                        modechanged();
                        break;
                    case "1283":	//remove
                        onremovesuccess();
                        break;
                    case "1284":	//cancel
                        oncancelsuccess();
                        break;
                    case "1285":	//restore
                        onrestoresuccess();
                        break;
                    case "1286":	//close
                        onclosesuccess();
                        break;
                    case "1288":	// next
                        modechanged();
                        break;
                    case "1289":	// previous
                        modechanged();
                        break;
                    case "1290":	// first
                        modechanged();
                        break;
                    case "1291":	// last
                        modechanged();
                        break;
                    case "1292":
                        if (is_CurrentMatrixUID == "")
                        {
                            SAPbouiCOM.Matrix oMatrix;
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(is_CurrentMatrixUID).Specific;
                            onaddrowsuccess(is_CurrentMatrixUID, oMatrix.RowCount);
                        }
                        else
                        {
                            onaddrowsuccess(is_CurrentMatrixUID, 0);
                        }
                        break;
                    case "1293":
                        ondeleterowsuccess(is_CurrentMatrixUID);
                        break;
                }
                return;
            }
            else
            {
                switch (pVal.MenuUID)
                {
                    case "771":	//cut
                        if (is_CurrentMatrixUID != "")
                            onmatrixeditchanging(is_CurrentMatrixUID, is_CurrentMatrixColUID, il_CurrentMatrixRow, false, ref BubbleEvent);
                        break;
                    case "773":	//paste
                        if (is_CurrentMatrixUID != "")
                            onmatrixeditchanging(is_CurrentMatrixUID, is_CurrentMatrixColUID, il_CurrentMatrixRow, false, ref BubbleEvent);
                        break;
                    case "1282":	//add
                        oninsert(ref BubbleEvent);
                        break;
                    case "1283":	//remove
                        onremove(ref BubbleEvent);
                        break;
                    case "1284":	//cancel
                        oncancel(ref BubbleEvent);
                        break;
                    case "1285":	//restore
                        onrestore(ref BubbleEvent);
                        break;
                    case "1286":	//close
                        onclose(ref BubbleEvent);
                        break;
                    case "1292":	// addrow
                        if (ib_udo)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        onaddrow(is_CurrentMatrixUID, ref BubbleEvent, false);
                        break;
                    case "1293":	// deleterow
                        if (ib_udo)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        ondeleterow(is_CurrentMatrixUID, il_CurrentMatrixRow, ref BubbleEvent, false);

                        break;
                    case "1295":	//copy from cell above
                        if (is_CurrentMatrixUID != "")
                            onmatrixeditchanging(is_CurrentMatrixUID, is_CurrentMatrixColUID, il_CurrentMatrixRow, false, ref BubbleEvent);
                        break;
                    case "1296":	//copy from cell below
                        if (is_CurrentMatrixUID != "")
                            onmatrixeditchanging(is_CurrentMatrixUID, is_CurrentMatrixColUID, il_CurrentMatrixRow, false, ref BubbleEvent);
                        break;
                }
                return;
            }
        }
        public virtual void validate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public virtual void onformattach()
        {
        }
        public virtual void onformattach(ref bool BubbleEvent)
        {
        }
        public virtual void onmatrixeditchanging(string MatrixUID, string coluid, int row, bool istabkey, ref bool BubbleEvent)
        {
        }
        public virtual void onGetCreationParams(ref SAPbouiCOM.BoFormBorderStyle io_BorderStyle, ref string is_FormType, ref string is_ObjectType, ref string xmlPath)
        {
        }
        public virtual void onFormCreate(ref bool ab_visible, ref bool ab_center)
        {
        }

        //on
        public virtual void oninsert(ref bool BubbleEvent)
        {
        }
        public virtual void onremove(ref bool BubbleEvent)
        {
        }
        public virtual void oncancel(ref bool BubbleEvent)
        {
        }
        public virtual void onrestore(ref bool BubbleEvent)
        {
        }
        public virtual void onclose(ref bool BubbleEvent)
        {
        }
        public virtual void onupdate(ref bool BubbleEvent)
        {
        }
        public virtual void onfind(ref bool BubbleEvent)
        {
        }
        public virtual void onok(ref bool BubbleEvent)
        {
        }
        public virtual void onadd(ref bool BubbleEvent)
        {
        }
        public virtual void onsave(int ai_mode, ref bool BubbleEvent)
        {
        }
        public virtual void onsetmatrixmenu(string matrixuid, bool focused)
        {
        }
        public virtual void onaddrow(string matrixuid, ref bool BubbleEvent, bool innerevent)
        {
        }
        public virtual void ondeleterow(string matrixuid, int row, ref bool BubbleEvent, bool innerevent)
        {
        }
        public virtual void onsetrequestor()
        {
        }
        public virtual void onloaddata()
        {
        }
        public virtual void ondisplaydata()
        {
        }
        

        //success
        public virtual void onaddsuccess()
        {
        }
        public virtual void onaddrowsuccess(string matrixuid, int matrixrow)
        {
        }
        public virtual void ondeleterowsuccess(string matrixuid)
        {
        }
        public virtual void oncancelsuccess()
        {
        }
        public virtual void onclosesuccess()
        {
        }
        public virtual void onupdatesuccess()
        {
        }
        public virtual void onsavesuccess(int ai_mode)
        {
        }
        public virtual void onremovesuccess()
        {
        }
        public virtual void onrestoresuccess()
        {
        }
        public virtual void onsaveerror(int ai_mode)
        {
        }
        public virtual void onupdateerror()
        {
        }
        public virtual void onadderror()
        {
        }
        //mode
        public virtual void onupdatemode()
        {
        }
        public virtual void onfindmode()
        {
        }
        public virtual void onokmode()
        {
        }
        public virtual void onaddmode()
        {
        }

        //functions
        public void attachForm(int ai_index, string FormUID)
        {
            ib_udo = false;
            oForm = UI.SBO_Application.Forms.Item(FormUID);
            globalvar.gdt_Form.Rows.Add(ai_index, FormUID, oForm.TypeEx);
            onformattach();
            modechanged();
        }
        public void attachForm(int ai_index, string FormUID, ref bool BubbleEvent)
        {
            ib_udo = false;
            oForm = UI.SBO_Application.Forms.Item(FormUID);

            globalvar.gdt_Form.Rows.Add(ai_index, FormUID, oForm.TypeEx);
            onformattach(ref BubbleEvent);
            modechanged();
        }
        public void createForm(int ai_index)
        {
            SAPbouiCOM.FormCreationParams creationPackage;
            SAPbouiCOM.BoFormBorderStyle BorderStyle;
            System.Xml.XmlDocument oXML;
            string ls_FormType, ls_ObjectType, FormUID, xmlPath, innerxml, appPath;
            bool lb_visible, lb_center;

            BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            lb_visible = true;
            lb_center = true;
            ls_FormType = "";
            ls_ObjectType = "";
            xmlPath = "";
            onGetCreationParams(ref BorderStyle, ref ls_FormType, ref ls_ObjectType, ref xmlPath);
            try
            {
                FormUID = UI.generateFormUID();
                creationPackage = (SAPbouiCOM.FormCreationParams)UI.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                if (string.IsNullOrEmpty(xmlPath))
                {
                    creationPackage.UniqueID = FormUID;
                    creationPackage.FormType = ls_FormType;

                    if (!string.IsNullOrEmpty(ls_ObjectType))
                        creationPackage.ObjectType = ls_ObjectType;

                    creationPackage.BorderStyle = BorderStyle;
                    oForm = UI.SBO_Application.Forms.AddEx(creationPackage);

                    globalvar.gdt_Form.Rows.Add(ai_index, FormUID, ls_FormType);

                }
                else
                {
                    appPath = System.Windows.Forms.Application.StartupPath;
                    oXML = new System.Xml.XmlDocument();
                    oXML.Load(appPath + "\\" + xmlPath);

                    innerxml = oXML.InnerXml;

                    creationPackage.XmlData = innerxml;
                    creationPackage.UniqueID = FormUID;

                    ls_FormType = oXML.SelectSingleNode("Application/forms/action/form/@FormType").Value;

                    oForm = UI.SBO_Application.Forms.AddEx(creationPackage);
                    globalvar.gdt_Form.Rows.Add(ai_index, FormUID, ls_FormType);
                }

                onFormCreate(ref lb_visible, ref lb_center);

                if (lb_center)
                {
                    centerForm();
                }
                if (lb_visible)
                {
                    oForm.Visible = true;
                }
                modechanged();
                creationPackage = null;
                GC.Collect();
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                GlobalFunction.f_error(e);
            }

        }
        public void setrequestor(Form aoForm, SAPbouiCOM.EditText ao_edittext)
        {
            targetEditText = ao_edittext;
            targetsboform = aoForm;
            onsetrequestor();
        }
        protected void modechanged()
        {
            switch (oForm.Mode)
            {
                case SAPbouiCOM.BoFormMode.fm_FIND_MODE:
                    onfindmode();
                    return;
                case SAPbouiCOM.BoFormMode.fm_ADD_MODE:
                    onaddmode();
                    return;
                case SAPbouiCOM.BoFormMode.fm_OK_MODE:
                    onokmode();
                    return;
                case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE:
                    onupdatemode();
                    return;
            }
        }
        protected void centerForm()
        {
            oForm.Left = (UI.SBO_Application.Desktop.Width / 2) - oForm.Width / 2;
            oForm.Top = ((UI.SBO_Application.Desktop.Height / 2) - oForm.Height / 2) - 60;
        }

        protected SAPbouiCOM.Item createStaticText(int left, int top, int width, int height, string name, string caption, string linkto)
        {
            SAPbouiCOM.StaticText oStaticText;
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;
            oItem.LinkTo = linkto;
            oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
            oStaticText.Caption = caption;
            return oItem;
        }
        protected SAPbouiCOM.Item createEditText(int left, int top, int width, int height, string name, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;
            if (bound)
            {
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                oEditText.DataBind.SetBound(bound, tablename, alias);
            }
            return oItem;
        }
        protected SAPbouiCOM.Item createEditText(int left, int top, int width, int height, string name)
        {
            return createEditText(left, top, width, height, name, false, "", "");
        }
        protected SAPbouiCOM.Item createExtEditText(int left, int top, int width, int height, string name, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;
            if (bound)
            {
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                oEditText.DataBind.SetBound(bound, tablename, alias);
                oEditText.ScrollBars = SAPbouiCOM.BoScrollBars.sb_None;
            }
            return oItem;
        }
        protected SAPbouiCOM.Item createExtEditText(int left, int top, int width, int height, string name)
        {
            return createExtEditText(left, top, width, height, name, false, "", "");
        }
        protected SAPbouiCOM.Item createCheckBox(int left, int top, int width, int height, string name, string valon, string valoff, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.CheckBox oCheckBox;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Width = width;
            oItem.Left = left;
            oItem.Top = top;
            oItem.Height = height;

            oCheckBox = (SAPbouiCOM.CheckBox)oItem.Specific;

            oCheckBox.ValOn = valon;
            oCheckBox.ValOff = valoff;
            if (bound)
            {
                oCheckBox.DataBind.SetBound(bound, tablename, alias);

            }
            return oItem;
        }
        protected SAPbouiCOM.Item createCheckBox(int left, int top, int width, int height, string name, string valon, string valoff)
        {
            return createCheckBox(left, top, width, height, name, valon, valoff, false, "", "");
        }
        protected SAPbouiCOM.Item createCombobox(int left, int top, int width, int height, string name, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.ComboBox oComboBox;
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;

            if (bound)
            {
                oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
                oComboBox.DataBind.SetBound(bound, tablename, alias);

            }
            return oItem;
        }
        protected SAPbouiCOM.Item createCombobox(int left, int top, int width, int height, string name)
        {
            return createCombobox(left, top, width, height, name, false, "", "");
        }
        protected SAPbouiCOM.Item createLinkButton(string name, SAPbouiCOM.Item linkto, SAPbouiCOM.BoLinkedObject linkobject)
        {
            SAPbouiCOM.Item oItem;

            SAPbouiCOM.LinkedButton oLink;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItem.Left = linkto.Left - 16;
            oItem.Top = linkto.Top + 2;
            oItem.Width = 16;
            oItem.Height = 9;
            oItem.LinkTo = linkto.UniqueID;
            oLink = (SAPbouiCOM.LinkedButton)oItem.Specific;
            oLink.LinkedObject = linkobject;

            return oItem;

        }
        protected SAPbouiCOM.Item createButton(int left, int top, int width, int height, string name, string caption)
        {

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Button oButton;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;
            if (caption != "")
            {
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = caption;
            }
            return oItem;
        }
        protected SAPbouiCOM.Item createOptionButton(int left, int top, int width, int height, string name, string valon, string valoff, Boolean bound, String tablename, string alias, string caption)
        {
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.OptionBtn oOptionBtn;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            oItem.Width = width;
            oItem.Left = left;
            oItem.Top = top;
            oItem.Height = height;

            oOptionBtn = (SAPbouiCOM.OptionBtn)oItem.Specific;

            oOptionBtn.ValOn = valon;
            oOptionBtn.ValOff = valoff;
            oOptionBtn.Caption = caption;
            if (bound)
            {
                oOptionBtn.DataBind.SetBound(bound, tablename, alias);

            }
            return oItem;
        }
        protected SAPbouiCOM.Item createRectangle(int left, int top, int width, int height, string name)
        {

            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;


            return oItem;
        }
        protected SAPbouiCOM.Item createMatrix(int left, int top, int width, int height, string name)
        {
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oItem.Left = left;
            oItem.Width = width;
            oItem.Top = top;
            oItem.Height = height;

            return oItem;

        }
        protected SAPbouiCOM.Column createMatrixEditText(string matrix, int width, string name, string caption, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add(name, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = caption;
            oColumn.Width = width;

            if (bound)
            {
                oColumn.DataBind.SetBound(bound, tablename, alias);

            }
            return oColumn;
        }
        protected SAPbouiCOM.Column createMatrixEditText(string matrix, int width, string name, string caption)
        {
            return createMatrixEditText(matrix, width, name, caption, false, "", "");
        }
        protected SAPbouiCOM.Column createMatrixCheckBox(string matrix, int width, string name, string caption, string valon, string valoff, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add(name, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = caption;
            oColumn.Width = width;
            oColumn.ValOn = valon;
            oColumn.ValOff = valoff;
            if (bound)
            {
                oColumn.DataBind.SetBound(bound, tablename, alias);

            }
            return oColumn;
        }
        protected SAPbouiCOM.Column createMatrixCheckBox(string matrix, int width, string name, string caption, string valon, string valoff)
        {
            return createMatrixCheckBox(matrix, width, name, caption, valon, valoff, false, "", "");
        }
        protected SAPbouiCOM.Column createMatrixComboBox(string matrix, int width, string name, string caption, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add(name, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = caption;
            oColumn.Width = width;
            if (bound)
            {
                oColumn.DataBind.SetBound(bound, tablename, alias);

            }
            return oColumn;
        }
        protected SAPbouiCOM.Column createMatrixComboBox(string matrix, int width, string name, string caption)
        {
            return createMatrixComboBox(matrix, width, name, caption, false, "", "");
        }
        protected SAPbouiCOM.Column creatematrixlinkedbutton(string matrix, int width, string name, string caption, SAPbouiCOM.BoLinkedObject oLinkedObject, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;
            SAPbouiCOM.LinkedButton oLinkButton;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add(name, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = caption;
            oColumn.Width = width;
            oLinkButton = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
            oLinkButton.LinkedObject = oLinkedObject;

            if (bound)
            {
                oColumn.DataBind.SetBound(bound, tablename, alias);

            }
            return oColumn;
        }
        protected SAPbouiCOM.Column creatematrixlinkedbutton(string matrix, int width, string name, string caption, SAPbouiCOM.BoLinkedObject oLinkedObject)
        {
            return creatematrixlinkedbutton(matrix, width, name, caption, oLinkedObject, false, "", "");
        }
        protected SAPbouiCOM.Item createFolder(int left, int top, int width, int height, string name, string caption, string groupwith, Boolean bound, String tablename, string alias)
        {
            SAPbouiCOM.Folder oFolder;
            SAPbouiCOM.Item oItem;

            oItem = oForm.Items.Add(name, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oItem.Left = left;
            oItem.Top = top;
            oItem.Height = height;
            oItem.Width = width;

            oFolder = (SAPbouiCOM.Folder)oItem.Specific;
            oFolder.Caption = caption;
            if (groupwith != "")
            {
                oFolder.GroupWith(groupwith);
            }
            if (bound)
            {
                oFolder.DataBind.SetBound(true, tablename, alias);
            }
            return oItem;
        }
        protected string getItemString(string itemUID)
        {
            SAPbouiCOM.EditText oEditText;

            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item(itemUID).Specific;
            return oEditText.String;
        }
        protected string getItemSelectedDesc(string itemUID, string as_default)
        {
            SAPbouiCOM.ComboBox oComboBox;

            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item(itemUID).Specific;
            if (!(oComboBox.Selected == null))
            {
                return oComboBox.Selected.Description;
            }
            else
            {
                return as_default;
            }
        }
        protected string getItemSelectedValue(string itemUID, string as_default)
        {
            SAPbouiCOM.ComboBox oComboBox;

            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item(itemUID).Specific;
            if (!(oComboBox.Selected == null))
            {
                return oComboBox.Selected.Value;
            }
            else
            {
                return as_default;
            }
        }
        protected object getItemSpecific(string itemUID)
        {
            return oForm.Items.Item(itemUID).Specific;
        }

        protected string getGridString(string grid, string ColUID, int row)
        {
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.EditTextColumn oEditText;

            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific;
                oEditText = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(ColUID);
                return oEditText.GetText(row);
            }
             catch (Exception e)
             {
                 GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                 //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                 return "";
             }
        }
        protected string getGridSelectedValue(string grid, string ColUID, int row)
        {
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.ComboBoxColumn oComboBox;

            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific;
                oComboBox = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(ColUID);
                return oComboBox.GetSelectedValue(row).Value;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                return "";
            }
        }
        protected string getGridSelectedDesc(string grid, string ColUID, int row)
        {
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.ComboBoxColumn oComboBox;

            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific;
                oComboBox = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(ColUID);
                return oComboBox.GetSelectedValue(row).Description;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                return "";
            }
        }
        protected string getColumnString(string matrix, string ColUID, int row, string as_string)
        {
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
                return oEditText.String;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                return "";
            }
        }
        protected string getColumnSelectedDesc(string matrix, string ColUID, int row, string as_string)
        {
            SAPbouiCOM.ComboBox oComboBox;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
                return oComboBox.Selected.Description;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                return "";
            }
        }
        protected string getColumnSelectedValue(string matrix, string ColUID, int row, string as_string)
        {
            SAPbouiCOM.ComboBox oComboBox;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
                return oComboBox.Selected.Value;

            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
                return "";
            }
        }
        protected object getColumnSpecific(string matrix,string ColUID,int row)
        {
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = (SAPbouiCOM.Matrix) oForm.Items.Item(matrix).Specific;
            return oMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
        }
        protected void setColumnString(string matrix, string ColUID, int row, string as_string)
        {
            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Matrix OMatrix;

            try
            {
                OMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oEditText = (SAPbouiCOM.EditText)OMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
                oEditText.String = as_string;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
            }
        }
        protected void setColumnSelectedValue(string matrix, string ColUID, int row, string as_string)
        {
            SAPbouiCOM.ComboBox oComboBox;
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(ColUID).Cells.Item(row).Specific;
                oComboBox.Select(as_string, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
            }
        }
        protected void setColumnEnabled(string matrix, string ColUID, bool enabled)
        {
            SAPbouiCOM.Matrix oMatrix;

            try
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific;
                oMatrix.Columns.Item(ColUID).Editable = enabled;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
            }
        }
        protected void setItemSelectedValue(string itemUID, string as_string)
        {
            SAPbouiCOM.ComboBox oComboBox;

            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item(itemUID).Specific;
            oComboBox.Select(as_string, SAPbouiCOM.BoSearchKey.psk_ByValue);
        }
        protected void setItemString(string itemUID, string as_string)
        {
            SAPbouiCOM.EditText oEditText;
            try
            {

                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item(itemUID).Specific;
                oEditText.String = as_string;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                //UI.SBO_Application.MessageBox(e.Message, 1, "Ok", "", "");
            }
        }
        protected void setStaticTextCaption(string itemUID, string as_string)
        {
            SAPbouiCOM.StaticText oStaticText;
            try
            {
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item(itemUID).Specific;
                oStaticText.Caption = as_string;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
            }
        }
        protected void setButtonCaption(string itemUID, string as_string)
        {
            SAPbouiCOM.Button oButton;
            try
            {
                oButton = (SAPbouiCOM.Button)oForm.Items.Item(itemUID).Specific;
                oButton.Caption = as_string;
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
            }
        }
        protected void setItemEnabled(string itemUID, bool enabled)
        {
            oForm.Items.Item(itemUID).Enabled = enabled;
        }
        protected void setItemVisible(string itemUID, bool visible)
        {
            oForm.Items.Item(itemUID).Visible = visible;
        }

        public virtual void setoverlapform()
        {
            overlapForm = null;
            overlap =  false;
        }
        public virtual void setoverlapform(Form ao_sboform)
        {
            overlapForm = ao_sboform;
            overlap = true;
        }
        protected void queryopen(int ai_idx, string as_sql)
        {
            if (globalvar.gb_Recordset.GetUpperBound(0) < ai_idx)
            {
                globalvar.gb_Recordset[ai_idx] = true;
                globalvar.go_Recordset[ai_idx] = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            }
            else
            {
                if (!globalvar.gb_Recordset[ai_idx])
                {
                    globalvar.gb_Recordset[ai_idx] = true;
                    globalvar.go_Recordset[ai_idx] = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                }
            }
            globalvar.go_Recordset[ai_idx].DoQuery(as_sql);
        }
        protected void queryfetch(int ai_idx, string as_action)
        {
            switch (as_action)
            {
                case "next":
                    globalvar.go_Recordset[ai_idx].MoveNext();
                    break;
                case "first":
                    globalvar.go_Recordset[ai_idx].MoveFirst();
                    break;
                case "last":
                    globalvar.go_Recordset[ai_idx].MoveLast();
                    break;
                case "previous":
                    globalvar.go_Recordset[ai_idx].MovePrevious();
                    break;
                default:
                    UI.SBO_Application.MessageBox("Invalid Query Action [" + as_action + "].", 1, "Ok", "", "");
                    break;
            }
        }

        public virtual void poponsave(ref bool BubbleEvent)
        {
        }
        protected void itemclick(string as_fldname)
        {
            oForm.Items.Item(as_fldname).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        }
        protected int queryrows(int ai_idx)
        {
            return globalvar.go_Recordset[ai_idx].RecordCount;
        }
        protected object queryvalue(int ai_idx, int ai_fldidx)
        {
            return globalvar.go_Recordset[ai_idx].Fields.Item(ai_fldidx).Value;
        }
        protected object queryvalue(int ai_idx, string as_fldname)
        {
            return globalvar.go_Recordset[ai_idx].Fields.Item(as_fldname).Value;
        }
        protected void queryclose(int ai_idx)
        {
            if (globalvar.gb_Recordset.GetUpperBound(0) < ai_idx)
            {
                UI.SBO_Application.MessageBox("Invalid Query Object [" + ai_idx.ToString() + "] to close.", 1, "Ok", "", "");
            }
            else
            {
                if (globalvar.gb_Recordset[ai_idx])
                {
                    globalvar.gb_Recordset[ai_idx] = false;
                    globalvar.go_Recordset[ai_idx] = null;
                    GC.Collect();
                }
                else
                {
                    //UI.SBO_Application.MessageBox("Query Object [" + ai_idx.ToString() + "] already close.", 1, "Ok", "", "");
                }
            }
        }
    }
}
