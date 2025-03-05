using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddOn
{
    public partial class FormCFL : AddOn.Form
    {
        protected DataTable oDataTable;
        public FormCFL()
        {
            InitializeComponent();
            
            System.Data.DataColumn[] key = new System.Data.DataColumn[1];
            oDataTable = new System.Data.DataTable("Items");
            
            key[0] = oDataTable.Columns.Add("Row", typeof(System.Int16));
            oDataTable.Columns.Add("Str1", typeof(System.String));
            oDataTable.Columns.Add("Str2", typeof(System.String));
            oDataTable.Columns.Add("Str3", typeof(System.String));
            oDataTable.Columns.Add("Str4", typeof(System.String));
            oDataTable.Columns.Add("Str5", typeof(System.String));
            oDataTable.Columns.Add("Str6", typeof(System.String));
            oDataTable.Columns.Add("Str7", typeof(System.String));
            oDataTable.PrimaryKey = key;
        }
        public SAPbouiCOM.Matrix oMatrix;
        public string is_findcolumn = "str1";
        public string is_findcolumntype = "string";
        public string is_sourcecolumn = "";
        public override void click(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.click(FormUID, ref pVal, ref BubbleEvent);
            if(!pVal.BeforeAction)
            {
                switch(pVal.ItemUID)
                {
                    case "grdItems":
                        if(pVal.Row > 0 )
                        {
                            if(pVal.Row <= oMatrix.RowCount)
                            {
                                il_CurrentMatrixRow = pVal.Row;
                            }
                            oMatrix.SelectRow(pVal.Row,true,false);
                        }
                        break;
                }
            }
        }
        public override void doubleclick(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.doubleclick(FormUID, ref pVal, ref BubbleEvent);
            if(!pVal.BeforeAction)
            {
                switch(pVal.ItemUID)
                {
                    case "grdItems":
                        string ls_select;

                        ls_select = getColumnString("grdItems",is_sourcecolumn,il_CurrentMatrixRow,"");
                        unsetoverlapform();
                        targetEditText.String = ls_select;
                        oDataTable.Clear();
                        UI.SBO_Application.ActivateMenuItem("514");
                        break;
                }
            }
		
        }
        public override void itempressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.itempressed(FormUID, ref pVal, ref BubbleEvent);
            if (!pVal.BeforeAction)
            {
                if (pVal.ItemUID == "Choose")
                {
                    string ls_select;
                    ls_select = getColumnString("grdItems", is_sourcecolumn, il_CurrentMatrixRow, "");
                    unsetoverlapform();
                    targetEditText.String = ls_select;
                    oDataTable.Clear();
                    UI.SBO_Application.ActivateMenuItem("514");
                }
            }

            //    elseif pval.itemuid = "grdItems" then
            //        if pval.row = 0 then
            //            sortdata(event onheaderclick(pval.coluid))
            //            event ondisplaydata()
            //            updateselection()
            //        end if	
            //    end if
            //end if	
        }
        public override void keydown(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            base.keydown(FormUID, ref pVal, ref BubbleEvent);
            if (!pVal.BeforeAction)
            {
                switch (pVal.ItemUID)
                {
                    case "U_Find":
                        switch(pVal.CharPressed)
                        {
                            case 40:  //arrowdown
                                if (il_CurrentMatrixRow < oMatrix.RowCount)
                                    focusrow(il_CurrentMatrixRow + 1);
                                break;
                            case 38:
                                if (il_CurrentMatrixRow >1)
                                    focusrow(il_CurrentMatrixRow - 1);
                                break;
                            case 34:
                                if (il_CurrentMatrixRow + 9 <= oMatrix.RowCount)
                                    focusrow(il_CurrentMatrixRow + 9);
                                else
                                    focusrow(oMatrix.RowCount);
                                break;
                            case 33:
                                if (il_CurrentMatrixRow - 9 > 1)
                                    focusrow(il_CurrentMatrixRow - 9);
                                else
                                    focusrow(1);
                                break;
                            default:
                                string ls_find;
                                DataRow[] oDataRow;

                                ls_find = getItemString("U_Find").ToUpper();
                                //if(is_findcolumntype == "string")
                                //{
                                
                                    oDataRow =  oDataTable.Select(is_findcolumn +" like '" + ls_find + "%'");
                                //}
                                //else
                                //{
                                //    oDataRow = globalvar.gdt_Form.Select("string("is_findcolumn + ") like " + ls_find );
                                //}
                                
                                if (oDataRow.Length > 0)
                                {
                                    focusrow(Convert.ToInt16(oDataRow[0]["Row"]) + 1);
                                        //globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].setoverlapform();

                                }
                                    
                                break;
//                                    string ls_find, ls_search	
//                long ll_ctr
//            ls_find = Upper(iole_form.Items.Item("edtFind").Specific.String)
//                if is_findcolumntype = "string" then
//                    ll_ctr = ids_cfl.find(is_findcolumn + " like '" + ls_find + "%'",1,ids_cfl.rowcount())
//                elseif is_findcolumntype = "number" then
////					ll_ctr = ids_cfl.find(is_findcolumn + " = " + ls_find,1,ids_cfl.rowcount())
//                    ll_ctr = ids_cfl.find("string(" + is_findcolumn + ") like '" + ls_find + "%'",1,ids_cfl.rowcount())
//                end if	
//                if ll_ctr > 0 then
//                    focusrow(ll_ctr)
//            end if    
                        }
                        break;
                }
            }

        }

        

        protected void focusrow(int ii_row)
        {
            	oMatrix.SelectRow(ii_row,true,false);
                il_CurrentMatrixRow = ii_row;
        }
        protected void generatedefaultform(string as_desc, bool ab_btnnew)
        {
            SAPbouiCOM.Item oItem;

            oForm.Title = as_desc;
            //oForm.Left = 350;
            oForm.Width = 443;
            //oForm.Top = 50;
            oForm.Height = 289;

            
            //oItem = createStaticText(6, 6, 400, 14, "stlist", as_desc, "");
            //oItem.FontSize = 15;
            //oItem.TextStyle = 1;
            oItem = createStaticText(6, 6, 74, 14, "stFindt", "Find", "U_Find");
            oItem = createEditText(80, 6, 161, 14, "U_Find");
            oItem = createButton(6, 230, 66, 19, "Choose", "Choose");
            oItem = createButton(77, 230, 66, 19, "2", "");
            if (ab_btnnew)
            {
                oItem = createButton(77, 233, 66, 19, "New", "&New");
            }
            oItem = createMatrix(6, 26, 428, 200, "grdItems");

            oMatrix = (SAPbouiCOM.Matrix) oItem.Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

        }
        protected void setfindcolumn(string as_findcolumn, string as_findcolumntype)
        {
            is_findcolumn = as_findcolumn;
            is_findcolumntype = as_findcolumntype;
        }
        protected void setsourcecolumn(string as_column)
        {
            is_sourcecolumn = as_column;
        }
        private void unsetoverlapform()
        {
            targetsboform.setoverlapform();
            //DataRow[] oDataRow;
            //oDataRow = globalvar.gdt_Form.Select("FormUID = '" + targetsboform.UniqueID + "'");
            //if (oDataRow.Length > 0)
            //{
            //    try
            //    {
            //        globalvar.sboform[System.Convert.ToInt16(oDataRow[0]["Index"])].setoverlapform();
            //    }
            //    catch (Exception e)
            //    {
            //    }
            //}
        }
        
       
    }
}

