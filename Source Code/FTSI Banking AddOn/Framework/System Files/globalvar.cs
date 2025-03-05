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
// CLASS NAME   : globalvar.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class globalvar : UserControl
    {
        public static DataTable gdt_Form;
        public static Form[] sboform;
        public static long FormCount;
        public static bool gb_installedUDO;
        public static SAPbouiCOM.Form CurrentForm;
        public static bool itemchanged;
        public static string userid;
        public static string filepath;
        public static string addondescription;
        public static SAPbobsCOM.Recordset[] go_Recordset;
        public static bool[] gb_Recordset;
        public static bool fromevent, msgresult, fromexit, withpopup;
        public static string reason;
        public static DateTime closedate;
        public static string FileName, gs_filter;
        public static SAPbouiCOM.Item g_Item;

        public static bool bl_addon = false;
            
        //public static SAPbobsCOM.Recordset Recordset;

        public static DataTable oDataTable;
        public static DataTable oAmtFin;
        public static string DelNo;
        public static string DocKey;

        public static string strDrftNum;

        public static DataTable oDTImpData = new DataTable("ImportData");
    }
}
