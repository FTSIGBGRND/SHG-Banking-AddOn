using System;
using System.Collections.Generic;
using System.Windows.Forms;

//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : Program.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string sConnectionString = null;

            GlobalFunction.filewrite();

            System.Data.DataColumn[] key = new System.Data.DataColumn[1];
            globalvar.gdt_Form = new System.Data.DataTable("FORMS");
            key[0] = globalvar.gdt_Form.Columns.Add("Index", typeof(System.Int16));
            globalvar.gdt_Form.Columns.Add("FormUID", typeof(System.String));
            globalvar.gdt_Form.Columns.Add("FormTypeEx", typeof(System.String));

            globalvar.gdt_Form.PrimaryKey = key;

            globalvar.sboform = new Form[100];
            globalvar.go_Recordset = new SAPbobsCOM.Recordset[100];
            globalvar.gb_Recordset = new bool[100];

            if (Environment.GetCommandLineArgs().Length > 1) sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            else sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            addon oAddon = new addon(sConnectionString);
            System.Windows.Forms.Application.Run();
        }
    }
}