using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace FTSIBankingAddOn
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

            AddOn.GlobalFunction.filewrite();

            System.Data.DataColumn[] key = new System.Data.DataColumn[1];
            AddOn.globalvar.gdt_Form = new System.Data.DataTable("FORMS");
            key[0] = AddOn.globalvar.gdt_Form.Columns.Add("Index", typeof(System.Int16));
            AddOn.globalvar.gdt_Form.Columns.Add("FormUID", typeof(System.String));
            AddOn.globalvar.gdt_Form.Columns.Add("FormTypeEx", typeof(System.String));

            AddOn.globalvar.gdt_Form.PrimaryKey = key;

            AddOn.globalvar.sboform = new AddOn.Form[100];
            AddOn.globalvar.go_Recordset = new SAPbobsCOM.Recordset[100];
            AddOn.globalvar.gb_Recordset = new bool[100];

            if (Environment.GetCommandLineArgs().Length > 1) sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            else sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            AddOn.addon oAddon = new AddOn.addon(sConnectionString);
            System.Windows.Forms.Application.Run();
        }
    }
}
