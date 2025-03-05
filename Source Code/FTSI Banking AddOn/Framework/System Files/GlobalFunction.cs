using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Data.OleDb;

//********************************************************************************************
// DATE CREATED : December 2008
// REMARKS      : JOHN WILSON DE LOS SANTOS ( PROGRAMMER )
// CLASS NAME   : GlobalFunction.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class GlobalFunction : UserControl
    {

        public static string f_convert_date_sbodate(DateTime adt_datetime)
        {
            string ls_datetime;

            if (adt_datetime.Day.ToString().Length == 1)
            {
                if (adt_datetime.Month.ToString().Length == 1)
                {
                    ls_datetime = adt_datetime.Year.ToString() + "0" + adt_datetime.Month.ToString() + "0" + adt_datetime.Day.ToString();
                }
                else
                {
                    ls_datetime = adt_datetime.Year.ToString() + adt_datetime.Month.ToString() + "0" + adt_datetime.Day.ToString();
                }
            }
            else
            {
                if (adt_datetime.Month.ToString().Length == 1)
                {
                    ls_datetime = adt_datetime.Year.ToString() + "0" + adt_datetime.Month.ToString() + adt_datetime.Day.ToString();
                }
                else
                {
                    ls_datetime = adt_datetime.Year.ToString() + adt_datetime.Month.ToString() + adt_datetime.Day.ToString();
                }
            }
            return ls_datetime;
        }

        public static string f_convertsbodate_tostring(string as_stringdate)
        {
            string ls_date;

            ls_date = as_stringdate.Substring(4, 2) + "/" + as_stringdate.Substring(6, 2) + "/" + as_stringdate.Substring(0, 4);

            return ls_date;
        }
        
        
        public static string f_convert_date_tostring(DateTime adt_datetime)
        {
            string ls_datetime;

            if (adt_datetime.Day.ToString().Length == 1)
            {
                ls_datetime = adt_datetime.Month.ToString() + "/0" + adt_datetime.Day.ToString() + "/" + adt_datetime.Year.ToString();
            }
            else
            {
                ls_datetime = adt_datetime.Month.ToString() + "/" + adt_datetime.Day.ToString() + "/" + adt_datetime.Year.ToString();
            }
            return ls_datetime;
        }
        public static string f_convert_string_to_date(string as_stringdate)
        {
            string ls_year, ls_month, ls_day;
            DateTime ldt_datetime;

            try
            {
                ls_year = as_stringdate.Substring(0, 4);
                ls_month = as_stringdate.Substring(4, 2);
                ls_day = as_stringdate.Substring(6, 2);

                ldt_datetime = System.Convert.ToDateTime(ls_month + "/" + ls_day + "/" + ls_year);

                return ldt_datetime.ToString();
            }
            catch (Exception e)
            {
                GlobalFunction.fileappend("Errmsg -" + e.Message + ".");
                return "";
            }
        }
        public static Int16 f_intcolumn(string as_column)
        {
            switch (as_column.ToUpper())
            {
                case "A":
                    return 1;
                case "B":
                    return 2;
                case "C":
                    return 3;
                case "D":
                    return 4;
                case "E":
                    return 5;
                case "F":
                    return 6;
                case "G":
                    return 7;
                case "H":
                    return 8;
                case "I":
                    return 9;
                case "J":
                    return 10;
                case "K":
                    return 11;
                case "L":
                    return 12;
                case "M":
                    return 13;
                case "N":
                    return 14;
                case "O":
                    return 15;
                case "P":
                    return 16;
                case "Q":
                    return 17;
                case "R":
                    return 18;
                case "S":
                    return 19;
                case "T":
                    return 20;
                case "U":
                    return 21;
                case "V":
                    return 22;
                case "W":
                    return 23;
                case "X":
                    return 24;
                case "Y":
                    return 25;
                case "Z":
                    return 26;
                case "AA":
                    return 27;
                case "AB":
                    return 28;
                case "AC":
                    return 29;
                case "AD":
                    return 30;
                case "AE":
                    return 31;
                case "AF":
                    return 32;
                case "AG":
                    return 33;
                case "AH":
                    return 34;
                case "AI":
                    return 35;
                case "AJ":
                    return 36;
                case "AK":
                    return 37;
                case "AL":
                    return 38;
                case "AM":
                    return 39;
                case "AN":
                    return 40;
                case "AO":
                    return 41;
                case "AP":
                    return 42;
                case "AQ":
                    return 43;
                case "AR":
                    return 44;
                case "AS":
                    return 45;
                case "AT":
                    return 46;
                case "AU":
                    return 47;
                case "AV":
                    return 48;
                case "AW":
                    return 49;
                case "AX":
                    return 50;
                case "AY":
                    return 51;
                case "AZ":
                    return 52;
                default:
                    return 0;
            }
        }
        public static string f_getCellValue(Microsoft.Office.Interop.Excel.Worksheet excelWorksheet, long al_row, string as_column)
        {
            long ll_column;
            Microsoft.Office.Interop.Excel.Range excelCell;

            ll_column = f_intcolumn(as_column);

            excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.Cells[al_row, ll_column];

            return excelCell.Text.ToString();
        }
        public static Double f_convertToDouble(string as_data)
        {
            if (string.IsNullOrEmpty(as_data))
            {
                return 0;
            }
            else
            {
                return System.Convert.ToDouble(as_data);
            }
        }
        public static bool f_checkexist(string TableName, string ls_where)
        {
            SAPbobsCOM.Recordset oRecordset;
            oRecordset = (SAPbobsCOM.Recordset)DI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("Select * from [" + TableName + "] where " + ls_where);
            if (oRecordset.RecordCount > 0)
            {
                oRecordset = null;
                return true;
            }
            else
            {
                oRecordset = null;
                return false;
            }
        }
        public static void f_error(Exception e)
        {
            string errmsg, ls_error;
            int errorno;

            errorno = DI.oCompany.GetLastErrorCode();
            errmsg = DI.oCompany.GetLastErrorDescription();
            if (string.IsNullOrEmpty(errmsg))
                errmsg = e.Message;

            ls_error = "Error Code        : " + errorno.ToString() + "\n\n" +
                       "Error Description : " + errmsg; // + "\n\n" +
            //"This Addon/Application will terminate.\n\n" +
            //"NOTE: All Addon/Application functionality will not be available anymore, contact system administrator.";

            UI.SBO_Application.MessageBox(ls_error, 1, "Ok", "", "");
            fileappend(ls_error);
        }
        public static void filewrite()
        {
            globalvar.filepath = System.Windows.Forms.Application.StartupPath + "\\ERROR LOG\\" + DateTime.Now.ToString("dd-MM-yy") + "- Error.txt";
            if (!File.Exists(globalvar.filepath))
            {
                FileInfo errorlog = new FileInfo(globalvar.filepath);
                StreamWriter streamwriter = errorlog.CreateText();
                streamwriter.WriteLine(DateTime.Today.ToString() + "          SAP Business One AddOn Error Log");
                streamwriter.Close();
            }
        }
        public static void fileappend(string text)
        {
            FileInfo errorlog = new FileInfo(globalvar.filepath);
            StreamWriter streamwriter = errorlog.AppendText();
            streamwriter.WriteLine(DateTime.Now.TimeOfDay.ToString() + "          "+ text);
            streamwriter.Close();
        }
        public static void showOpenFileDialog(SAPbouiCOM.Item oItem, string as_filter)
        {
            try
            {
                globalvar.g_Item = oItem;
                globalvar.gs_filter = as_filter;
                System.Threading.Thread ShowFileBrowserThread;
                ShowFileBrowserThread = new System.Threading.Thread(ShowFileBrowser);
                if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFileBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFileBrowserThread.Start();
                }
                else if (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFileBrowserThread.Start();
                    ShowFileBrowserThread.Join();
                }
                while (ShowFileBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception ex)
            {

            }
        }
        public static void ShowFileBrowser()
        {
            Process[] MyProcs = new Process[Process.GetProcessesByName("SAP Business One").Length];
            globalvar.FileName = "";
            OpenFileDialog OpenFile = new OpenFileDialog();
            try
            {
                OpenFile.Multiselect = false;
                OpenFile.Filter = globalvar.gs_filter;
                Int32 filterindex = 0;
                try
                {
                    filterindex = 0;
                }
                catch (Exception e)
                {
                }
                OpenFile.FilterIndex = filterindex;
                OpenFile.RestoreDirectory = true;
                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (Int32 i = 0; i < MyProcs.Length; i++)
                {
                    if (MyProcs[i].MainWindowTitle == UI.SBO_Application.Desktop.Title)
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        DialogResult ret = OpenFile.ShowDialog(MyWindow);
                        if (ret == DialogResult.OK)
                        {
                            SAPbouiCOM.EditText oEditText;
                            oEditText = (SAPbouiCOM.EditText)globalvar.g_Item.Specific;
                            oEditText.String = OpenFile.FileName;
                            oEditText = null;
                            globalvar.g_Item = null;
                            OpenFile.Dispose();
                        }
                        else
                        {
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                globalvar.FileName = "";
            }
            finally
            {
                OpenFile.Dispose();
            }
        }
        public static void showFolderBrowserDialog(SAPbouiCOM.Item oItem)
        {
            try
            {
                globalvar.g_Item = oItem;
                System.Threading.Thread ShowFolderBrowserThread;
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception ex)
            {

            }
        }
        public static void ShowFolderBrowser()
        {
            Process[] MyProcs = new Process[Process.GetProcessesByName("SAP Business One").Length];
            globalvar.FileName = "";
            FolderBrowserDialog OpenFolder = new FolderBrowserDialog();
            try
            {
                OpenFolder.Description = "Select the directory that you want to use as the default.";

                // Do not allow the user to create new files via the FolderBrowserDialog. 
                OpenFolder.ShowNewFolderButton = false;

                // Default to the My Documents folder. 
                OpenFolder.RootFolder = Environment.SpecialFolder.Personal;


                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (Int32 i = 0; i < MyProcs.Length; i++)
                {
                    if (MyProcs[i].MainWindowTitle == UI.SBO_Application.Desktop.Title)
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        DialogResult result = OpenFolder.ShowDialog(MyWindow);
                        if (result == DialogResult.OK)
                        {
                            string folderName = OpenFolder.SelectedPath;

                            SAPbouiCOM.EditText oEditText;
                            oEditText = (SAPbouiCOM.EditText)globalvar.g_Item.Specific;
                            oEditText.String = folderName;
                            oEditText = null;
                            globalvar.g_Item = null;
                        }
                        else
                        {
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                globalvar.FileName = "";
            }
            finally
            {
                OpenFolder.Dispose();
            }
        }
        public static bool importXLSX(string strXLSPath, string strHeader, string strSheet)
        {
            try
            {
                globalvar.oDTImpData.Clear();

                string connString = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 8.0; HDR={1};' ", strXLSPath, strHeader);
                
                OleDbConnection oledbConn = new OleDbConnection(connString);
                    
                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}$]", strSheet), oledbConn);

                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                oleda.Fill(globalvar.oDTImpData);

                oledbConn.Close();

                return true;

            }
            catch (Exception ex)
            {

                GlobalFunction.fileappend("Errmsg -" + ex.Message + ".");
                return false;
            }

        }
    }
}
