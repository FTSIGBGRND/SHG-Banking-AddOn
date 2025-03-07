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
// CLASS NAME   : nv_string.cs
// VERSION      : Version 2.0
// NOTE         : THIS CODE AND INFORMATION IS PROVIDED 'AS IS' WITHOUT WARRANTY OF
//                ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
//                THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//                PARTICULAR PURPOSE.
//********************************************************************************************

namespace AddOn
{
    public partial class nv_string : UserControl
    {
        public static string of_getToken(ref string as_source, string as_separator)
        {
            string ls_ret;
            int li_pos;

            if (as_source == null || as_separator == "")
            {
                string ls_null;
                ls_null = null;
                return ls_null;
            }

            li_pos = as_source.IndexOf(as_separator);
            if (li_pos == -1)
            {
                ls_ret = as_source;
                as_source = "";
            }
            else
            {
                ls_ret = as_source.Substring(0, li_pos);
                as_source = as_source.Substring(li_pos + 1);
            }

            return ls_ret;


        }
        public static string encrypt(string as_string)
        {
            int j, mod;

            string ls_enctext = "", ls_temp;
            const string CRYPT_KEY = "FarEastExpressIncorporated";

            j = as_string.Length;
            for (int i = 0; i < j; i++)
            {
                Math.DivRem(i, 10, out mod);
                ls_temp = CRYPT_KEY.Substring(mod + 1, 1);
                if (ls_temp == "‘" || ls_temp == "’" || ls_temp == "'" || ls_temp == "	" || ls_temp == '"'.ToString())
                {
                    ls_temp = i.ToString();
                }
                ls_enctext = ls_enctext + ls_temp;

                ls_temp = Convert.ToString(255 - Convert.ToInt32(Convert.ToChar(as_string.Substring(i, 1))));
                if (ls_temp == "‘" || ls_temp == "’" || ls_temp == "'" || ls_temp == "	" || ls_temp == '"'.ToString())
                {
                    ls_temp = i.ToString();
                }
                ls_enctext = ls_enctext + ls_temp;

            }
            //ls_enctext = ls_enctext.Replace("", "?");
            return ls_enctext;
        }
    }
}
