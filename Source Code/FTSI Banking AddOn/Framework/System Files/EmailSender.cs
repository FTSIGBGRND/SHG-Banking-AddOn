using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;

namespace AddOn
{
    public partial class EmailSender : UserControl
    {
        private static string strSMTPHost, strEmailUserName, strEmailPassword, 
                              strEmailSubject, strEmailTo, strEmailCC;

        private static int intEmailPort;
        public EmailSender()
        {
            InitializeComponent();
        }
        public static bool getSMTPCredentials(string strPathConnect)
        {
            string[] strLines;

            try
            {
                strLines = File.ReadAllLines(strPathConnect);

                strSMTPHost = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                intEmailPort = Convert.ToInt32(strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1));
                strEmailUserName = strLines[2].ToString().Substring(strLines[2].IndexOf("=") + 1);
                strEmailPassword = strLines[3].ToString().Substring(strLines[3].IndexOf("=") + 1);
                strEmailTo = strLines[4].ToString().Substring(strLines[4].IndexOf("=") + 1);
                strEmailCC = strLines[5].ToString().Substring(strLines[5].IndexOf("=") + 1);
                strEmailSubject = strLines[6].ToString().Substring(strLines[6].IndexOf("=") + 1);

            }
            catch (Exception ex)
            {
                GlobalFunction.fileappend("Errmsg -" + ex.Message + ".");
                return false;
            }
            return true;
        }
        public static bool sendSMTPEmail(string strSubject, string strMailTo, string strMailCC, string strBody)
        {
            string[] strAEMailTo, strAEMailCC;

            string strAttPath = System.Windows.Forms.Application.StartupPath + "\\E-Mail Settings\\Attachment\\";

            string strSMTPSettings = System.Windows.Forms.Application.StartupPath + "\\E-Mail Settings\\E-Mail Connect Settings.ini";

            try
            {
                
                if (getSMTPCredentials(strSMTPSettings))
                {
                    strEmailTo = strEmailTo + strMailTo;
                    strEmailCC = strEmailCC + strMailCC;

                    strAEMailTo = strEmailTo.Split(Convert.ToChar(";"));
                    strAEMailCC = strEmailCC.Split(Convert.ToChar(";"));

                    MailMessage emailmsg = new MailMessage();
                    SmtpClient smtpServer = new SmtpClient(strSMTPHost, intEmailPort);

                    emailmsg.From = new MailAddress(strEmailUserName);

                    for (int intTo = 0; intTo < strAEMailTo.Length; intTo++)
                    {
                        if (!string.IsNullOrEmpty(strAEMailTo[intTo].Trim()))
                            emailmsg.To.Add(strAEMailTo[intTo].Trim());
                    }

                    for (int intCC = 0; intCC < strAEMailCC.Length; intCC++)
                    {
                        if (!string.IsNullOrEmpty(strAEMailCC[intCC].Trim()))
                            emailmsg.CC.Add(strAEMailCC[intCC].Trim());
                    }

                    emailmsg.Subject = strEmailSubject + strSubject;

                    emailmsg.Body = strBody;

                    smtpServer.EnableSsl = true;

                    foreach (var strFile in Directory.GetFiles(strAttPath, "*.*"))
                    {
                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(strFile);
                        emailmsg.Attachments.Add(attachment);
                    }

                    smtpServer.Credentials = new System.Net.NetworkCredential(strEmailUserName, strEmailPassword);
                    smtpServer.ServicePoint.MaxIdleTime = 2;
                    smtpServer.Send(emailmsg);

                    return true;

                }
                else
                {

                    GlobalFunction.fileappend("Errmsg - Please check E-Mail Connection Settings.");
                    return false;
                }
            }
            catch (Exception ex)
            {

                GlobalFunction.fileappend("Errmsg -" + ex.Message + ".");
                return false;
            }
        }

    }
}
