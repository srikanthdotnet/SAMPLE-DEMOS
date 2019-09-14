using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Data;
using System.Text;
using Microsoft.VisualBasic;
using System.IO;
using System.Configuration;
using System.Net.Mail;
using System.Collections;
using System.Security.Cryptography;
namespace Kloc.Models
{
    public class Util
    {
        #region "GetErrorCode"
        public static DataSet GetErrorcode(string Errorcode, string Errmsg)
        {
            DataTable dtException = new DataTable();
            dtException.Columns.Add("Error_Code");
            dtException.Columns.Add("Error_Message");
            DataRow dr = default(DataRow);
            dr = dtException.NewRow();
            dr["Error_Code"] = Errorcode;
            dr["Error_Message"] = Errmsg;
            dtException.Rows.Add(dr);
            DataSet DsException = new DataSet();
            DsException.Tables.Add(dtException);
            DsException.Tables[0].TableName = "Table";
            return DsException;
        }
        #endregion

        public static void WriteLog(string Message, string Source, string StackTrace,string metodname)
        {

            string path = System.AppDomain.CurrentDomain.BaseDirectory + "\\LogFile\\Error" + Strings.Format(DateTime.Now, "MMMyyyy") + ".log";
            if (!Directory.Exists(System.AppDomain.CurrentDomain.BaseDirectory + "\\LogFile"))
            {
                Directory.CreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory + "\\LogFile");
            }
            if (!File.Exists(path))
            {
                File.Create(path).Dispose();
            }
            using (StreamWriter errFile = System.IO.File.AppendText(path))
            {
                StringBuilder sbError = new StringBuilder("");
                sbError.Append("MethodName:" + Constants.vbTab + metodname + Constants.vbNewLine);
                sbError.Append("Err Date:" + Constants.vbTab + DateTime.Now + Constants.vbNewLine);
                sbError.Append("Err Message:" + Constants.vbTab + Message + Constants.vbNewLine);
                sbError.Append("Err Source" + Constants.vbTab + Source + Constants.vbTab + "" + Constants.vbTab + "" + Constants.vbNewLine);
                sbError.Append("Err Trace:" + Constants.vbTab + StackTrace + Constants.vbNewLine);
                sbError.Append("*********************************************************************" + Constants.vbTab + Constants.vbNewLine + Constants.vbNewLine);
                errFile.WriteLine(sbError.ToString());
                errFile.Flush();
                errFile.Close();
                errFile.Dispose();
            }
        }

        #region "Encryption and decryption
        public static string SHA1EncryptionNew(string stringToEncrypt)
        {
            byte[] key = { };

            byte[] inputByteArray = null;
            try
            {

                SHA1CryptoServiceProvider des = new SHA1CryptoServiceProvider();

                inputByteArray = System.Text.Encoding.UTF8.GetBytes(stringToEncrypt);

                MemoryStream ms = new MemoryStream();

                CryptoStream cs = new CryptoStream(ms, des, CryptoStreamMode.Write);

                cs.Write(inputByteArray, 0, inputByteArray.Length);

                cs.FlushFinalBlock();

                return Convert.ToBase64String(ms.ToArray());

            }
            catch
            {
                return (string.Empty);
            }
        }
        public static string SHA1DecryptNew(string strText)
        {
            byte[] bKey = new byte[21];
            byte[] IV = {
                     10,
                     20,
                     30,
                     40,
                     50,
                     60,
                     70,
                     80
              };

            try
            {

                SHA1CryptoServiceProvider DES = new SHA1CryptoServiceProvider();

                byte[] inputByteArray = Convert.FromBase64String(strText);

                MemoryStream ms = new MemoryStream();

                CryptoStream cs = new CryptoStream(ms, DES, CryptoStreamMode.Write);

                cs.Write(inputByteArray, 0, inputByteArray.Length);

                cs.FlushFinalBlock();

                System.Text.Encoding encoding = System.Text.Encoding.UTF8;

                return encoding.GetString(ms.ToArray());
            }
            catch (Exception ex)
            {
                throw ex;

            }

        }
        #endregion
        #region Encrypt Password
        public static string GetEncryptData(string Value1)
        {
            string functionReturnValue = null;
            string a = string.Empty;
            dynamic TT = null;
            dynamic X = null;
            dynamic Y = null;
            dynamic i = null;
            TT = "";
            functionReturnValue = "";
            for (i = 1; i <= Strings.Len(Value1); i++)
            {
                X = Strings.Asc(Strings.Mid(Value1, i, 1));
                Y = X ^ 0x54;
                Y = Y ^ 200;
                TT = TT + Strings.Chr(Y);
            }
            functionReturnValue = TT;
            return functionReturnValue;
        }
        #endregion
        #region "sendmail"
        public static void mailsend(string fromaddr, string toaddr, string subject, string body)
        {
            var message = new MailMessage();
            ArrayList list_emails = new ArrayList();
            //list_emails.Add("v-hcl.ibrahim@karvy.com");
            //list_emails.Add("venkatasiva.nagaraju@karvy.com");
            fromaddr = ConfigurationManager.AppSettings["smtpUserName"];
            if (toaddr != "")
            {
                toaddr = "venkatasiva.nagaraju@karvy.com";
                list_emails.Add(toaddr);
                foreach (var item in list_emails)
                {
                    message.To.Add(new MailAddress(item.ToString()));
                }
                message.From = new MailAddress(fromaddr);
                using (var smtp = new SmtpClient())
                {
                    var credential = new NetworkCredential
                    {
                        UserName = fromaddr,  // replace with valid value
                        Password = ConfigurationManager.AppSettings["smtpPassword"]  // replace with valid value
                    };
                    message.IsBodyHtml = true;
                    message.Subject = subject;
                    message.Body = body;
                    smtp.Credentials = credential;
                    smtp.Host = ConfigurationManager.AppSettings["Server"];
                    smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["smtpserverport"]);
                    smtp.EnableSsl = false;
                    smtp.Send(message);
                }
            }

        }
        #endregion
    }
}