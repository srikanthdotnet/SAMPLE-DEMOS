using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Kloc.Models;
using System.IO;
using System.Text;
namespace Kloc.Models
{
    public class ExceptionFactory : ILog
    {
        public  void LogError(Exception ex, string controller, string method, string LineNumber)
        {
            StreamWriter errfile;
            StringBuilder sbError = new StringBuilder();
            if (!System.IO.Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile"))
            {
                System.IO.Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile");
            }
            errfile = File.AppendText(System.AppDomain.CurrentDomain.BaseDirectory + "\\LogFile\\Error" + DateTime.Now.ToString("ddMMMyyyy") + ".log");
            sbError.Append("Err Date:" + DateTime.Now.ToString() + "\r\n");
            sbError.Append("Err Desc:" + ex.Message + "\r\n");
            sbError.Append("Controller:" + controller + "\r\n");
            sbError.Append("Method:" + method + "\r\n");
            sbError.Append("LineNumber:" + "\t" + LineNumber + "\r\n");
            sbError.Append("===============================================" + "\r\n" + "\r\n");
            errfile.WriteLine(sbError.ToString());
            errfile.Close();      
        }

        public void ActionMethodLog(string username, string controller, string method, string CurrentFund)
        {
            StreamWriter errfile;
            StringBuilder sbError = new StringBuilder();
            if (!System.IO.Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile"))
            {
                System.IO.Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile");
            }
            errfile = File.AppendText(System.AppDomain.CurrentDomain.BaseDirectory + "\\LogFile\\Userlog" + DateTime.Now.ToString("ddMMMyyyy") + ".log");
            sbError.Append("LogDate:" + DateTime.Now.ToString() + "\r\n");
            sbError.Append("CurrentUserName:" + username + "\r\n");
            sbError.Append("Controller:" + controller + "\r\n");
            sbError.Append("Method:" + method + "\r\n");
            sbError.Append("CurrentFund:" + CurrentFund + "\r\n");
            sbError.Append("===============================================" + "\r\n" + "\r\n");
            errfile.WriteLine(sbError.ToString());
            errfile.Close();
        }
        //public void OnException(ExceptionContextfilterContext)  
        //{  
        //    if (!filterContext.ExceptionHandled && filterContext.ExceptionisArgumentOutOfRangeException) {  
        //        filterContext.Result = newRedirectResult("~/Content/RangeErrorPage.html");  
        //        filterContext.ExceptionHandled = true;  
        //    }  
  
        //}  
    }
}