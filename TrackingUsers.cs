using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Net;
namespace Kloc.Models
{
    public class TrackingUsers : ActionFilterAttribute
    {
        ILog Logger = new ExceptionFactory();
        //{
        //    public static void CreateTraceFile(string str)
        //    {
        //        File.AppendAllText(HttpContext.Current.Server.MapPath("~/TraceLog/Trace.txt"), str);
        //    }
        //    public override void OnActionExecuting(ActionExecutingContext filterContext)
        //    {
        //        string str1 = "controllername" + " " + filterContext.ActionDescriptor.ControllerDescriptor.ControllerName + "------->" +
        //                      "actionmethodname" + " " + filterContext.ActionDescriptor.ActionName + "------->"
        //                      + "executed on :" + " " + DateTime.Now.ToString() + "------->"
        //                      + "executed by:" + " " + HttpContext.Current.Session["UM_NAME"].ToString() + "------->"
        //                      + "current fund:" + " " + HttpContext.Current.Session["currentFund"].ToString();
        //        CreateTraceFile(str1);

        //        //base.OnActionExecuting(filterContext);
        //    }

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            string controllername = filterContext.RouteData.Values["controller"].ToString();
            string actionname = filterContext.RouteData.Values["action"].ToString();
            var _currentUser = HttpContext.Current.Session["UM_NAME"] == null ? "" : Convert.ToString(HttpContext.Current.Session["UM_NAME"]);
            var _currentFund = HttpContext.Current.Session["currentFund"] == null ? "" : Convert.ToString(HttpContext.Current.Session["currentFund"]);
            Logger.ActionMethodLog(_currentUser, controllername, actionname, _currentFund);
        }
    }
}

//  filterContext.Result = new RedirectResult("~/Kloc/Login");