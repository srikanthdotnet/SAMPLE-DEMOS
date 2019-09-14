using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
namespace Kloc.Models
{
    public class BaseController : Controller, IExceptionFilter, IActionFilter,IResultFilter
    {
        ILog Logger = new ExceptionFactory();
        public BaseController()
        {

        }
       // void OnException(ExceptionContext filterContext);
        protected override void OnException(ExceptionContext filterContext)
        {   //using system.diagnostics stack trace class is there,by using this one  we can get the line number(where the eception occured).
            var st = new StackTrace(filterContext.Exception.InnerException, true);
            // Get the top stack frame
            var frame = st.GetFrame(0);
            // Get the line number from the stack frame       
            string line = frame.GetFileLineNumber().ToString();
            string controllerName = filterContext.RouteData.Values["controller"].ToString();
            string actionName = filterContext.RouteData.Values["action"].ToString();
            Logger.LogError(filterContext.Exception.InnerException, controllerName, actionName, line);
            //filterContext.Result = new RedirectResult("~/Shared/ErrorPage.cshtml");
            //filterContext.Result = new ViewResult
            //{
            //    ViewName = "~/Views/Shared/ErrorPage.cshtml"
            //};
            filterContext.ExceptionHandled = true;

        }
    }
}