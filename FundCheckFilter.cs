using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
namespace Kloc.Models
{
    public class FundCheckFilter : ActionFilterAttribute
    {
        //void OnAuthorization(AuthorizationContext filterContext);
        //void OnException(ExceptionContext filterContext);
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            HttpContext ctx = HttpContext.Current;

            if (HttpContext.Current.Session["currentFund"] == null)
            {
                filterContext.Result = new RedirectResult("~/Kloc/Dashbord");
                HttpContext.Current.Session["FundMessage"] = "Please select Fund to view the report";
                return;
            }
            base.OnActionExecuting(filterContext);//called by asp.net mvc framework before the action method exceutes.
            // base.OnActionExecuted(filterContext);//called by asp.net mvc framework  after the action method exceutes.
        }
      
    }
}