using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Kloc.Models
{
    public class UserLogin : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            HttpContext ctx = HttpContext.Current;

            //if (HttpContext.Current.Session["UM_NAME"] == null || HttpContext.Current.Session.SessionID.ToString() == "")
            if (HttpContext.Current.Session.SessionID.ToString() == null)
            {
                filterContext.Result = new RedirectResult("~/Kloc/Login");
                return;
            }
          //  base.OnActionExecuting(filterContext);
        }
    }
}
