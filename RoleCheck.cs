using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Kloc.Models
{
    public class RoleCheck:IAuthorizationFilter
    {
        //void OnAuthorization(AuthorizationContext filterContext);
        public void OnAuthorization(AuthorizationContext filterContext)
        {
            if (HttpContext.Current.Session["currentFund"] == null)
            {
                filterContext.Result = new RedirectResult("~/Kloc/Dashbord");
                HttpContext.Current.Session["FundMessage"] = "Please select Fund to view the report";
                return;
            }
        }
    }
}

