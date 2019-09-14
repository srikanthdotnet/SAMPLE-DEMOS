using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
namespace Kloc.Models
{
    public class CustomResultFilter : ActionFilterAttribute, IResultFilter
    {
       
        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {

            var test = "hai2";
        }
        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            var test = "hai1";
        }
    }
}