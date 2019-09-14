using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Kloc.Models
{
    interface ILog
    {
        void LogError(Exception ex, string controller, string method, string Linenum);
        void ActionMethodLog(string username, string controller, string method, string CurrentFund);
    }
}