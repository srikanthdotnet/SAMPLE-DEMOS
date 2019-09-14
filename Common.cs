using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Xml;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Kloc.Models
{
    public  static class Common
    {
        public static bool Datevalidation(DateTime fromdate, DateTime Todate)
        {
            int i = Convert.ToInt32((Todate - fromdate).TotalDays);
            if (i >= 91)
            {
                return false;

            }
            return true;
        }

        public static string SerializeToXml<T>(T obj)
        {
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "");
            StringWriter Output = new StringWriter(new StringBuilder());
            XmlSerializer ser = new XmlSerializer(obj.GetType());
            ser.Serialize(Output, obj, ns);
            return Output.ToString();
        }
    }
}