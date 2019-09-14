using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;
using System.IO;
using System.Text;
namespace KlocBal
{
    public class BllCommonUtility
    {
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