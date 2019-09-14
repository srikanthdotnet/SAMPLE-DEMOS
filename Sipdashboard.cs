using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.ComponentModel.DataAnnotations;
namespace KlocModel
{
    [DataContract]
    public sealed class Sipdashboard
    {

        private Sipdashboard()
        {

        }
        // private static Lazy<Image> imgInstance = new Lazy<Image>(() => new Image());
        private static Lazy<Sipdashboard> instance = new Lazy<Sipdashboard>(() => new Sipdashboard());
        private static readonly object lockObject = new object(); //create lock object restricting (multithreading senario creating objects).
        //private static Sipdashboard instance = null;


        public static Sipdashboard GetInstance
        {
            get
            {
                return instance.Value;
            }
        }

        //public static Sipdashboard GetInstance
        //{
        //    get
        //    {
        //        lock (lockObject)
        //        {
        //            if (instance == null)
        //            {
        //                instance = new Sipdashboard();
        //            }
        //        }
        //        return instance;
        //    }
        //}

        [DataMember]
        public string fund { get; set; }
        [DataMember]
        public string flg { get; set; }
        [DataMember]
        public string Fromdt { get; set; }
        [DataMember]
        public string Todate { get; set; }
        [DataMember]
        public string Mode { get; set; }
        [DataMember]
        public string  Remarks { get; set; }
    }
}
