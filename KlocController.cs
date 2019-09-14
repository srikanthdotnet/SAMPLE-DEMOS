using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Kloc.KarvyLoginService;
using Kloc.KlocBalService;
using System.Net;
using System.Data;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.Web.Security;
using Kloc.Models;
using KlocModel;
using Newtonsoft.Json;
using System.Globalization;
using System.Threading;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;
using System.Web.Script.Serialization;
using System.Text;
using System.ServiceModel.Web;
using Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;
using System.Configuration;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using KlocSqlHelper;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;
using System.Collections;

namespace Kloc.Controllers
{     //CriticalSms
    public class KlocController : BaseController
    {

        #region Login
        [HttpGet]
        public ActionResult Login()
        {

            //X509Certificate2Collection certificates = new X509Certificate2Collection();
            //DataSet obj1 = new DataSet();
            Session.Clear();
            return View();
        }

        public DataSet test()
        {
            DataSet ds = new DataSet("Dataset1");
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("name", typeof(string));
            dt.Rows.Add(1, "a");
            dt.Rows.Add(2, "b");
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Id1", typeof(int));
            dt1.Columns.Add("name1", typeof(string));
            dt1.Rows.Add(11, "aa");
            dt1.Rows.Add(22, "bb");
            ds.Tables.Add(dt);
            ds.Tables.Add(dt1);
            return ds;
        }
        [HttpPost]

        public ActionResult Login(string userid, string Password, string Fund)
        {
            try
            {

                string msg = "";
                KarvyLoginService.Login obj = new KarvyLoginService.Login();
                KarvyLoginService.NexgBllServiceClient serviceObj = new KarvyLoginService.NexgBllServiceClient();
                obj.UserID = userid;
                obj.Password = EncryptString(Password);
                obj.Fund = Fund;
                obj.UserType = "E";
                obj.sessionid = Session.SessionID.ToString();
                string hostName = Dns.GetHostName();// Retrive the Name of HOST
                obj.IPAddress = Dns.GetHostByName(hostName).AddressList[0].ToString();
                System.Data.DataTable dtObj = new System.Data.DataTable();
                dtObj = serviceObj.BvalidateLogin_Mobile(obj);
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();


                comm = balobj.GetAllFunds(userid);
                var data = comm.ds.Tables[0];
                List<SelectListItem> fundslistobj = GetFundDetails(comm);
                Session["Funddata"] = fundslistobj;

                if (dtObj != null && dtObj.Rows.Count > 0)
                {
                    if (dtObj.Columns.Contains("UM_NAME"))
                    {
                        HttpContext.Session["UM_NAME"] = dtObj.Rows[0]["UM_NAME"].ToString();
                        HttpContext.Session["UM_EMAILID"] = dtObj.Rows[0]["UM_EMAILID"].ToString();
                        HttpContext.Session["UM_UID"] = dtObj.Rows[0]["UM_UID"].ToString();

                        return Json(Newtonsoft.Json.JsonConvert.SerializeObject(dtObj), JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(Newtonsoft.Json.JsonConvert.SerializeObject(dtObj), JsonRequestBehavior.AllowGet);
                    }

                }
                if (dtObj.Rows.Count == 0)
                {
                    ViewBag.error = "Invalid Password";
                    msg = ViewBag.error;
                    return Json(new { ErrorMsg = "Invalid Password", Status = false });
                }
            }
            catch (Exception ex)
            {
                DataSet resultDataSet = new DataSet();
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Loginpost");
                resultDataSet = Util.GetErrorcode("100", ex.Message);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            return RedirectToAction("Login");
        }
        #endregion

        #region SMS_MIS
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult SMSMIS()
        {

            return View();
        }

        public JsonResult Getsms_misdashboard(string Fund)
        {
            DataSet resultDataSet = null;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            try
            {
                resultDataSet = balobj.GetSmsMis("RMF");
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Getdcrdashboard");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region CriticalSms
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult CriticalSms()
        {

            return View();
        }

        public JsonResult GetCriticalSms(string Fund)
        {
            DataSet resultDataSet = null;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            try
            {
                resultDataSet = balobj.GetCriticalSms(Fund);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Getdcrdashboard");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }
        public void CriticalSmsExcel(string Fund)
        {

            DataSet data = null;
            DataSet resultDataSet = new DataSet("CRITICALSMS");
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            resultDataSet = balobj.GetCriticalSms(Fund);




            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + "CriticalSms" + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }



        #endregion

        #region Dashboard
        [UserLogin]
        [HttpGet]
        public ActionResult Dashbord()
        {
            return View();
        }
        #endregion Dahboard

        #region GetUserFunds

        public List<SelectListItem> GetFundDetails(KlocModel.CommonReturnType objCommon)
        {
            List<SelectListItem> listItems = new List<SelectListItem>();
            listItems.Add(new SelectListItem
            {
                Text = "Select Fund",
                Value = "",

            });

            foreach (DataRow row in objCommon.ds.Tables[0].Rows)
            {
                listItems.Add(new SelectListItem
                {
                    Text = row["fundname"].ToString(),
                    Value = row["Fund"].ToString(),
                });
            }


            return listItems;
        }
        #endregion

        #region GetFundDetailsBasedOnId
        public ActionResult GetFundDetailsBasedOnId(string value)
        {
            DataSet resultset = new DataSet();
            try
            {

                var userbasedfunds = Session["Funddata"] as List<SelectListItem>;
                List<SelectListItem> currentobj = new List<SelectListItem>();
                SelectListItem obj = new SelectListItem();
                Userfunds fundobj = new Userfunds();
                foreach (var item in userbasedfunds)
                {
                    if (value == item.Value)
                    {
                        fundobj.UserID = item.Value;
                        fundobj.Fund = item.Text;
                        currentobj.Add(obj);
                        break;
                    }
                }

                Session["currentFund"] = fundobj.UserID.ToString();
                Session["currentFundName"] = fundobj.Fund.ToString().ToUpper();
                resultset = Util.GetErrorcode(fundobj.UserID, fundobj.Fund);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultset), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "GetFundDetailsBasedOnId");
                resultset = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultset), JsonRequestBehavior.AllowGet);
            }

        }
        #endregion

        #region sipdashboard
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult SipDashboard()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Getallsipdata(string Fromdate, string Todate, string Fund, int Flg)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            KlocBalServiceClient balobj = new KlocBalServiceClient();

            try
            {
                string fund = Fund;
                string format = "dd/MM/yyyy";

                DateTime fromdate1;
                DateTime todate1;
                if (DateTime.TryParseExact(Fromdate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out fromdate1) && DateTime.TryParseExact(Todate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out todate1) && Common.Datevalidation(fromdate1, todate1))
                {

                    if (Flg == 1)
                    {
                        string flg = Flg.ToString();
                        var fromdateformat = Fromdate;
                        var todateformat = Todate;
                        var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                        var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                        string fromdate = fromdateChanged;
                        string todate = todateChanged;
                        Session["fund"] = fund;
                        KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();

                        comm = balobj.Getallsipdashboarddata(fund, flg, fromdate, todate);
                        data = comm.ds;
                        // ---Remining Reports dataload-----

                        var flg1 = "2";

                        KlocModel.CommonReturnType channel = new KlocModel.CommonReturnType();
                        KlocModel.CommonReturnType exchangeobj = new KlocModel.CommonReturnType();
                        KlocModel.CommonReturnType isipobj = new KlocModel.CommonReturnType();
                        KlocBalServiceClient balobj1 = new KlocBalServiceClient();
                        channel = balobj1.GetChannelSIPDashboard(fund, flg1, fromdate, todate);//Getting channel data
                        exchangeobj = balobj1.GetExchangeSIPDashboard(fund, flg1, fromdate, todate);//getting exchange data
                        isipobj = balobj1.GetISIPDashboard(fund, flg1, fromdate, todate);//getting isip data.
                        System.Data.DataTable channeldtobj = channel.ds.Tables[0].Copy();
                        channeldtobj.TableName = "channel";
                        // channeldtobj.TableName = "Table15";
                        System.Data.DataTable exchangedtobj = exchangeobj.ds.Tables[0].Copy();
                        exchangedtobj.TableName = "exchange";
                        // channeldtobj.TableName = "Table16";
                        System.Data.DataTable isipdtobj = isipobj.ds.Tables[0].Copy();
                        isipdtobj.TableName = "isip";
                        System.Data.DataTable sipcancellation = isipobj.ds.Tables[1].Copy();
                        sipcancellation.TableName = "sipcancellation";
                        isipdtobj.TableName = "isip";
                        // channeldtobj.TableName = "Table17";
                        data.Tables.Add(channeldtobj);
                        data.Tables.Add(exchangedtobj);
                        data.Tables.Add(isipdtobj);
                        data.Tables.Add(sipcancellation);
                    }
                    else
                    {
                        Thread.Sleep(1000);
                        data = balobj.GetSIPDashboardScheduler(fund);

                    }
                    //end of remaining reports load----
                    return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
                }
                else
                {
                    resultDataSet = Util.GetErrorcode("100", "Report cannot be generated for the given Timelines.Please select a Valid range of 3 months to generate the Report");
                    resultDataSet.Tables[0].TableName = "ErrorTable";
                    return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
                }
            }

            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Getallsipdata");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }

        }
        #endregion

        #region onlinesipdashboard
        public ActionResult Onlinesipdashboard()
        {
            return View();
        }

        #endregion

        #region NctDashboard
        [UserLogin]
        [FundCheckFilter]
        public ActionResult NctDashboard()
        {
            return View();
        }
        [HttpGet]
        public ActionResult GetallNctDashboarddata(string Fromdate, string Todate, string Fund, string flg, string conditionflag)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            DataSet dtdata = new DataSet();
            try
            {
                string format = "dd/MM/yyyy";
                DateTime fromdate1;
                DateTime todate1;
                var a = DateTime.TryParseExact(Fromdate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out fromdate1);
                var b = DateTime.TryParseExact(Todate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out todate1);
                var c = Common.Datevalidation(fromdate1, todate1);
                if (DateTime.TryParseExact(Fromdate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out fromdate1) && DateTime.TryParseExact(Todate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out todate1) && Common.Datevalidation(fromdate1, todate1))
                {
                    string fund = Fund;
                    var fromdateformat = Fromdate;
                    var todateformat = Todate;
                    var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                    var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                    string fromdate = fromdateChanged;
                    string todate = todateChanged;
                    KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                    KlocBalServiceClient balobj = new KlocBalServiceClient();
                    comm = balobj.GetNCTDashboard(fund, flg, fromdate, todate);
                    if (conditionflag == "1")
                    {
                        data = comm.ds;
                        return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        //For export to excel total Dashboard Dump
                        data = comm.ds;
                        string filename = "AllNctdashboards";
                        var currentfilename = NCTDataSheetwise(data, filename, Todate);
                        var filepath = currentfilename;
                        return Json(new { Path = filepath + ".xlsx" }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    resultDataSet = Util.GetErrorcode("100", "Report cannot be generated for the given Timelines.Please select a Valid range of 3 months to generate the Report");
                    resultDataSet.Tables[0].TableName = "ErrorTable";
                    return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "GetallNctDashboarddata");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }


        public string NCTDataSheetwise(System.Data.DataSet dt, string FileName, string fromdate)
        {
            try
            {
                string file = FileName + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss").ToString();
                string filename = Server.MapPath("~/AccountStmts/") + file;
                XLWorkbook xlWorkbook = new XLWorkbook();
                xlWorkbook.Worksheets.Add(dt.Tables[0], "NCTReportedvsPending");
                xlWorkbook.Worksheets.Add(dt.Tables[1], "QRCbifurcationonpendingnumbers");
                xlWorkbook.Worksheets.Add(dt.Tables[2], "TATAdherence");
                xlWorkbook.Worksheets.Add(dt.Tables[3], "Zonewithpendency");
                xlWorkbook.Worksheets.Add(dt.Tables[4], "Top10NCTpendingSubjects");
                xlWorkbook.SaveAs(Server.MapPath("~/AccountStmts/") + file + ".xlsx");
                return file;
            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportTOExcel1()");
                return "";
            }

        }
        public void NctExportExcel(string Fromdate, string Todate, string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            var fromdateformat = Fromdate;
            var todateformat = Todate;
            var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
            var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
            string fromdate = fromdateChanged;
            string todate = todateChanged;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            comm = balobj.GetNCTDashboard(Fund, flg, fromdate, todate);
            resultDataSet = comm.ds;

            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }

        #endregion

        #region NctClientDashboard
        [UserLogin]
        [FundCheckFilter]
        [NoCache]
        public ActionResult NctClientDashboard()
        {
            return View();
        }
        [HttpGet]
        public ActionResult GetallNctClientDashboarddata(string Fromdate, string Todate, string Fund, string flg)
        {
            DataSet resultDataSet = null;

            DataSet dtdata = new DataSet();
            try
            {
                DataSet dss = new DataSet();
                System.Data.DataTable dt = new System.Data.DataTable();
             
                dss.Tables.Add(dt1);
                dss.DataSetName = "Dataset1";

                return Json(JsonConvert.SerializeObject(dss), JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "GetallNctClientDashboarddata");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }

        //[HttpGet]
        //public ActionResult GetallNctClientDashboarddata(string Fromdate, string Todate, string Fund, string flg)
        //{
        //    DataSet resultDataSet = null;

        //    DataSet dtdata = new DataSet();
        //    try
        //    {
        //        DataSet dss = new DataSet();
        //        System.Data.DataTable dt = new System.Data.DataTable();
        //        //dt.Columns.Add("Category", typeof(string));
        //        //dt.Columns.Add("Withinclosed", typeof(string));
        //        //dt.Columns.Add("within%", typeof(string));
        //        //dt.Columns.Add("withinopen", typeof(string));
        //        //dt.Columns.Add("withinopen%", typeof(string));
        //        //dt.Columns.Add("withintotal", typeof(string));
        //        //dt.Columns.Add("beyondclosed", typeof(string));
        //        //dt.Columns.Add("beyond%", typeof(string));
        //        //dt.Columns.Add("beyondopen", typeof(string));
        //        //dt.Columns.Add("beyondopen%", typeof(string));
        //        //dt.Columns.Add("beyondtotal%", typeof(string));
        //        //dt.Columns.Add("GrandTotal", typeof(string));
        //        //dt.Columns.Add("Total%", typeof(string));
        //        dt.Columns.Add("a", typeof(string));
        //        dt.Columns.Add("b", typeof(string));
        //        dt.Columns.Add("c", typeof(string));
        //        dt.Columns.Add("d", typeof(string));
        //        dt.Columns.Add("e", typeof(string));
        //        dt.Columns.Add("f", typeof(string));
        //        dt.Columns.Add("g", typeof(string));
        //        dt.Columns.Add("h", typeof(string));
        //        dt.Columns.Add("i", typeof(string));
        //        dt.Columns.Add("j", typeof(string));
        //        dt.Columns.Add("k", typeof(string));
        //        dt.Columns.Add("l", typeof(string));
        //        dt.Columns.Add("m", typeof(string));
        //        dt.Rows.Add("Exchange", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12");
        //        dt.Rows.Add("Channel", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42");
        //        dt.Rows.Add("MFU", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62");
        //        dt.Rows.Add("Grand Total", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt.TableName = "Table1";
        //        dss.Tables.Add(dt);
        //        //second datatable---

        //        System.Data.DataTable dt1 = new System.Data.DataTable();
        //        //dt1.Columns.Add("Subject Categorization", typeof(string));
        //        //dt1.Columns.Add("Withinclosed", typeof(string));
        //        //dt1.Columns.Add("within%", typeof(string));
        //        //dt1.Columns.Add("withinopen", typeof(string));
        //        //dt1.Columns.Add("withinopen%", typeof(string));
        //        //dt1.Columns.Add("withintotal", typeof(string));
        //        //dt1.Columns.Add("beyondclosed", typeof(string));
        //        //dt1.Columns.Add("beyond%", typeof(string));
        //        //dt1.Columns.Add("beyondopen", typeof(string));
        //        //dt1.Columns.Add("beyondopen%", typeof(string));
        //        //dt1.Columns.Add("beyondtotal%", typeof(string));
        //        //dt1.Columns.Add("GrandTotal", typeof(string));
        //        //dt1.Columns.Add("Total%", typeof(string));
        //        dt1.Columns.Add("a", typeof(string));
        //        dt1.Columns.Add("b", typeof(string));
        //        dt1.Columns.Add("c", typeof(string));
        //        dt1.Columns.Add("d", typeof(string));
        //        dt1.Columns.Add("e", typeof(string));
        //        dt1.Columns.Add("f", typeof(string));
        //        dt1.Columns.Add("g", typeof(string));
        //        dt1.Columns.Add("h", typeof(string));
        //        dt1.Columns.Add("i", typeof(string));
        //        dt1.Columns.Add("j", typeof(string));
        //        dt1.Columns.Add("k", typeof(string));
        //        dt1.Columns.Add("l", typeof(string));
        //        dt1.Columns.Add("m", typeof(string));
        //        dt1.Rows.Add("Change in Investor Profile", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("DEE", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("SOArelated", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("T+0", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Brokerage", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Change In Folio Details", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("DEE (Correction & Changes Broker/EUIN/scheme)", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Transmission", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("T+3", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Brokerage (Revalidation/AUM merger)", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("De-mat/Re-mat", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Dividend", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Purchase", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Redemption", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("SIP", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Switch", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("T+7", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.Rows.Add("Grand Total", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
        //        dt1.TableName = "Table2";

        //        dss.Tables.Add(dt1);
        //        dss.DataSetName = "Dataset1";

        //        return Json(JsonConvert.SerializeObject(dss), JsonRequestBehavior.AllowGet);

        //    }
        //    catch (Exception ex)
        //    {
        //        Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "GetallNctClientDashboarddata");
        //        resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
        //        return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
        //    }
        //}

        [HttpPost]
        public ActionResult NctClientDashboardSendemail(string Fromdate, string Todate, string Fund)
        {
            DataSet resultDataSet = null;

            DataSet dtdata = new DataSet();
            try
            {
                DataSet dss = new DataSet();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("Category", typeof(string));
                dt.Columns.Add("Withinclosed", typeof(string));
                dt.Columns.Add("within%", typeof(string));
                dt.Columns.Add("withinopen", typeof(string));
                dt.Columns.Add("withinopen%", typeof(string));
                dt.Columns.Add("withintotal", typeof(string));
                dt.Columns.Add("beyondclosed", typeof(string));
                dt.Columns.Add("beyond%", typeof(string));
                dt.Columns.Add("beyondopen", typeof(string));
                dt.Columns.Add("beyondopen%", typeof(string));
                dt.Columns.Add("beyondtotal%", typeof(string));
                dt.Columns.Add("GrandTotal", typeof(string));
                dt.Columns.Add("Total%", typeof(string));
                dt.Rows.Add("Exchange", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12");
                dt.Rows.Add("Channel", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42");
                dt.Rows.Add("MFU", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62");
                dt.Rows.Add("Grand Total", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt.TableName = "Table1";
                //second datatable---
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("Subject Categorization", typeof(string));
                dt1.Columns.Add("Withinclosed", typeof(string));
                dt1.Columns.Add("within%", typeof(string));
                dt1.Columns.Add("withinopen", typeof(string));
                dt1.Columns.Add("withinopen%", typeof(string));
                dt1.Columns.Add("withintotal", typeof(string));
                dt1.Columns.Add("beyondclosed", typeof(string));
                dt1.Columns.Add("beyond%", typeof(string));
                dt1.Columns.Add("beyondopen", typeof(string));
                dt1.Columns.Add("beyondopen%", typeof(string));
                dt1.Columns.Add("beyondtotal%", typeof(string));
                dt1.Columns.Add("GrandTotal", typeof(string));
                dt1.Columns.Add("Total%", typeof(string));
                dt1.Rows.Add("Change in Investor Profile", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("DEE", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("SOArelated", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("T+0", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Brokerage", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Change In Folio Details", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("DEE (Correction & Changes Broker/EUIN/scheme)", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Transmission", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("T+3", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Brokerage (Revalidation/AUM merger)", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("De-mat/Re-mat", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Dividend", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Purchase", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Redemption", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("SIP", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("Switch", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("T+7", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.Rows.Add("GrandTotal", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100", "100");
                dt1.TableName = "Table2";
                dss.Tables.Add(dt1);
                dss.DataSetName = "Dataset1";
                StringBuilder mailBody = new StringBuilder();
                //                mailBody.Append(@"<table width='65%' align='center' style='color: rgb(51, 51, 51); font-family: Arial, Helvetica, sans-serif; font-size: 14px;' border='0' cellspacing='0' cellpadding='0'>
                //                               <tbody><tr><td>Dear User,</td></tr><br><tr><td>Please find the below dashboard of Complaints/Query & Request and T+0. PFA  row level data </td></tr><tr>");
                //                mailBody.Append(@"<tr></tr></tbody></table>");
                mailBody.Append(@"Dear User,");
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");
                mailBody.Append("</br>");
                mailBody.Append(@"Please find the below dashboard of Complaints/Query & Request and T+0. PFA  row level data");
                string body = mailBody.ToString();
                mailBody.Append(@"</br>");
                //Firsttable starts here
                mailBody.Append(@"<div style='border: 2px solid red; width: 100%''>
            <div style='text-align: center; font-weight: bold; background-color: #f6700e;color:#fff; width: 100%'>
            COMPLAINTS CATEGORY WISE TAT%
            </div>");
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color: green; color: white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Category</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background: #a6a6a6'>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                //end of First  Table. 
                //Second Table Starts Here.
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added
                var firstTest = mailBody.ToString();
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt1.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color: green; color: white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Subject Categorization</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background: #a6a6a6'>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt1.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt1.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                ////end of second  Table.

                mailBody.Append(@"</div>");
                //--------------------------ending of First Task-------------------
                //--------------------------Second Task start here-------------------
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added
                //Firsttable starts here
                mailBody.Append(@"<div style='border: 2px solid red; width: 100%''>
            <div style='text-align: center; font-weight: bold; background-color: #f6700e;color:#fff; width: 100%'>
            Query Request CATEGORY WISE TAT%
            </div>");
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color:green;color:white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Category</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background:#a6a6a6'>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                //end of First  Table. 
                //Second Table Starts Here.
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt1.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color:green;color:white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Subject Categorization</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2'  style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background:#a6a6a6'>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt1.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt1.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                ////end of second  Table.

                mailBody.Append(@"</div>");
                //--------------------------Third Task Start Here-------------------
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added
                mailBody.Append(@"<div style='border: 2px solid red; width: 100%'>");
                //Firsttable starts here
                mailBody.Append(@"
            <div style='text-align: center; font-weight: bold; background: #f6700e;color:#fff; width: 100%'>
            T+0   DASHBOARD
            </div>");
                mailBody.Append(@"</br>");
                mailBody.Append(@"<div style='text-align: center; font-weight: bold; background-color: green; width: 100%'>
            Total
            </div>");
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color: green; color: white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Category</th>");
                    mailBody.Append(@"<th colspan='5'style='text-align: center; color: white; font-weight: bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align: center; color: white; font-weight: bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align: center; color: white; font-weight: bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background:#a6a6a6'>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>CLOSED</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>OPEN</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>%</th>");
                    mailBody.Append(@"<th style='background: #a6a6a6;'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                //end of First  Table. 
                //Second Table Starts Here.
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added
                mailBody.Append(@"<div style='text-align: center; font-weight: bold; background: #f6700e;color:#fff; width: 100%'>
           Before 3pm
            </div>");
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color: green; color: white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Category</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align:center;color:white;font-weight:bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align:center;color:white;font-weight:bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background:#a6a6a6'>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                ////end of second  Table.
                mailBody.Append(@"</br>");
                mailBody.Append(@"</br>");//added

                mailBody.Append(@"<div style='text-align: center; font-weight: bold; background: #f6700e;color:#fff; width: 100%'>
            After 3pm
            </div>");
                //Third Table Start Here--
                mailBody.Append(@"<table  border='1' style='width:100%;border-collapse:collapse;'><thead><tr>");
                if (dt.Columns.Count > 0)
                {
                    mailBody.Append(@"<tr style='background-color:green;color:white;'>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Category</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align:center;color:white;font-weight:bold;background: #00b050;'>WITHIN</th>");
                    mailBody.Append(@"<th colspan='5' style='text-align:center;color:white;font-weight:bold;background: #ff0000;'>Beyond</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Grand Total</th>");
                    mailBody.Append(@"<th rowspan='2' style='text-align:center;color:white;font-weight:bold;background: #3f61a7;'>Total%</th>");
                    mailBody.Append(@"</tr>");
                    mailBody.Append(@"<tr style='background:#a6a6a6'>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>CLOSED</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>OPEN</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>%</th>");
                    mailBody.Append(@"<th style='background:#a6a6a6'>TOTAL</th>");
                    mailBody.Append(@"</tr>");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        mailBody.Append(@"<tr style='color:BLACK;font-weight:bold'>");

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            mailBody.Append(@"<td>" + dt.Rows[i][j] + "</td>");
                        }
                        mailBody.Append(@"</tr>");
                    }
                }
                mailBody.Append(@"</table>");
                //---Third Table End Here---
                mailBody.Append(@"</br>");//added
                mailBody.Append(@"</div>");
                mailBody.Append(@"</div>");
                var c = mailBody.ToString();

                //string fromaddr = ConfigurationManager.AppSettings["smtpUserName"];
                //string pwd = ConfigurationManager.AppSettings["smtpPassword"];
                //--Zip file configuration---
                //ZipFile zip = new ZipFile()


                string filename = "NctClientExcel" + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xlsx";
                string cc = ExportToExcelnctlientdashboardSingleSheet(dss, filename, DateTime.Now.ToString(), "nctclientdahboard");
                //cc="NctClientExcel21052019152315.xlsx"
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + cc;
                System.Net.Mail.Attachment attachment;
                System.Net.Mail.Attachment attachment1;
                attachment = new System.Net.Mail.Attachment(currentfilename);
                attachment1 = new System.Net.Mail.Attachment(currentfilename);
                //---end of zip file configuration
                string fromaddr = "samfd@karvy.com";
                string pwd = " ";
                string toAddr = "srikanth.14@karvy.com";
                using (MailMessage mail = new MailMessage(fromaddr, toAddr))
                {
                    mail.Subject = "NctClient Dashboard report is for the period" + Fromdate + "to" + Todate;
                    mail.IsBodyHtml = true;
                    mail.Body = mailBody.ToString();
                    mail.Attachments.Add(attachment);
                    mail.Attachments.Add(attachment1);

                    SmtpClient smtp = new SmtpClient(("smtp.karvy.com"));
                    //smtp.Host = ConfigurationManager.AppSettings["Server"];
                    smtp.Host = "192.168.14.80";
                    smtp.EnableSsl = false;
                    NetworkCredential networkCredential = new NetworkCredential(fromaddr, pwd);
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = networkCredential;

                    //smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["smtpserverport"]);
                    smtp.Port = Convert.ToInt32("25");
                    smtp.Send(mail);//public void Send(MailMessage message);
                }

                return Json(new { Result = "ok" });
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "NctClientDashboardSendemail");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }



        public string ExportToExcelnctlientdashboardSingleSheet(System.Data.DataSet dt, string filename, string fromdate, string currentHeading)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);
                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];

                xlWorkSheetToExport.Cells[1, 1] = currentHeading;

                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:N1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                for (int i = 1; i < dt.Tables[0].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[3, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[3, i] = dt.Tables[0].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[0].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[0].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 4, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 4, k + 1] = dt.Tables[0].Rows[j].ItemArray[k].ToString();
                    }
                }

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
                // SAVE THE FILE IN A FOLDER.
                //filename="NctClientExcel21052019152315.xlsx"
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;

            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportToExcelSingleSheet()");
                return "";
            }

        }
        public void NctClientExportExcel(string Fromdate, string Todate, string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            var fromdateformat = Fromdate;
            var todateformat = Todate;
            var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
            var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
            string fromdate = fromdateChanged;
            string todate = todateChanged;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            comm = balobj.GetNCTDashboard(Fund, flg, fromdate, todate);
            resultDataSet = comm.ds;

            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }

        #endregion

        #region DCRDashboard
        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        public ActionResult DCRDashboard()
        {
            return View();
        }
        public JsonResult Getdcrdashboard(string Fund)
        {
            DataSet resultDataSet = null;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            string flg = "0";
            try
            {
                resultDataSet = balobj.GetDcrDashboardScheduler(Fund, flg);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Getdcrdashboard");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }

        public void ExportExcel(string Fund, string flag, string FileName)
        {
            string strFileName = string.Empty;
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            DataSet resultDataSet = null;
            resultDataSet = balobj.GetDcrDashboardScheduler(Fund, flag);
            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, " ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }
        public JsonResult DCR_ExportExcel(string Fund, string flag, string FileName)
        {
            string strFileName = string.Empty;
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            DataSet resultDataSet = null;
            resultDataSet = balobj.GetDcrDashboardScheduler(Fund, flag);
            string currentfilenamewithrandomnumber = FileName + Strings.Format(DateTime.Now, " ddMMyyyyHHmmss") + ".xlsx";
            var currentfilename = ExportTOExcel(resultDataSet.Tables[0], currentfilenamewithrandomnumber, FileName, flag);
            return Json(new { savedfilename = currentfilename }, JsonRequestBehavior.AllowGet);
        }
        public string ExportTOExcel(System.Data.DataTable dt, string filename, string mainfilename, string flag)
        {
            string currentfilename;
            string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
            string path = currentdirectory + "Exportedfiles\\";
            currentfilename = path + filename;
            if (System.IO.File.Exists(currentfilename))
            {
                System.IO.File.Delete(currentfilename);
            }
            if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
            {
                Directory.CreateDirectory(path);
            }
            currentfilename = path + filename;
            if (System.IO.File.Exists(currentfilename))
            {
                System.IO.File.Delete(currentfilename);
            }
            // ADD A WORKBOOK USING THE EXCEL APPLICATION.
            Microsoft.Office.Interop.Excel.Application xlAppToExport = new Microsoft.Office.Interop.Excel.Application();
            xlAppToExport.Workbooks.Add("");

            // ADD A WORKSHEET.
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport = default(Microsoft.Office.Interop.Excel.Worksheet);
            xlWorkSheetToExport = (Microsoft.Office.Interop.Excel.Worksheet)xlAppToExport.Sheets["Sheet1"];

            // ROW ID FROM WHERE THE DATA STARTS SHOWING.
            int iRowCnt = 4;
            var fromdate = DateTime.Now.ToString();
            string currentHeading = mainfilename + " " + "Report as on " + fromdate;
            // SHOW THE HEADER.
            xlWorkSheetToExport.Cells[1, 1] = currentHeading;

            Microsoft.Office.Interop.Excel.Range range = xlWorkSheetToExport.Cells[1, 12];
            range.EntireRow.Font.Name = "Calibri";
            range.EntireRow.Font.Bold = true;
            range.EntireRow.Font.Size = 20;

            xlWorkSheetToExport.Range["A1:I1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
            if (flag == "1A")
            {
                // SHOW COLUMNS ON THE TOP.
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "DTRentryDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "BatchcloseDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "BrokerCode";
                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["DTRentryDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["BatchcloseDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "1B")
            {
                // SHOW COLUMNS ON THE TOP.
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "DTRentryDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "BatchcloseDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["DTRentryDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["BatchcloseDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["BrokerCode"];
                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "2")
            {
                // SHOW COLUMNS ON THE TOP.
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "QCDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "BrokerCode";
                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["QCDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "3A")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "ProcessedDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "ProcessedFlag";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["ProcessedDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["ProcessedFlag"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["BrokerCode"];
                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "3B")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "ProcessedDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "ProcessedFlag";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "BrokerCode";
                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["ProcessedDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["ProcessedFlag"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "4")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "FundingDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "FundingStandard";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["FundingDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["FundingStandard"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["BrokerCode"];
                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "5")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "FundingDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "FundingStandard";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["FundingDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["FundingStandard"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["BrokerCode"];
                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "6")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "BrokerCode";


                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["BrokerCode"];
                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "7")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "8")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "Navdt";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "QCDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "DispatchDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "EmailID";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 17] = "DispTAT";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 18] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 19] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["Navdt"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["QCDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["DispatchDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["EmailID"];
                    xlWorkSheetToExport.Cells[iRowCnt, 17] = dt.Rows[i]["DispTAT"];
                    xlWorkSheetToExport.Cells[iRowCnt, 18] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 19] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "9")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "ProcessedDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Trno";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "ProcessedDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "DispatchDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "EmailID";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "DispTAT";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 17] = "Mode";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 18] = "NatureofDiscrepancy";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["ProcessedDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Trno"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["ProcessedDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["DispatchDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["EmailID"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["DispTAT"];
                    xlWorkSheetToExport.Cells[iRowCnt, 17] = dt.Rows[i]["Mode"];
                    xlWorkSheetToExport.Cells[iRowCnt, 18] = dt.Rows[i]["NatureofDiscrepancy"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "10")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "UCRDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "WeekDay";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Scheme";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Plan";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "IhNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "TranType";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Nameoftheinvestor";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "Amount";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "TrDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "NAVDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "BatchcloseDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "Ageing";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "NatureofDiscrepancy";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "BrokerCode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["UCRDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["WeekDay"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["Scheme"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Plan"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["IhNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["TranType"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Nameoftheinvestor"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["Amount"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["TrDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["NAVDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["BatchcloseDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["Ageing"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["NatureofDiscrepancy"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["BrokerCode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            else if (flag == "11")
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "Sl.no";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "FundHouse";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "FolioNo";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Branch";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Subject Code";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "Subject";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "IHNum";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "weekday";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "Inward_status";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 10] = "InvName";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 11] = "InwardDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 12] = "OutwardDate";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 13] = "Status";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 14] = "Ageing";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 15] = "ReqCompSource";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 16] = "Mode";

                int i;
                for (i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i]["Sl.no"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i]["FundHouse"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i]["FolioNo"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i]["Branch"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i]["Subject Code"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Rows[i]["Subject"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Rows[i]["IHNum"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Rows[i]["weekday"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Rows[i]["Inward_status"];
                    xlWorkSheetToExport.Cells[iRowCnt, 10] = dt.Rows[i]["InvName"];
                    xlWorkSheetToExport.Cells[iRowCnt, 11] = dt.Rows[i]["InwardDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 12] = dt.Rows[i]["OutwardDate"];
                    xlWorkSheetToExport.Cells[iRowCnt, 13] = dt.Rows[i]["Status"];
                    xlWorkSheetToExport.Cells[iRowCnt, 14] = dt.Rows[i]["Ageing"];
                    xlWorkSheetToExport.Cells[iRowCnt, 15] = dt.Rows[i]["ReqCompSource"];
                    xlWorkSheetToExport.Cells[iRowCnt, 16] = dt.Rows[i]["Mode"];

                    iRowCnt = iRowCnt + 1;
                }
            }
            // FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION
            Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1];
            range1.AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatList3);

            // SAVE THE FILE IN A FOLDER.
            xlWorkSheetToExport.SaveAs(path + filename);
            // CLEAR.
            xlAppToExport.Workbooks.Close();

            xlAppToExport.Quit();
            xlAppToExport = null;
            xlWorkSheetToExport = null;
            return filename;

        }

        #region oldcode
        //public JsonResult DCR_ExportExcel(string Fund, string flag, string FileName)
        //{
        //    string strFileName = string.Empty;
        //    KlocBalServiceClient balobj = new KlocBalServiceClient();
        //    DataSet resultDataSet = null;
        //    resultDataSet = balobj.GetDcrDashboardScheduler(Fund, flag);
        //    //string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
        //    //string path = currentdirectory + "AccountStmts\\" + FileName;
        //    string currentfilename;
        //    string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
        //    string path = currentdirectory + "Exportedfiles\\";
        //    double dRandomNo = VBMath.Rnd(1) * 10000;
        //    string currentfilenamewithrandomnumber = FileName + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + dRandomNo + ".xlsx";
        //    currentfilename = path + currentfilenamewithrandomnumber;
        //    if (System.IO.File.Exists(currentfilename))
        //    {
        //        System.IO.File.Delete(currentfilename);
        //    }
        //    if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
        //    {
        //        Directory.CreateDirectory(path);
        //    }

        //    if (System.IO.File.Exists(currentfilename))
        //    {
        //        System.IO.File.Delete(currentfilename);
        //    }
        //    var isExcel = XLSX(resultDataSet.Tables[0], currentfilename);
        //    return Json(new { obj = isExcel,savedfilename=currentfilenamewithrandomnumber }, JsonRequestBehavior.AllowGet);
        //}
        #endregion

        //public bool XLSX(System.Data.System.Data.DataTable dt, string FileNameIncludingDirectoryFolder)
        //{
        //    try
        //    {
        //        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        //        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        //        object misValue = System.Reflection.Missing.Value;


        //        xlWorkBook = xlApp.Workbooks.Add(misValue);

        //        xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        //        WriteArray(dt, xlWorkSheet);


        //        xlWorkBook.SaveAs(FileNameIncludingDirectoryFolder, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //        xlWorkBook.Close(true, misValue, misValue);
        //        xlApp.Quit();
        //        return true;
        //    }
        //    catch (Exception)
        //    {
        //        return false;
        //    }
        //}

        //public void WriteArray(System.Data.System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        //{

        //    for (int i = dt.Columns.Count - 1; i >= 0; i--)
        //    {
        //        Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range("A1", Type.Missing);
        //        rng.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight,
        //                                 Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        //        worksheet.Range["A1", Type.Missing].Value2 = dt.Columns[i].ColumnName;
        //    }

        //    var data1 = new object[dt.Rows.Count, dt.Columns.Count];
        //    for (var row = 0; row < dt.Rows.Count; row++)
        //    {
        //        for (var column = 0; column < dt.Columns.Count; column++)
        //        {
        //            data1[row, column] = dt.Rows[row][column];
        //        }
        //    }
        //    var startCell1 = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[2, 1];

        //    var endCell1 = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[dt.Rows.Count + 1, dt.Columns.Count];
        //    var writeRange1 = worksheet.Range[startCell1, endCell1];
        //    writeRange1.Value2 = data1;
        //}
        #endregion

        #region BankingDashboard
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult BankingDashboard()
        {
            var d = Convert.ToString(Session.SessionID);
            return View();
        }
        [HttpPost]
        public ActionResult GetallBankingDashboarddata(string Fromdate, string Fund, int flg)
        {
            DataSet resultDataSet = null;
            string fund = null;
            try
            {
                DataSet data = null;
                fund = Fund;
                var fromdateformat = Fromdate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                Session["fromdate"] = fromdate;
                Session["fund"] = fund;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                if (flg == 0)
                {
                    comm = balobj.GetBANKINGDashboard(fund, fromdate);
                    data = comm.ds;
                }
                if (flg == 1)
                {
                    data = balobj.GetBankingDashboardScheduler(fund);
                }

                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "GetallBankingDashboarddata()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }

        }


        #endregion

        #region ExchangeDashboard
        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]

        public ActionResult ExchangeDashboard()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ExchangeDashboard(string Todate, string Fund, string flg)
        {
            DataSet resultDataSet = null;
            var intervaldata = string.Empty;
            try
            {
                var todateformat = Todate;
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                intervaldata = todateChanged;

                var yesterdaydate = intervaldata.ToString();
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                resultDataSet = balobj.GetExchangeDashboard(Fund, flg, yesterdaydate);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExchangeDashboard()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }


        public void ExchangeExportExcel(string Fund, string flag, string FileName, string Todate, string exchangeType)
        {
            string strFileName = string.Empty;
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            DataSet resultDataSet = null;
            var intervaldata = string.Empty;
            //var intervaldata = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy");
            var todateformat = Todate;
            var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
            intervaldata = todateChanged;
            var yesterdaydate = intervaldata.ToString();
            resultDataSet = balobj.GetExchangeDashboard(Fund, flag, yesterdaydate);

            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                if (exchangeType == "Purchase")
                {
                    gv.DataSource = resultDataSet.Tables[0];
                }
                else
                {
                    gv.DataSource = resultDataSet.Tables[1];
                }
                gv.DataBind();
                gv.HeaderRow.Style.Add("background-color", "#FFFFFF");
                gv.BorderStyle = BorderStyle.Solid;
                gv.BorderWidth = 2;
                gv.BackColor = System.Drawing.Color.Yellow;
                gv.GridLines = GridLines.Both;
                gv.Font.Name = "Verdana";
                gv.Font.Size = FontUnit.Medium;
                gv.HeaderStyle.BackColor = System.Drawing.Color.GreenYellow;
                gv.HeaderStyle.ForeColor = System.Drawing.Color.Gray;
                gv.RowStyle.HorizontalAlign = HorizontalAlign.Left;
                gv.RowStyle.VerticalAlign = VerticalAlign.Top;
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + yesterdaydate + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }


        //public bool ExchangeExportExcel(string Fund, string flag, string FileName, string Todate)
        //{
        //    DataSet ds = test();
        //    try
        //    {
        //        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
        //        foreach (System.Data.DataTable table in ds.Tables)
        //        {
        //            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
        //            excelWorkSheet.Name = table.TableName;
        //            for (int i = 1; i <= table.Columns.Count + 1 - 1; i++)
        //            {
        //                excelWorkSheet.Cells(1, i) = table.Columns(i - 1).ColumnName;
        //                excelWorkSheet.Cells(1, i).Font.Bold = true;
        //                excelWorkSheet.Cells(i).ColumnWidth = 30;
        //            }
        //            for (int j = 0; j <= table.Rows.Count - 1; j++)
        //            {
        //                for (int k = 0; k <= table.Columns.Count - 1; k++)
        //                    excelWorkSheet.Cells(j + 2, k + 1) = table.Rows(j).ItemArray(k).ToString();
        //            }
        //            if ((!System.IO.File.Exists(FileName)))
        //            {
        //                excelWorkSheet.SaveAs(FileName);
        //                excelWorkBook = excelApp.Workbooks.Open(FileName);
        //            }
        //        }
        //        excelWorkBook.Save();
        //        excelWorkBook.Close();
        //        excelApp.Quit();
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

        [HttpPost]

        public ActionResult RMExportData(string Todate, string Fund, string flg)
        {
            DataSet resultDataSet = null;
            try
            {
                var dt = test();
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1 = dt.Tables[0];

                var todateformat = Todate;
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                var yesterdaydate = todateChanged.ToString();
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                resultDataSet = balobj.GetExchangeDashboard(Fund, flg, yesterdaydate);
                string filename = "ExchangeDashboardReport.xls";
                var currentfilename = ExportTOExcel(resultDataSet, filename, Todate);
                var filepath = currentfilename;
                return Json(new { Path = filepath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "RMExportData");
                return Json(new { Path = "" }, JsonRequestBehavior.AllowGet);
            }

        }

        [HttpPost]

        public ActionResult ExchangeExportData(string Todate, string Fund, string flg)
        {
            DataSet resultDataSet = null;
            try
            {
                var dt = test();
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1 = dt.Tables[0];

                var todateformat = Todate;
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                var yesterdaydate = todateChanged.ToString();
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                resultDataSet = balobj.GetExchangeDashboard(Fund, flg, yesterdaydate);
                string filename = "ExchangeDashboardReport";
                var currentfilename = ExchangeDataSheetwise(resultDataSet, filename, Todate);
                var filepath = currentfilename;
                return Json(new { Path = filepath + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "RMExportData");
                return Json(new { Path = "" }, JsonRequestBehavior.AllowGet);
            }

        }


        public string ExchangeDataSheetwise(System.Data.DataSet dt, string FileName, string fromdate)
        {
            try
            {
                string file = FileName + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss").ToString();
                string filename = Server.MapPath("~/AccountStmts/") + file;
                XLWorkbook xlWorkbook = new XLWorkbook();
                xlWorkbook.Worksheets.Add(dt.Tables[0], "Purchase");
                xlWorkbook.Worksheets.Add(dt.Tables[1], "Redemption");
                xlWorkbook.SaveAs(Server.MapPath("~/AccountStmts/") + file + ".xlsx");
                return file;
            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportTOExcel1()");
                return "";
            }

        }

        public string ExportTOExcel(System.Data.DataSet dt, string filename, string fromdate)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);

                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];
                // ROW ID FROM WHERE THE DATA STARTS SHOWING.
                int iRowCnt = 4;
                // var fromdate = DateTime.Now.ToString();
                string currentHeading = "Purchase Report as on " + fromdate;
                string secondheading = "AXISMUTUAL FUND";

                //------
                // SHOW THE HEADER.
                xlWorkSheetToExport.Cells[1, 1] = currentHeading;
                xlWorkSheetToExport.Cells[2, 1] = secondheading;
                // xlWorkSheetToExport.Cells[2, 1] = "AXIS MUTUAL FUND EXHANGE DATA";
                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;
                xlWorkSheetToExport.Range["A1:I1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                //-----------
                // SHOW THE HEADER.

                // xlWorkSheetToExport.Cells[2, 1] = "AXIS MUTUAL FUND EXHANGE DATA";
                Range range11 = xlWorkSheetToExport.Cells[2, 12] as Range;
                range11.EntireRow.Font.Name = "Calibri";
                range11.EntireRow.Font.Bold = true;
                range11.EntireRow.Font.Size = 20;
                ////  ---row colors styles added
                //Range maxln = xlWorkSheetToExport.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                //Range range222 = xlWorkSheetToExport.get_Range("A1", maxln);
                //xlWorkSheetToExport.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, range222, Type.Missing, XlYesNoGuess.xlNo, Type.Missing).Name = "MyTableStyle";
                //xlWorkSheetToExport.ListObjects.get_Item("MyTableStyle").TableStyle = "TableStyleMedium1";
                ////  --- end of row colors styles added


                // xlWorkSheetToExport.Range["A2:I2"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                // SHOW COLUMNS ON THE TOP.
                xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "Mode";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Reported";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Rejected";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Processed";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 5] = "Pending";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 6] = "CaExecuted";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 7] = "NsdlDp";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 8] = "CdslDp";
                xlWorkSheetToExport.Cells[iRowCnt - 1, 9] = "CaPending";
                int i;
                for (i = 0; i <= dt.Tables[0].Rows.Count - 1; i++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Tables[0].Rows[i]["Mode"];
                    xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Tables[0].Rows[i]["Reported"];
                    xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Tables[0].Rows[i]["Rejected"];
                    xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Tables[0].Rows[i]["Processed"];
                    xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Tables[0].Rows[i]["Pending"];
                    xlWorkSheetToExport.Cells[iRowCnt, 6] = dt.Tables[0].Rows[i]["CaExecuted"];
                    xlWorkSheetToExport.Cells[iRowCnt, 7] = dt.Tables[0].Rows[i]["NsdlDp"];
                    xlWorkSheetToExport.Cells[iRowCnt, 8] = dt.Tables[0].Rows[i]["CdslDp"];
                    xlWorkSheetToExport.Cells[iRowCnt, 9] = dt.Tables[0].Rows[i]["CaPending"];
                    iRowCnt = iRowCnt + 1;
                }
                // FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Range;
                range1.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatList3);

                ////-------------sheet2 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                //Microsoft.Office.Interop.Excel.Application xlAppToExport = new Microsoft.Office.Interop.Excel.Application();
                ////Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                //xlAppToExport1.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport1 = default(Worksheet);
                xlWorkSheetToExport1 = (Worksheet)xlAppToExport.Sheets["Sheet2"];
                // ROW ID FROM WHERE THE DATA STARTS SHOWING.
                int iRowCnt1 = 4;
                var fromdate1 = DateTime.Now.ToString();
                string currentHeading1 = "Redemption Report as on " + fromdate1;
                // SHOW THE HEADER.
                xlWorkSheetToExport1.Cells[1, 1] = currentHeading1;
                Range range111 = xlWorkSheetToExport1.Cells[1, 12] as Range;
                range111.EntireRow.Font.Name = "Calibri";
                range111.EntireRow.Font.Bold = true;
                range111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport1.Range["A1:I1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                // SHOW COLUMNS ON THE TOP.




                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 1] = "Mode";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 2] = "Reported";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 3] = "Rejected";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 4] = "Processed";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 5] = "Pending";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 6] = "CaExecuted";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 7] = "NsdlDp";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 8] = "CdslDp";
                xlWorkSheetToExport1.Cells[iRowCnt1 - 1, 9] = "CaPending";


                int i1;
                for (i1 = 0; i1 <= dt.Tables[1].Rows.Count - 1; i1++)
                {
                    xlWorkSheetToExport1.Cells[iRowCnt1, 1] = dt.Tables[1].Rows[i1]["Mode"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 2] = dt.Tables[1].Rows[i1]["Reported"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 3] = dt.Tables[1].Rows[i1]["Rejected"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 4] = dt.Tables[1].Rows[i1]["Processed"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 5] = dt.Tables[1].Rows[i1]["Pending"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 6] = dt.Tables[1].Rows[i1]["CaExecuted"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 7] = dt.Tables[1].Rows[i1]["NsdlDp"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 8] = dt.Tables[1].Rows[i1]["CdslDp"];
                    xlWorkSheetToExport1.Cells[iRowCnt1, 9] = dt.Tables[1].Rows[i1]["CaPending"];
                    iRowCnt1 = iRowCnt1 + 1;
                }




                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.   --bug occured this below line
                //Range range2 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Range;
                //range2.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatList2);









                //-------------end of sheet2 data-------------
                // SAVE THE FILE IN A FOLDER.
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;
            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportTOExcel()");
                return "";
            }

        }

        #endregion


        public ActionResult RMExportData1(string Todate, string Fund, string flg)
        {
            DataSet resultDataSet = null;
            try
            {
                var dt = test();
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1 = dt.Tables[0];

                var todateformat = Todate;
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                var yesterdaydate = todateChanged.ToString();
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                resultDataSet = balobj.GetExchangeDashboard(Fund, flg, yesterdaydate);
                string filename = "ExchangeDashboardReport.xls";
                var currentfilename = ExportTOExcel1(resultDataSet, filename, Todate);
                var filepath = currentfilename;
                return Json(new { Path = filepath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "RMExportData1");
                return Json(new { Path = "" }, JsonRequestBehavior.AllowGet);
            }

        }

        public string ExportTOExcel1(System.Data.DataSet dt, string filename, string fromdate)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);
                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];
                //// ROW ID FROM WHERE THE DATA STARTS SHOWING.
                var currentdate = DateTime.Now.ToString();
                string currentHeading = "Axis Mutual Fund Purchase Report as on " + currentdate.ToString();

                xlWorkSheetToExport.Cells[1, 1] = currentHeading;

                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:N1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                for (int i = 1; i < dt.Tables[0].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[4, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[4, i] = dt.Tables[0].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[0].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[0].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 5, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 5, k + 1] = dt.Tables[0].Rows[j].ItemArray[k].ToString();
                    }
                }

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);


                ////-------------sheet2 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Microsoft.Office.Interop.Excel.Application xlAppToExport1 = new Microsoft.Office.Interop.Excel.Application();

                xlAppToExport1.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport1 = default(Worksheet);
                xlWorkSheetToExport1 = (Worksheet)xlAppToExport.Sheets["Sheet2"];
                var fromdate1 = DateTime.Now.ToString();
                string currentHeading1 = "Axis Mutual Fund  Redemption Report as on " + fromdate1;
                // SHOW THE HEADER.
                xlWorkSheetToExport1.Cells[1, 1] = currentHeading1;
                Range range111 = xlWorkSheetToExport1.Cells[1, 12] as Range;
                range111.EntireRow.Font.Name = "Calibri";
                range111.EntireRow.Font.Bold = true;
                range111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport1.Range["A1:N1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                //// SHOW COLUMNS ON THE TOP.



                for (int i = 1; i < dt.Tables[1].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport1.Cells[4, i].Font.Bold = true;
                    xlWorkSheetToExport1.Cells[4, i] = dt.Tables[1].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[1].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[1].Columns.Count; k++)
                    {
                        xlWorkSheetToExport1.Cells[j + 5, k + 1].Font.Bold = true;
                        xlWorkSheetToExport1.Cells[j + 5, k + 1] = dt.Tables[1].Rows[j].ItemArray[k].ToString();
                    }
                }
                ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

                //-------------end of sheet2 data-------------
                // SAVE THE FILE IN A FOLDER.
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;

            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportTOExcel1()");
                return "";
            }

        }

        #region ExchangeofflineCADashboard
        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult ExchangeofflineCADashboard()
        {
            return View();
        }
        [HttpGet]
        public ActionResult getExchangeofflineCADashboard(string Fund)
        {
            DataSet resultDataSet = null;
            var intervaldata = string.Empty;
            try
            {

                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                resultDataSet = balobj.GetExchangeofflineCA(Fund);
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "getExchangeofflineCADashboard()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }
        public void ExchangeOfflineCAExportExcel(string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            switch (flg)
            {
                case "offlineCounts":
                    resultDataSet = balobj.GetExchangeofflineCA(Fund);
                    break;
                case "1":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "2":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "3":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "4":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "5":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "6":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "7":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "8":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "9":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
                case "10":
                    resultDataSet = balobj.GetExchageofflineDetails(Fund, flg, "", "");
                    break;
            }
            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }
        #endregion

        #region PaytmDashboard
        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult PAYTMDashboard()
        {
            return View();
        }
        [HttpPost]
        public ActionResult PAYTMDashboard(string Fromdate, string Todate, string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            try
            {
                string fund = Fund;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                comm = balobj.GetPaytmDashboard(fund, flg, fromdate, todate, FileName);
                data = comm.ds;
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "PAYTMDashboard()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }

        public void PaytmExportExcel(string Fromdate, string Todate, string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            var fromdateformat = Fromdate;
            var todateformat = Todate;
            var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
            var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
            string fromdate = fromdateChanged;
            string todate = todateChanged;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            comm = balobj.GetPaytmDashboard(Fund, flg, fromdate, todate, FileName);
            resultDataSet = comm.ds;

            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[1];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }
        #endregion

        #region FundingandPayoutDashboard
        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult FundingandPayoutDashboard()
        {
            return View();
        }
        [HttpPost]
        public ActionResult FundingandPayoutDashboard(string Fromdate, string Todate, string Fund)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            try
            {
                string format = "dd/MM/yyyy";
                DateTime fromdate1;
                DateTime todate1;
                var a = DateTime.TryParseExact(Fromdate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out fromdate1);
                var b = DateTime.TryParseExact(Todate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out todate1);
                var c = Common.Datevalidation(fromdate1, todate1);
                if (DateTime.TryParseExact(Fromdate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out fromdate1) && DateTime.TryParseExact(Todate, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out todate1) && Common.Datevalidation(fromdate1, todate1))
                {
                    DataSet dss = new DataSet();
                    var fromdateChanged = Fromdate.Split('/')[1] + "/" + Fromdate.Split('/')[0] + "/" + Fromdate.Split('/')[2];
                    var todateChanged = Todate.Split('/')[1] + "/" + Todate.Split('/')[0] + "/" + Todate.Split('/')[2];

                    KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                    KlocBalServiceClient balobj = new KlocBalServiceClient();
                    comm = balobj.GetFundingandPayoutDashboard(Fund, "", fromdateChanged, todateChanged);
                    data = comm.ds;
                    return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
                }
                else
                {
                    resultDataSet = Util.GetErrorcode("100", "Report cannot be generated for the given Timelines.Please select a Valid range of 3 months to generate the Report");
                    resultDataSet.Tables[0].TableName = "ErrorTable";
                    return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "FundingandPayoutDashboard()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }
        public void FundingandPayoutExportExcel(string Fromdate, string Todate, string Fund)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            DataSet dss = new DataSet();
            var fromdateChanged = Fromdate.Split('/')[1] + "/" + Fromdate.Split('/')[0] + "/" + Fromdate.Split('/')[2];
            var todateChanged = Todate.Split('/')[1] + "/" + Todate.Split('/')[0] + "/" + Todate.Split('/')[2];

            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            comm = balobj.GetFundingandPayoutDashboard(Fund, "", fromdateChanged, todateChanged);
            data = comm.ds;
            resultDataSet = comm.ds;

            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[1];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + "FundingandPayoutDashboard" + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }
        #endregion

        #region CPDDashboard

        [HttpGet]
        [UserLogin]
        [FundCheckFilter]
        [TrackingUsers]
        public ActionResult CPDDashboard()
        {
            return View();
        }
        [HttpPost]

        public ActionResult CPDDashboard(string Fromdate, string Todate, string Fund, string flg, string FileName)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            try
            {
                string fund = Fund;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                comm = balobj.GetCPDDashboard(fund, flg, fromdate, todate, FileName);
                data = comm.ds;
                return Json(Newtonsoft.Json.JsonConvert.SerializeObject(data), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "CPDDashboard()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }


        public ActionResult CPDDashboardTotalDump(string Fromdate, string Todate, string Fund, string flg, string FileName, string currentfundtext)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            try
            {
                string fund = Fund;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = Fund;
                obj.flg = flg;
                obj.Fromdt = fromdateChanged;
                obj.Todate = todateChanged;
                string xmldata = Common.SerializeToXml(obj);// we are directly calling the sp here,beacause of maxbufer size issue.
                //all connectionstrings reading---------
                string constring = null;
                string ConnectionString = null;
                var ConStr = Convert.ToString(Fund);
                if (ConStr == "Mfdwebtest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Mfdwebtest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Kbolttest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Kbolttest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Karvymfstest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Karvymfstest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "KBOLT")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["KBOLT"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                    // ConnectionString = "Data Source=192.168.14.147;Initial Catalog=Reliance;User ID=rmfsecondary;Password=%bn745NY~";
                }

                else if (ConStr == "108")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["108"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "101")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["101"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "102")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["102"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "103")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["103"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "104")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["104"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "105")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["105"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "107")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["107"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "113")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["113"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "116")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["116"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "117")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["117"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "118")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["118"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "120")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["120"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "125")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["125"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "128")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["128"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "129")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["129"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "130")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["130"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "135")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["135"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                //string ConnectionString = "Data Source=192.168.10.20;Initial catalog=axismf;User Id=migration;Password=mig0106@s;";
                SqlConnection con = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "GetCPD_DashBorad_kloc";
                cmd.CommandType = CommandType.StoredProcedure;
                if (!string.IsNullOrWhiteSpace(xmldata))
                {
                    cmd.Parameters.AddWithValue("@Param", xmldata);
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 500;
                da.SelectCommand.CommandTimeout = 1000;
                DataSet resds = new DataSet();
                da.Fill(resds);
                //resultDataSet = DBHelper.ExecuteSP_GetDataSet(Fund, "GetCPD_DashBorad_ExcelReport_kloc", xmldata);
                //  string file = "FileName" + DateTime.Now.ToString("ddMMyyyyhhmmssfff");
                string file = FileName + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss").ToString();
                string filename = Server.MapPath("~/AccountStmts/") + file;
                XLWorkbook xlWorkbook = new XLWorkbook();
                //var ws = xlWorkbook.Worksheets.Add(resds.Tables[1]);

                //ws.Cell(1, 1).Value = "Purchase Data from   to  date";  // sets excel sheet header

                //var rangeTitle = ws.Range(3, 1, 3, 10);   // range for row 3, column 1 to row 3, column titles.Count
                //rangeTitle.AddToNamed("Titles");
                //// styles
                //var titlesStyle = xlWorkbook.Style;
                //titlesStyle.Font.Bold = true;
                //titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                //titlesStyle.Fill.BackgroundColor = XLColor.Amber;

                //// style titles row
                //xlWorkbook.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;
                //ws.Columns().AdjustToContents();
                xlWorkbook.Worksheets.Add(resds.Tables[1], "Purchase");
                xlWorkbook.Worksheets.Add(resds.Tables[3], "Redemption");
                xlWorkbook.Worksheets.Add(resds.Tables[5], "Switch");
                xlWorkbook.SaveAs(Server.MapPath("~/AccountStmts/") + file + ".xlsx");
                return Json(new { Flag = 0, Path = file + ".xlsx" });
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "CPDDashboardTotalDump()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }

        public void CpdExportExcel(string Fund, string flg, string FileName, string Todate, string Fromdate, string Mode, string Remarks)
        {
            try
            {
                DataSet resultDataSet = null;
                DataSet data = null;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();

                //comm = balobj.GetCPDDashboardDetails(Fund, flg, fromdate, todate, Mode, Remarks, FileName);
                //resultDataSet = comm.ds;
                ////// ROW ID FROM WHERE THE DATA STARTS SHOWING.
                //var currentdate = DateTime.Now.ToString();
                //string currentHeading = "Axis Mutual Fund Purchase Report as on " + currentdate.ToString();
                //var currentfilename = ExportToExcelSingleSheet(resultDataSet, FileName + ".xls", Todate, currentHeading);
                //var filepath = currentfilename;
                //return Json(new { Path = filepath }, JsonRequestBehavior.AllowGet);


                //comm = balobj.GetCPDDashboardDetails(Fund, flg, fromdate, todate, Mode, Remarks, FileName);
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = Fund;
                obj.flg = flg;
                obj.Fromdt = fromdateChanged;
                obj.Todate = todateChanged;
                obj.Mode = Mode;
                obj.Remarks = Remarks;
                string xmldata = Common.SerializeToXml(obj);// we are directly calling the sp here,beacause of maxbufer size issue.
                //all connectionstrings reading---------
                string constring = null;
                string ConnectionString = null;
                var ConStr = Convert.ToString(Fund);
                if (ConStr == "Mfdwebtest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Mfdwebtest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Kbolttest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Kbolttest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Karvymfstest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Karvymfstest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "KBOLT")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["KBOLT"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }


                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "108")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["108"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "101")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["101"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "102")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["102"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "103")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["103"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "104")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["104"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "105")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["105"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "107")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["107"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "113")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["113"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "116")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["116"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "117")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["117"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "118")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["118"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "120")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["120"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "125")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["125"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "128")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["128"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "129")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["129"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "130")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["130"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "135")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["135"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                //string ConnectionString = "Data Source=192.168.10.20;Initial catalog=axismf;User Id=migration;Password=mig0106@s;";
                SqlConnection con = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "GetCPD_DashBorad_ExcelReport_kloc";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 500;
                if (!string.IsNullOrWhiteSpace(xmldata))
                {
                    cmd.Parameters.AddWithValue("@Param", xmldata);
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.SelectCommand.CommandTimeout = 1000;
                DataSet resds = new DataSet();
                da.Fill(resds);
                if (resds.Tables[0].Rows.Count > 0)
                {
                    GridView gv = new GridView();
                    gv.DataSource = resds.Tables[0];
                    gv.DataBind();
                    Response.ClearContent();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", "attachment; filename=" + Mode + '(' + FileName + ')' + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xlsx");
                    Response.ContentType = "application/ms-excel";
                    Response.Charset = "";
                    StringWriter sw = new StringWriter();
                    HtmlTextWriter htw = new HtmlTextWriter(sw);
                    gv.RenderControl(htw);
                    Response.Output.Write(sw.ToString());
                    Response.Flush();
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "CpdExportExcel()");

            }
        }
        public ActionResult CpdExportExcel12(string Fund, string flg, string FileName, string Todate, string Fromdate, string Mode, string Remarks)
        {

            try
            {
                DataSet resultDataSet = null;
                DataSet data = null;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                //comm = balobj.GetCPDDashboardDetails(Fund, flg, fromdate, todate, Mode, Remarks, FileName);
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = Fund;
                obj.flg = flg;
                obj.Fromdt = fromdateChanged;
                obj.Todate = todateChanged;
                obj.Mode = Mode;
                obj.Remarks = Remarks;
                string xmldata = Common.SerializeToXml(obj);// we are directly calling the sp here,beacause of maxbufer size issue.
                //all connectionstrings reading---------
                string constring = null;
                string ConnectionString = null;
                var ConStr = Convert.ToString(Fund);
                if (ConStr == "Mfdwebtest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Mfdwebtest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Kbolttest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Kbolttest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Karvymfstest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Karvymfstest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "KBOLT")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["KBOLT"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                    // ConnectionString = "Data Source=192.168.14.147;Initial Catalog=Reliance;User ID=rmfsecondary;Password=%bn745NY~";
                }

                else if (ConStr == "108")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["108"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "101")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["101"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "102")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["102"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "103")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["103"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "104")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["104"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "105")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["105"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "107")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["107"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "113")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["113"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "116")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["116"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "117")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["117"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "118")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["118"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "120")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["120"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "125")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["125"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "128")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["128"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "129")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["129"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "130")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["130"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "135")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["135"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                //string ConnectionString = "Data Source=192.168.10.20;Initial catalog=axismf;User Id=migration;Password=mig0106@s;";

                SqlConnection con = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "GetCPD_DashBorad_ExcelReport_kloc";
                cmd.CommandType = CommandType.StoredProcedure;
                if (!string.IsNullOrWhiteSpace(xmldata))
                {
                    cmd.Parameters.AddWithValue("@Param", xmldata);
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 500;
                da.SelectCommand.CommandTimeout = 1000;
                DataSet resds = new DataSet();
                da.Fill(resds);
                //resultDataSet = DBHelper.ExecuteSP_GetDataSet(Fund, "GetCPD_DashBorad_ExcelReport_kloc", xmldata);
                //  string file = "FileName" + DateTime.Now.ToString("ddMMyyyyhhmmssfff");
                string file = Mode + '(' + FileName + ')' + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss").ToString();
                string filename = Server.MapPath("~/AccountStmts/") + file;
                string sheetname = Mode + '(' + FileName + ')';
                XLWorkbook xlWorkbook = new XLWorkbook();
                xlWorkbook.Worksheets.Add(resds.Tables[0], sheetname);
                xlWorkbook.SaveAs(Server.MapPath("~/AccountStmts/") + file + ".xlsx");
                return Json(new { Flag = 0, Path = file + ".xlsx" });

            }
            catch (Exception e)
            {

                Util.WriteLog(e.Message, e.Source, e.StackTrace, "CpdExportExcel12()");
                return Json(new { Flag = 0, Path = "" });
            }
        }
        public string ExportToExcelSingleSheet(System.Data.DataSet dt, string filename, string fromdate, string currentHeading)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);
                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];

                xlWorkSheetToExport.Cells[1, 1] = currentHeading;

                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:N1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                for (int i = 1; i < dt.Tables[0].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[3, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[3, i] = dt.Tables[0].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[0].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[0].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 4, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 4, k + 1] = dt.Tables[0].Rows[j].ItemArray[k].ToString();
                    }
                }

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
                // SAVE THE FILE IN A FOLDER.
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;

            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportToExcelSingleSheet()");
                return "";
            }

        }
        //public string ExportToExcelMultipleSheets(System.Data.DataSet dt, string filename, string fromdate, string Todate, string currentfundtext)
        //{
        //    try
        //    {
        //        string currentfilename;
        //        string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
        //        string path = currentdirectory + "AccountStmts\\";
        //        currentfilename = path + filename;
        //        if (System.IO.File.Exists(currentfilename))
        //        {
        //            System.IO.File.Delete(currentfilename);
        //        }
        //        if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
        //        {
        //            Directory.CreateDirectory(path);
        //        }
        //        currentfilename = path + filename;
        //        if (System.IO.File.Exists(currentfilename))
        //        {
        //            System.IO.File.Delete(currentfilename);
        //        }
        //        // ADD A WORKBOOK USING THE EXCEL APPLICATION.
        //        Application xlAppToExport = new Application();
        //        //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
        //        xlAppToExport.Workbooks.Add("");

        //        // ADD A WORKSHEET.
        //        Worksheet xlWorkSheetToExport = default(Worksheet);
        //        xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];
        //        //// ROW ID FROM WHERE THE DATA STARTS SHOWING.

        //        string currentHeading = currentfundtext + " Purchase Report from " + " " + fromdate.ToString() + " " + "To" + " " + Todate.ToString();

        //        xlWorkSheetToExport.Cells[1, 1] = currentHeading;

        //        Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
        //        range.EntireRow.Font.Name = "Calibri";
        //        range.EntireRow.Font.Bold = true;
        //        range.EntireRow.Font.Size = 20;

        //        xlWorkSheetToExport.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

        //        for (int i = 1; i < dt.Tables[1].Columns.Count + 1; i++)
        //        {
        //            xlWorkSheetToExport.Cells[4, i].Font.Bold = true;
        //            xlWorkSheetToExport.Cells[4, i] = dt.Tables[1].Columns[i - 1].ColumnName;
        //        }

        //        for (int j = 0; j < dt.Tables[1].Rows.Count; j++)
        //        {
        //            for (int k = 0; k < dt.Tables[1].Columns.Count; k++)
        //            {
        //                xlWorkSheetToExport.Cells[j + 5, k + 1].Font.Bold = true;
        //                xlWorkSheetToExport.Cells[j + 5, k + 1] = dt.Tables[1].Rows[j].ItemArray[k].ToString();
        //            }
        //        }

        //        //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
        //        Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
        //        range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
        //        ////-------------sheet2 data starts---------
        //        //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
        //        Microsoft.Office.Interop.Excel.Application xlAppToExport1 = new Microsoft.Office.Interop.Excel.Application();

        //        xlAppToExport1.Workbooks.Add("");

        //        //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
        //        // ADD A WORKSHEET.
        //        Worksheet xlWorkSheetToExport1 = default(Worksheet);
        //        xlWorkSheetToExport1 = (Worksheet)xlAppToExport.Sheets["Sheet2"];
        //        string currentHeading1 = currentfundtext + " Redemption Report from " + " " + fromdate.ToString() + "  " + "  " + "To" + " " + Todate.ToString();
        //        // SHOW THE HEADER.
        //        xlWorkSheetToExport1.Cells[1, 1] = currentHeading1;
        //        Range range111 = xlWorkSheetToExport1.Cells[1, 12] as Range;
        //        range111.EntireRow.Font.Name = "Calibri";
        //        range111.EntireRow.Font.Bold = true;
        //        range111.EntireRow.Font.Size = 20;

        //        xlWorkSheetToExport1.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
        //        //// SHOW COLUMNS ON THE TOP.



        //        for (int i = 1; i < dt.Tables[3].Columns.Count + 1; i++)
        //        {
        //            xlWorkSheetToExport1.Cells[4, i].Font.Bold = true;
        //            xlWorkSheetToExport1.Cells[4, i] = dt.Tables[3].Columns[i - 1].ColumnName;
        //        }

        //        for (int j = 0; j < dt.Tables[3].Rows.Count; j++)
        //        {
        //            for (int k = 0; k < dt.Tables[3].Columns.Count; k++)
        //            {
        //                xlWorkSheetToExport1.Cells[j + 5, k + 1].Font.Bold = true;
        //                xlWorkSheetToExport1.Cells[j + 5, k + 1] = dt.Tables[3].Rows[j].ItemArray[k].ToString();
        //            }
        //        }
        //        ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
        //        //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
        //        //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

        //        //-------------end of sheet2 data-------------

        //        ////-------------sheet3 data starts---------
        //        //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
        //        Microsoft.Office.Interop.Excel.Application xlAppToExport2 = new Microsoft.Office.Interop.Excel.Application();

        //        xlAppToExport2.Workbooks.Add("");

        //        //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
        //        // ADD A WORKSHEET.
        //        Worksheet xlWorkSheetToExport2 = default(Worksheet);
        //        xlWorkSheetToExport2 = (Worksheet)xlAppToExport.Sheets["Sheet3"];
        //        string currentHeading2 = currentfundtext + " Switch Report from " + " " + fromdate.ToString() + " " + " " + "To" + " " + Todate.ToString();
        //        // SHOW THE HEADER.
        //        xlWorkSheetToExport2.Cells[1, 1] = currentHeading2;
        //        Range range1111 = xlWorkSheetToExport2.Cells[1, 12] as Range;
        //        range1111.EntireRow.Font.Name = "Calibri";
        //        range1111.EntireRow.Font.Bold = true;
        //        range1111.EntireRow.Font.Size = 20;

        //        xlWorkSheetToExport2.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
        //        //// SHOW COLUMNS ON THE TOP.

        //        for (int i = 1; i < dt.Tables[5].Columns.Count + 1; i++)
        //        {
        //            xlWorkSheetToExport2.Cells[4, i].Font.Bold = true;
        //            xlWorkSheetToExport2.Cells[4, i] = dt.Tables[5].Columns[i - 1].ColumnName;
        //        }

        //        for (int j = 0; j < dt.Tables[5].Rows.Count; j++)
        //        {
        //            for (int k = 0; k < dt.Tables[5].Columns.Count; k++)
        //            {
        //                xlWorkSheetToExport2.Cells[j + 5, k + 1].Font.Bold = true;
        //                xlWorkSheetToExport2.Cells[j + 5, k + 1] = dt.Tables[5].Rows[j].ItemArray[k].ToString();
        //            }
        //        }
        //        ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
        //        //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
        //        //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

        //        //-------------end of sheet2 data-------------
        //        // SAVE THE FILE IN A FOLDER.
        //        xlWorkSheetToExport.SaveAs(path + filename);
        //        // CLEAR.
        //        xlAppToExport.Workbooks.Close();
        //        xlAppToExport.Quit();
        //        xlAppToExport = null;
        //        xlWorkSheetToExport = null;
        //        // ---
        //        return filename;

        //    }
        //    catch (Exception ex)
        //    {

        //        Util.WriteLog(ex.Message, ex.Source, ex.StackTrace);
        //        return "";
        //    }

        //}
        //public string ExportToExcelMultipleSheetsclosedxml(System.Data.DataSet dt, string filename, string fromdate)
        //{
        //    try
        //    {
        //        string currentfilename;
        //        string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
        //        string path = currentdirectory + "AccountStmts\\";
        //        currentfilename = path + filename;
        //        if (System.IO.File.Exists(currentfilename))
        //        {
        //            System.IO.File.Delete(currentfilename);
        //        }
        //        if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
        //        {
        //            Directory.CreateDirectory(path);
        //        }
        //        currentfilename = path + filename;
        //        if (System.IO.File.Exists(currentfilename))
        //        {
        //            System.IO.File.Delete(currentfilename);
        //        }


        //        // ---
        //        return filename;

        //    }
        //    catch (Exception ex)
        //    {

        //        Util.WriteLog(ex.Message, ex.Source, ex.StackTrace);
        //        return "";
        //    }

        //}
        public string ExportToExcelMultipleSheets(System.Data.DataSet dt, string filename, string fromdate, string Todate, string currentfundtext)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);
                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];
                //// ROW ID FROM WHERE THE DATA STARTS SHOWING.

                string currentHeading = currentfundtext + " Purchase Report from " + " " + fromdate.ToString() + " " + "To" + " " + Todate.ToString();

                xlWorkSheetToExport.Cells[1, 1] = currentHeading;

                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                for (int i = 1; i < dt.Tables[1].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[4, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[4, i] = dt.Tables[1].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[1].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[1].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 5, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 5, k + 1] = dt.Tables[1].Rows[j].ItemArray[k].ToString();
                    }
                }

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
                ////-------------sheet2 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Microsoft.Office.Interop.Excel.Application xlAppToExport1 = new Microsoft.Office.Interop.Excel.Application();

                xlAppToExport1.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport1 = default(Worksheet);
                xlWorkSheetToExport1 = (Worksheet)xlAppToExport.Sheets["Sheet2"];
                string currentHeading1 = currentfundtext + " Redemption Report from " + " " + fromdate.ToString() + "  " + "  " + "To" + " " + Todate.ToString();
                // SHOW THE HEADER.
                xlWorkSheetToExport1.Cells[1, 1] = currentHeading1;
                Range range111 = xlWorkSheetToExport1.Cells[1, 12] as Range;
                range111.EntireRow.Font.Name = "Calibri";
                range111.EntireRow.Font.Bold = true;
                range111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport1.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                //// SHOW COLUMNS ON THE TOP.



                for (int i = 1; i < dt.Tables[3].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport1.Cells[4, i].Font.Bold = true;
                    xlWorkSheetToExport1.Cells[4, i] = dt.Tables[3].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[3].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[3].Columns.Count; k++)
                    {
                        xlWorkSheetToExport1.Cells[j + 5, k + 1].Font.Bold = true;
                        xlWorkSheetToExport1.Cells[j + 5, k + 1] = dt.Tables[3].Rows[j].ItemArray[k].ToString();
                    }
                }
                ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

                //-------------end of sheet2 data-------------

                ////-------------sheet3 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Microsoft.Office.Interop.Excel.Application xlAppToExport2 = new Microsoft.Office.Interop.Excel.Application();

                xlAppToExport2.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport2 = default(Worksheet);
                xlWorkSheetToExport2 = (Worksheet)xlAppToExport.Sheets["Sheet3"];
                string currentHeading2 = currentfundtext + " Switch Report from " + " " + fromdate.ToString() + " " + " " + "To" + " " + Todate.ToString();
                // SHOW THE HEADER.
                xlWorkSheetToExport2.Cells[1, 1] = currentHeading2;
                Range range1111 = xlWorkSheetToExport2.Cells[1, 12] as Range;
                range1111.EntireRow.Font.Name = "Calibri";
                range1111.EntireRow.Font.Bold = true;
                range1111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport2.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                //// SHOW COLUMNS ON THE TOP.

                for (int i = 1; i < dt.Tables[5].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport2.Cells[4, i].Font.Bold = true;
                    xlWorkSheetToExport2.Cells[4, i] = dt.Tables[5].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[5].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[5].Columns.Count; k++)
                    {
                        xlWorkSheetToExport2.Cells[j + 5, k + 1].Font.Bold = true;
                        xlWorkSheetToExport2.Cells[j + 5, k + 1] = dt.Tables[5].Rows[j].ItemArray[k].ToString();
                    }
                }
                ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

                //-------------end of sheet2 data-------------
                // SAVE THE FILE IN A FOLDER.
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;

            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "ExportToExcelMultipleSheets()");
                return "";
            }

        }
        public ActionResult CPDDashboardTotalDump1(string Fromdate, string Todate, string Fund, string flg, string FileName, string currentfundtext)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            try
            {
                string fund = Fund;
                var fromdateformat = Fromdate;
                var todateformat = Todate;
                var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
                var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
                string fromdate = fromdateChanged;
                string todate = todateChanged;
                KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
                KlocBalServiceClient balobj = new KlocBalServiceClient();
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = Fund;
                obj.flg = flg;
                obj.Fromdt = fromdateChanged;
                obj.Todate = todateChanged;
                string xmldata = Common.SerializeToXml(obj);// we are directly calling the sp here,beacause of maxbufer size issue.
                //all connectionstrings reading---------
                string constring = null;
                string ConnectionString = null;
                var ConStr = Convert.ToString(Fund);
                if (ConStr == "Mfdwebtest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Mfdwebtest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Kbolttest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Kbolttest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "Karvymfstest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Karvymfstest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "KBOLT")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["KBOLT"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                    // ConnectionString = "Data Source=192.168.14.147;Initial Catalog=Reliance;User ID=rmfsecondary;Password=%bn745NY~";
                }

                else if (ConStr == "108")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["108"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "101")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["101"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "102")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["102"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "103")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["103"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "104")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["104"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "105")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["105"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "107")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["107"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "113")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["113"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "116")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["116"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "117")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["117"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "118")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["118"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "120")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["120"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "123")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "125")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["125"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "127")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["127"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "128")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["128"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "129")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["129"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "130")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["130"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else if (ConStr == "135")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["135"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                //string ConnectionString = "Data Source=192.168.10.20;Initial catalog=axismf;User Id=migration;Password=mig0106@s;";
                SqlConnection con = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandTimeout = 500;
                cmd.CommandText = "GetCPD_DashBorad_kloc";
                cmd.CommandType = CommandType.StoredProcedure;
                if (!string.IsNullOrWhiteSpace(xmldata))
                {
                    cmd.Parameters.AddWithValue("@Param", xmldata);
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.SelectCommand.CommandTimeout = 1000;
                DataSet resds = new DataSet();
                da.Fill(resds);
                //resultDataSet = DBHelper.ExecuteSP_GetDataSet(Fund, "GetCPD_DashBorad_ExcelReport_kloc", xmldata);
                //  string file = "FileName" + DateTime.Now.ToString("ddMMyyyyhhmmssfff");
                //string file = FileName + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss").ToString();
                //string filename = Server.MapPath("~/AccountStmts/") + file;
                var currentfile = Singlesheet(resds, FileName, fromdate, todate, currentfundtext);
                return Json(new { Flag = 0, Path = currentfile });
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "CPDDashboardTotalDump()");
                resultDataSet = Util.GetErrorcode("100", ex.Message.ToString());
                return Json(JsonConvert.SerializeObject(resultDataSet), JsonRequestBehavior.AllowGet);
            }
        }
        public string Singlesheet(System.Data.DataSet dt, string filename, string fromdate, string Todate, string currentfundtext)
        {
            try
            {
                string currentfilename;
                string currentdirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentdirectory + "AccountStmts\\";
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                {
                    Directory.CreateDirectory(path);
                }
                currentfilename = path + filename;
                if (System.IO.File.Exists(currentfilename))
                {
                    System.IO.File.Delete(currentfilename);
                }
                // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Application xlAppToExport = new Application();
                //Microsoft.Office.Interop.Excel.ApplicationClass xlAppToExport = new Microsoft.Office.Interop.Excel.ApplicationClass();
                xlAppToExport.Workbooks.Add("");

                // ADD A WORKSHEET.
                Worksheet xlWorkSheetToExport = default(Worksheet);
                xlWorkSheetToExport = (Worksheet)xlAppToExport.Sheets["Sheet1"];
                //// ROW ID FROM WHERE THE DATA STARTS SHOWING.

                string currentHeading = currentfundtext + " Purchase Report from " + " " + fromdate.ToString() + " " + "To" + " " + Todate.ToString();

                xlWorkSheetToExport.Cells[1, 1] = currentHeading;

                Range range = xlWorkSheetToExport.Cells[1, 12] as Range;
                range.EntireRow.Font.Name = "Calibri";
                range.EntireRow.Font.Bold = true;
                range.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A1:Y1"].MergeCells = true;// MERGE CELLS OF THE HEADER.

                for (int i = 1; i < dt.Tables[1].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[2, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[2, i] = dt.Tables[1].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[1].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[1].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 3, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 3, k + 1] = dt.Tables[1].Rows[j].ItemArray[k].ToString();
                    }
                }

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
                ////-------------sheet2 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Microsoft.Office.Interop.Excel.Application xlAppToExport1 = new Microsoft.Office.Interop.Excel.Application();

                xlAppToExport1.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.

                string currentHeading1 = currentfundtext + " Redemption Report from " + " " + fromdate.ToString() + "  " + "  " + "To" + " " + Todate.ToString();
                // SHOW THE HEADER.
                xlWorkSheetToExport.Cells[11, 11] = currentHeading1;
                Range range111 = xlWorkSheetToExport.Cells[11, 25] as Range;
                range111.EntireRow.Font.Name = "Calibri";
                range111.EntireRow.Font.Bold = true;
                range111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A11:Y11"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                //// SHOW COLUMNS ON THE TOP.



                for (int i = 1; i < dt.Tables[3].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[12, i].Font.Bold = true;
                    xlWorkSheetToExport.Cells[12, i] = dt.Tables[3].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[3].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[3].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 13, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 13, k + 1] = dt.Tables[3].Rows[j].ItemArray[k].ToString();
                    }
                }
                ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);

                //-------------end of sheet2 data-------------

                ////-------------sheet3 data starts---------
                //// ADD A WORKBOOK USING THE EXCEL APPLICATION.
                Microsoft.Office.Interop.Excel.Application xlAppToExport2 = new Microsoft.Office.Interop.Excel.Application();

                xlAppToExport2.Workbooks.Add("");

                //  Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport1 = default(Worksheet);
                // ADD A WORKSHEET.

                string currentHeading2 = currentfundtext + " Switch Report from " + " " + fromdate.ToString() + " " + " " + "To" + " " + Todate.ToString();
                // SHOW THE HEADER.
                xlWorkSheetToExport.Cells[20, 20] = currentHeading2;
                Range range1111 = xlWorkSheetToExport.Cells[20, 30] as Range;
                range1111.EntireRow.Font.Name = "Calibri";
                range1111.EntireRow.Font.Bold = true;
                range1111.EntireRow.Font.Size = 20;

                xlWorkSheetToExport.Range["A20:Y20"].MergeCells = true;// MERGE CELLS OF THE HEADER.
                //// SHOW COLUMNS ON THE TOP.

                for (int i = 1; i < dt.Tables[5].Columns.Count + 1; i++)
                {
                    xlWorkSheetToExport.Cells[21, i].Font.Bold = true;

                    xlWorkSheetToExport.Cells[21, i] = dt.Tables[5].Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Tables[5].Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Tables[5].Columns.Count; k++)
                    {
                        xlWorkSheetToExport.Cells[j + 22, k + 1].Font.Bold = true;
                        xlWorkSheetToExport.Cells[j + 22, k + 1] = dt.Tables[5].Rows[j].ItemArray[k].ToString();
                    }
                }
                ////// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Microsoft.Office.Interop.Excel.Range range12 = xlAppToExport1.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                //range12.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList2);
                //-------------end of sheet2 data-------------
                // SAVE THE FILE IN A FOLDER.
                xlWorkSheetToExport.SaveAs(path + filename);
                // CLEAR.
                xlAppToExport.Workbooks.Close();
                xlAppToExport.Quit();
                xlAppToExport = null;
                xlWorkSheetToExport = null;
                // ---
                return filename;

            }
            catch (Exception ex)
            {

                Util.WriteLog(ex.Message, ex.Source, ex.StackTrace, "Singlesheet()");
                return "";
            }

        }
        public void CpdExportExcel1(string Fund, string flg, string FileName, string Todate, string Fromdate, string Mode, string Remarks)
        {
            DataSet resultDataSet = null;
            DataSet data = null;
            var fromdateformat = Fromdate;
            var todateformat = Todate;
            var fromdateChanged = fromdateformat.Split('/')[1] + "/" + fromdateformat.Split('/')[0] + "/" + fromdateformat.Split('/')[2];
            var todateChanged = todateformat.Split('/')[1] + "/" + todateformat.Split('/')[0] + "/" + todateformat.Split('/')[2];
            string fromdate = fromdateChanged;
            string todate = todateChanged;
            KlocModel.CommonReturnType comm = new KlocModel.CommonReturnType();
            KlocBalServiceClient balobj = new KlocBalServiceClient();
            // comm = balobj.GetCPDDashboardDetails(Fund, flg, fromdate, todate, Mode, Remarks, FileName);
            resultDataSet = comm.ds;
            if (resultDataSet.Tables[0].Rows.Count > 0)
            {
                GridView gv = new GridView();
                gv.DataSource = resultDataSet.Tables[0];
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ' ' + Strings.Format(DateTime.Now, "ddMMyyyyHHmmss") + ".xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter sw = new StringWriter();
                HtmlTextWriter htw = new HtmlTextWriter(sw);
                gv.RenderControl(htw);
                Response.Output.Write(sw.ToString());
                Response.Flush();
                Response.End();
            }
        }
        #endregion

        #region Logout

        public ActionResult Keepalive()
        {
            return Json("OK", JsonRequestBehavior.AllowGet);
        }
        [TrackingUsers]
        public ActionResult LogOut()
        {
            FormsAuthentication.SignOut();
            Session.Clear();
            Session.Abandon();
            Session.RemoveAll();
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Cache.SetExpires(DateTime.UtcNow.AddHours(-1));
            Response.Cache.SetNoStore();
            return RedirectToAction("Login", "Kloc");
        }

        #endregion

        #region EncryptString
        /// <summary>
        /// This Method is used for password encryption
        /// </summary>
        /// <param name="strTxt">Password Parameter</param>
        /// <returns>encrypted string</returns>
        [NonAction]
        public string EncryptString(string strTxt)
        {
            int i = 0;
            //string c = null;
            char c;
            string cTemp = "";

            for (i = 1; i <= Strings.Len(strTxt); i++)
            {
                if (Strings.Asc(Strings.Mid(strTxt, i, 1)) < 127)
                {
                    c = Strings.Chr(Strings.Asc(Strings.Mid(strTxt, i, 1)) + 127);
                    cTemp = cTemp + c;
                }
                else
                {
                    cTemp = cTemp + Strings.Mid(strTxt, i, 1);
                }
            }
            return cTemp;
        }

        #endregion

        public ActionResult testResult()
        {
            return View();
        }

        #region PrivateMethods
        [NonAction]
        public static string HttpContent(string url)
        {
            WebRequest objRequest = System.Net.HttpWebRequest.Create(url);
            StreamReader sr = new StreamReader(objRequest.GetResponse().GetResponseStream());
            string result = sr.ReadToEnd();
            sr.Close();
            return result;
        }
        [NonAction]
        public static string TamperProofStringDecode(string value, string key)
        {
            string dataValue = "";
            string calcHash = "";
            string storedHash = "";

            System.Security.Cryptography.MACTripleDES mac3des = new System.Security.Cryptography.MACTripleDES();
            System.Security.Cryptography.MD5CryptoServiceProvider md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            mac3des.Key = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(key));

            try
            {
                dataValue = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(value.Split('-')[0]));
                storedHash = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(value.Split('-')[1]));
                calcHash = System.Text.Encoding.UTF8.GetString(mac3des.ComputeHash(System.Text.Encoding.UTF8.GetBytes(dataValue)));
                if (storedHash != calcHash)
                {
                    throw new ArgumentException("Hash value does not match");
                    //This error is immediately caught below
                }
            }
            catch (System.Exception ex)
            {
                throw new ArgumentException("Invalid TamperProofString");
            }

            return dataValue;

        }
        #endregion

    }
}

