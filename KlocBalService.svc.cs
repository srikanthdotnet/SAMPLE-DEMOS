using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using KlocBal.KlocDalService;
using KlocModel;
using System.Data;
using KlocBal;
using System.IO;
using System.ServiceModel.Web;
using System.Net;
namespace KlocBal
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "KlocBalService" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select KlocBalService.svc or KlocBalService.svc.cs at the Solution Explorer and start debugging.
    public class KlocBalService : IKlocBalService
    {
        string defMsg = "";
        public void DoWork()
        {
        }
        #region GetAllFunds
        public KlocModel.CommonReturnType GetAllFunds(string userid)
        {
            IncomingWebRequestContext request = WebOperationContext.Current.IncomingRequest;
            WebHeaderCollection headers = request.Headers;
            
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {
                string xmldata = "";
                NexPurchase obj = new NexPurchase();
                obj.UserID = userid;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetAllFunds(xmldata);
                    cmd.Status = true;
                }


            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }

        #endregion

        #region Getallsipdashboarddata
        public KlocModel.CommonReturnType Getallsipdashboarddata(string fund, string flg, string Fromdt, string Todate)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                

                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.Getallsipdashboarddata(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

        #region GetChannelSIPDashboard
        public CommonReturnType GetChannelSIPDashboard(string fund, string flg, string Fromdt, string Todate)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetChannelSIPDashboard(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;

        }
        #endregion

        #region GetExchangeSIPDashboard
        public CommonReturnType GetExchangeSIPDashboard(string fund, string flg, string Fromdt, string Todate)
        {

            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {

                    dtdata = dalObjecet.GetExchangeSIPDashboard(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;

        }
        #endregion

        #region GetISIPDashboard
        public CommonReturnType GetISIPDashboard(string fund, string flg, string Fromdt, string Todate)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {

                    dtdata = dalObjecet.GetISIPDashboard(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

        #region GetNCTDashboard
        public CommonReturnType GetNCTDashboard(string fund, string flg, string Fromdt, string Todate)
        {

            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {

                    dtdata = dalObjecet.GetNCTDashboard(xmldata, flg, fund, Fromdt, Todate);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

        #region GetBANKINGDashboard
        public CommonReturnType GetBANKINGDashboard(string fund, string Fromdt)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.Fromdt = Fromdt;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetBANKINGDashboard(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

        #region GetBANKINGDashboardExportExcelData
        public CommonReturnType GetBANKINGDashboardExportExcelData(string fund, string Fromdt, string flg)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.Fromdt = Fromdt;
                obj.flg = flg;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetBANKINGDashboardExportExcelData(xmldata, fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;




        }
        #endregion

        #region GetBankingDashboardScheduler
        public DataSet GetBankingDashboardScheduler(string fund)
        {

            DataSet dtdata = new DataSet();
            try
            {
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetBankingDashboardScheduler(fund);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        #endregion

        #region GetDcrDashboardScheduler
        public DataSet GetDcrDashboardScheduler(string fundcode, string flg)
        {

            DataSet dtdata = new DataSet();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.flg = flg;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetDcrDashboardScheduler(fundcode, flg, xmldata);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;


        }
        #endregion

        #region GetExchangeDashboard
        public DataSet GetExchangeDashboard(string fundcode, string flg, string Todate)
        {
            DataSet dtdata = new DataSet();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.flg = flg;
                obj.Todate = Todate;
                obj.fund = fundcode;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetExchangeDashboard(xmldata, fundcode);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        public DataSet GetSmsMis(string param)
        {
            DataSet dtdata = new DataSet();
            try
            {
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetSmsMis(param);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        public DataSet GetCriticalSms(string param)
        {
            DataSet dtdata = new DataSet();
            try
            {
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetCriticalSms(param);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        public DataSet GetExchageofflineDetails(string fundcode, string flg, string Todate, string FileName)
        {
            DataSet dtdata = new DataSet();
            try
            {
                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.flg = flg;
                obj.Todate = Todate;
                obj.fund = fundcode;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetExchageofflineDetails(xmldata, fundcode, FileName);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        public DataSet GetExchangeofflineCA(string fundcode)
        {
            DataSet dtdata = new DataSet();
            try
            {
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetExchangeofflineCA(fundcode);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }

        #endregion

        #region GetSIPDashboardScheduler
        public DataSet GetSIPDashboardScheduler(string fund)
        {
            DataSet dtdata = new DataSet();
            try
            {
                CommonReturnType cmd = new CommonReturnType();
                KlocDalServiceClient dalObjecet = new KlocDalServiceClient();
                dtdata = dalObjecet.GetSIPDashboardScheduler(fund);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }

            return dtdata;
        }
        #endregion

        #region PaytmDashboard
        public CommonReturnType GetPaytmDashboard(string fund, string flg, string Fromdt, string Todate, string FileName)
        {

            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetPaytmDashboard(xmldata, fund, FileName);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

        public CommonReturnType GetCPDDashboard(string fund, string flg, string Fromdt, string Todate, string FileName)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetCPDDashboard(xmldata, fund, FileName);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        public CommonReturnType GetCPDDashboardDetails(string fund, string flg, string Fromdt, string Todate, string Mode, string Remarks, string FileName)
        {
            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                obj.Mode = Mode;
                obj.Remarks = Remarks;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetCPDDashboardDetails(xmldata, fund, FileName);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }

        #region FundingandPayoutDashboard
        public CommonReturnType GetFundingandPayoutDashboard(string fund, string flg, string Fromdt, string Todate)
        {

            DataSet dtdata = null;
            CommonReturnType cmd = new CommonReturnType();
            try
            {

                string xmldata = "";
                Sipdashboard obj = Sipdashboard.GetInstance;
                obj.fund = fund;
                obj.flg = flg;
                obj.Fromdt = Fromdt;
                obj.Todate = Todate;
                xmldata = BllCommonUtility.SerializeToXml(obj);
                using (KlocDalServiceClient dalObjecet = new KlocDalServiceClient())
                {
                    dtdata = dalObjecet.GetFundingandPayoutDashboard(xmldata,fund);
                    cmd.Status = true;
                }
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                cmd.Status = false;
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
                cmd.JSONData = cmd.Serialize_JsonData(dtdata);
                cmd.ds = dtdata;
            }
            return cmd;
        }
        #endregion

    }
}
