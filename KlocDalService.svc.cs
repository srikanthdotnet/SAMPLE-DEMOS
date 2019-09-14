using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.ServiceModel.Web;
using System.Data;
using KlocSqlHelper;
using System.Data.SqlClient;
namespace KlocDal
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "KlocDalService" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select KlocDalService.svc or KlocDalService.svc.cs at the Solution Explorer and start debugging.
    public class KlocDalService : IKlocDalService
    {
        #region  VariableDeclartion
        public static string KBOLT = "KBOLT";
        // public static string axismf = "128";
        static string axitest = "axistest";
        string defMsg = "";
        #endregion

        #region GetAllFunds
        public DataSet GetAllFunds(string param)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(KBOLT, "knex_getfundsbyuser", param);
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


        #region GetSMSMIS
        public DataSet GetSmsMis(string param)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                List<SqlParameter> plist = new List<SqlParameter>();
                SqlParameter p;
                p = new SqlParameter("@fund", SqlDbType.VarChar, 50);
                p.Value = Convert.ToString(param);
                plist.Add(p);
               
                dtdata = DBHelper.ExecuteDataSet("SMS_MIS_Triggerstatus_Proc",plist, param,"");
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

        #region GetCriticalSms
        public DataSet GetCriticalSms(string param)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                List<SqlParameter> plist = new List<SqlParameter>();
                SqlParameter p;
                p = new SqlParameter("@i_Fund", SqlDbType.VarChar, 50);
                p.Value = Convert.ToString(param);
                plist.Add(p);
                dtdata = DBHelper.ExecuteDataSet("SMS_MIS_GetCriticalData", plist, param, "CriticalSms");
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
        #region Getallsipdashboarddata
        public DataSet Getallsipdashboarddata(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "SIPRegistrationDashboarddata_FINALOUTPUT_OLD", param);
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

        #region GetChannelSIPDashboard
        public DataSet GetChannelSIPDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "ChannelSIPDashboard_xml", param);
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

        #region GetExchangeSIPDashboard

        public DataSet GetExchangeSIPDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "EXCHANGESIPDashboard_xml", param);
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

        #region  GetISIPDashboard
        public DataSet GetISIPDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "ISIPDashboard_xml", param);
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

        #region GetNCTDashboard
        public DataSet GetNCTDashboard(string param, string flg, string fundcode, string Fromdt, string Todate)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "Nct_Dashboard_Report_test_xml", param);

            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            finally
            {
            }
            return dtdata;
        }

        #endregion

        #region GetExchangeDashboard
        public DataSet GetExchangeDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();

                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "NSE_NSE_CREDIT_MAPPING_FILE_KLOC_VER1", param);

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

        #region GetBANKINGDashboard
        public DataSet GetBANKINGDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "CrdtNotIdenUpload_BANKINGDASHBOARD_xml", param);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }


        public DataSet GetBANKINGDashboardExportExcelData(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "CrdtNotIdenUpload_BANKINGDASHBOARD_xml", param);
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

        #region GetBankingDashboardScheduler

        public DataSet GetBankingDashboardScheduler(string fundcode)
        {

            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "CrdtNotIdenUpload_BANKINGDASHBOARD_Scheduler_klocAllData");
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
        public DataSet GetSIPDashboardScheduler(string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "Kloc_GetAllsipSchedulertabledata");
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

        #region GetDCRDashboardScheduler

        public DataSet GetDcrDashboardScheduler(string fundcode, string flg, string param)
        {

            DataSet dtdata = null;
            try
            {
                if (flg == "0")
                {
                    dtdata = new DataSet();
                    dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "klock_DCRGetdata");
                }
                else
                {
                    dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "Kloc_DCR_Details_DailyControlReport", param);

                }
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

        #region GetNCTClientDashboard
        public DataSet GetNCTClientDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "CrdtNotIdenUpload_BANKINGDASHBOARD_xml", param);
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

        #region CPDdashboard

        public DataSet GetCPDDashboard(string param, string fundcode, string FileName)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "GetCPD_DashBorad_kloc", param);
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;
        }
        public DataSet GetCPDDashboardDetails(string param, string fundcode, string FileName)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "GetCPD_DashBorad_ExcelReport_kloc", param);

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

        #region GetExchageofflineDetails
        public DataSet GetExchangeofflineCA(string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "KLOC_DPVALIDATION_STAGE_DATA");
            }
            catch (Exception ex)
            {
                Util.WriteLog(ex.Source, ex.Message, ex.StackTrace);
                defMsg = "An Error Accoured While Data Processing!";
                dtdata = Util.GetErrorcode("1", defMsg);
            }
            return dtdata;



        }
        public DataSet GetExchageofflineDetails(string param, string fundcode, string FileName)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "KLOC_DPVALIDATION_STAGE_DATA_MIS", param);

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

        #region paytmdashboard

        public DataSet GetPaytmDashboard(string param, string fundcode, string FileName)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                if (FileName == "PaytmDashboard")
                {
                    dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "GetPAYTMtransactionReport_kloc", param);
                }
                else
                {
                    dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "GetPAYTMPendingtransactionReport_kloc", param);
                }
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

        #region FundingandPayoutDashboard

        public DataSet GetFundingandPayoutDashboard(string param, string fundcode)
        {
            DataSet dtdata = null;
            try
            {
                dtdata = new DataSet();
                dtdata = DBHelper.ExecuteSP_GetDataSet(fundcode, "FundingandPayoutDashboard", param);
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
    }
}
//ChannelSIPDashboard_Scheduler   
//EXCHANGESIPDashboard_Scheduler  
//ISIPDashboard_Scheduler
//KLOConlinesipdashboard  ---call this procedure only 
//Kloc_Deltetabledataforscheduler
//NSE_NSE_CREDIT_MAPPING_FILE_KLOC
