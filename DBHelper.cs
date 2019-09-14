using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Xml;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
namespace KlocSqlHelper
{
    public class DBHelper
    {

        #region "Properties"

        private static int Int_ErrorCode
        {
            get { return 5; }
        }

        private static int Int_ErrorMsg
        {
            get { return 500; }
        }

        private static string Str_Col_DefCol
        {
            get { return "DefaultColumn"; }
        }

        private static string Str_Tbl_ResTbl
        {
            get { return "ResultTable"; }
        }

        private static string Str_Ds_ResDs
        {
            get { return "ResultDataSet"; }
        }

        private static string Str_Tbl_ErrTbl
        {
            get { return "ErrorTable"; }
        }

        private static string Str_Col_ErrCode
        {
            get { return "ErrorCode"; }
        }

        private static string Str_Col_ErrMsg
        {
            get { return "ErrorMsg"; }
        }

        private static string Str_Col_ErrLine
        {
            get { return "ErrorLine"; }
        }

        private static string Str_Col_ErrProc
        {
            get { return "ErrorProc"; }
        }

        private static string Str_Col_ErrState
        {
            get { return "ErrorState"; }
        }

        private static string Str_Col_ErrSev
        {
            get { return "ErrorSeverity"; }
        }

        private static string Str_Col_DBName
        {
            get { return "DBName"; }
        }

        private static string Str_Col_ServerName
        {
            get { return "ServerName"; }
        }

        private static string Str_Col_PKey
        {
            get { return "PrimaryKey"; }
        }
        public static string AppSettingCon(string ConString, string XmlParamString)
        {

            string ret = string.Empty;

            DataTable resdt = null;

            DataSet ds = new DataSet();

            string fund = "";

            string connectiondecider = "";

            string ConnectionString = null;

            XmlDocument xmldocument = new XmlDocument();

            xmldocument.LoadXml(XmlParamString);

            ds.ReadXml(new XmlNodeReader(xmldocument));


            if (ConString == "Mfdwebtest")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["Mfdwebtest"];

                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");




            }
            else if (ConString == "Kbolttest")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["Kbolttest"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");

            }

            else if (ConString == "Karvymfstest")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["Karvymfstest"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }

            else if (ConString == "RMF")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }
            else if (ConString == "123")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["123"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }
            else if (ConString == "127")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["127"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }
            else if (ConString == "Reliance")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }

            else if (ConString == "Karvymfsalt")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["Karvymfsalt"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }

            else if (ConString == "KBOLT")
            {
                ConString = System.Configuration.ConfigurationManager.AppSettings["KBOLT"];
                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");
            }


            else
            {
                //If ds.Tables(0).Rows.Count > 0 Then

                //    fund = "501"

                //End If
                ConString = System.Configuration.ConfigurationManager.AppSettings[fund];
                //    ConString = System.Configuration.ConfigurationManager.AppSettings(fund);

                ConnectionString = TamperProofStringDecode(ConString, "KBCONSTR");

            }

            return ConnectionString;

        }
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

        private static string Str_DBCon
        {
            //Return Convert.ToString(ConfigurationManager.ConnectionStrings("WMUDB"))
            //Return Convert.ToString(ConfigurationManager.ConnectionStrings("CitiDB"))
            //Return Convert.ToString(ConfigurationManager.ConnectionStrings("IndiaBullsDB"))
            // Return Convert.ToString(ConfigurationManager.ConnectionStrings("PiramalDB"))
            get { return Convert.ToString(ConfigurationManager.ConnectionStrings["MfdwebTest"]); }
        }

        #endregion

        #region "Public Methods"

        /// <summary>
        /// Used to get a single value as a result
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <returns>Returns the result as DataTable</returns>
        /// <remarks></remarks>
        public static DataTable ExecuteScalar(string ConStr, string SPName, string XmlParamString)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable resdt = null;
            try
            {
                ConStr = Str_DBCon;
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resdt = op;
                    return resdt;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resdt = op;
                        return resdt;
                    }
                    else
                    {
                        op = ValidateParamDt(XmlParamString, "XmlParamString");
                        if (op != null)
                        {
                            resdt = op;
                            return resdt;
                        }
                    }
                }
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Param", XmlParamString);
                da = new SqlDataAdapter(cmd);
                resdt = new DataTable();
                da.Fill(resdt);

                if (!string.IsNullOrWhiteSpace(resdt.Rows[0][Int_ErrorCode].ToString()))
                {
                    resdt.TableName = Str_Tbl_ErrTbl;
                }
                else
                {
                    resdt.TableName = Str_Tbl_ResTbl;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
                if (resdt != null)
                {
                    resdt.Dispose();
                    resdt = null;
                }
            }
            return resdt;
        }

        /// <summary>
        /// Used to get count of rows affected, list of out parameters with corresponding values 
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <param name="CmdTimeOut">Pass the execution timeout(Optional)</param>
        /// <returns>Returns count of rows affected</returns>
        /// <remarks></remarks>
        public static DataTable ExecuteNonQuery(string ConStr, string SPName, string XmlParamString, int? CmdTimeOut = null)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            DataTable res = null;
            //Dim res As Integer = -1
            //Dim op As String = String.Empty
            try
            {
                res = new DataTable(Str_Tbl_ResTbl);
                ConStr = Str_DBCon;
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                //op = ValidateParam(ConStr, "ConStr")
                if (op != null)
                {
                    res = op;
                    return res;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    //op = ValidateParam(SPName, "SPName")
                    if (op != null)
                    {
                        res = op;
                        return res;
                    }
                    else
                    {
                        op = ValidateParamDt(XmlParamString, "XmlParamString");
                        if (op != null)
                        {
                            res = op;
                            return res;
                        }
                    }
                }
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Param", XmlParamString);
                //cmd.Parameters.Add(AddParam("@ErrorCode", Int_ErrorCode))
                //cmd.Parameters.Add(AddParam("@ErrorMsg", Int_ErrorMsg))
                con.Open();
                if (CmdTimeOut > 30)
                {
                    cmd.CommandTimeout = Convert.ToInt32(CmdTimeOut);
                }
                dynamic read = cmd.ExecuteReader();
                if (read.HasRows)
                {
                    DataTable dt = default(DataTable);
                    while (read.Read())
                    {
                        if (!string.IsNullOrWhiteSpace(read.Item(Str_Col_ErrCode)))
                        {
                            dt = GetTable();
                        }
                        else
                        {
                            dt = GetTable(false);
                        }
                        if (!string.IsNullOrWhiteSpace(read.Item(Str_Col_ErrCode)))
                        {
                            dt.Rows.Add(Convert.ToString(read.Item(Str_Col_ErrCode)), Convert.ToString(read.Item(Str_Col_ErrMsg)), Convert.ToString(read.Item(Str_Col_ErrLine)), Convert.ToString(read.Item(Str_Col_ErrProc)), Convert.ToString(read.Item(Str_Col_ErrState)), Convert.ToString(read.Item(Str_Col_ErrSev)), Convert.ToString(read.Item(Str_Col_DBName)), Convert.ToString(read.Item(Str_Col_ServerName)));
                        }
                        else
                        {
                            dt.Rows.Add(Convert.ToString(read.Item(Str_Col_ErrCode)), Convert.ToString(read.Item(Str_Col_PKey)));
                        }
                    }
                    read.Close();
                    res = dt;
                }

                //Purpose: For Integer result only
                //res = cmd.ExecuteNonQuery()
                //Dim intop = cmd.ExecuteNonQuery()
                //res.Columns.Add("Param")
                //res.Columns.Add("Result")
                //'If intop > 0 Then
                //res.Rows.Add("@RowsAffected", "(" + Convert.ToString(intop) + " row(s) affected)")
                //'End If
                //res.Rows.Add("@ErrorCode", Convert.ToString(cmd.Parameters("@ErrorCode").Value))
                //res.Rows.Add("@ErrorMsg", Convert.ToString(cmd.Parameters("@ErrorMsg").Value))
                //Catch ex As Exception
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
            }
            return res;
        }

        /// <summary>
        /// Used to get DataTable
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <returns>Returns DataTable</returns>
        /// <remarks></remarks>
        public static DataTable ExecuteSP_GetDataTable(string ConStr, string SPName)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable resdt = null;
            try
            {
                resdt = new DataTable(Str_Tbl_ResTbl);
                ConStr = Str_DBCon;
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resdt = op;
                    return resdt;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resdt = op;
                        return resdt;
                    }
                }
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                da = new SqlDataAdapter(cmd);
                da.Fill(resdt);
                if (resdt.Rows.Count > 0)
                {
                    if (!string.IsNullOrWhiteSpace(resdt.Rows[0][Str_Col_ErrCode].ToString()))
                    {
                        resdt.TableName = Str_Tbl_ErrTbl;
                    }
                    else
                    {
                        resdt.TableName = Str_Tbl_ResTbl;
                    }
                }
                else
                {
                    resdt.TableName = Str_Tbl_ResTbl;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }

        /// <summary>
        /// Used to get DataTable
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <returns>Returns DataTable</returns>
        /// <remarks></remarks>
        public static DataTable ExecuteSP_GetDataTable(string ConStr, string SPName, string XmlParamString)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;

            SqlDataAdapter da = null;
            DataTable resdt = null;
            try
            {
                resdt = new DataTable(Str_Tbl_ResTbl);
                ConStr = Str_DBCon;
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resdt = op;
                    return resdt;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resdt = op;
                        return resdt;
                    }
                    else
                    {
                        op = ValidateParamDt(XmlParamString, "XmlParamString");
                        if (op != null)
                        {
                            resdt = op;
                            return resdt;
                        }
                    }
                }
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Param", XmlParamString);
                da = new SqlDataAdapter(cmd);
                da.Fill(resdt);

                //If resdt.Rows.Count > 0 Then
                //    If Not String.IsNullOrWhiteSpace(resdt.Rows(0)(Str_Col_ErrCode)) Then
                //        resdt.TableName = Str_Tbl_ErrTbl
                //    Else
                //        resdt.TableName = Str_Tbl_ResTbl
                //    End If
                //Else
                //    resdt.TableName = Str_Tbl_ResTbl
                //End If

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }

        /// <summary>
        /// Used to get DataTable
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <param name="WithFlag">Pass boolean 'false' if 'ErrorCode' column is not needed</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DataTable ExecuteSP_GetDataTable(string ConStr, string SPName, string XmlParamString, bool WithFlag = true)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable resdt = null;
            try
            {
                resdt = new DataTable(Str_Tbl_ResTbl);
                ConStr = Str_DBCon;
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resdt = op;
                    return resdt;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resdt = op;
                        return resdt;
                    }
                    else
                    {
                        op = ValidateParamDt(XmlParamString, "XmlParamString");
                        if (op != null)
                        {
                            resdt = op;
                            return resdt;
                        }
                    }
                }
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Param", XmlParamString);
                da = new SqlDataAdapter(cmd);
                da.Fill(resdt);
                if (resdt.Columns.Contains(Str_Col_ErrCode))
                {
                    resdt.TableName = Str_Tbl_ErrTbl;
                }
                else
                {
                    resdt.TableName = Str_Tbl_ResTbl;
                }
                //If resdt.Rows.Count > 0 Then
                //    If Not String.IsNullOrWhiteSpace(resdt.Rows(0)(Str_Col_ErrCode)) Then
                //        resdt.TableName = Str_Tbl_ErrTbl
                //    Else
                //        resdt.TableName = Str_Tbl_ResTbl
                //    End If
                //Else
                //    resdt.TableName = Str_Tbl_ResTbl
                //End If
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }

        //============================================================================================================================================

        public static DataSet ExecuteMastertablesSP_And_GetDataSet(string ConStr, string Tablename, string DMLType, string xmldata)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataSet resds = null;
            string SPName = "P_MasterBuild";
            SqlParameter outp1 = null;
            SqlParameter outp2 = null;
            ConStr = Str_DBCon;

            try
            {
                resds = new DataSet();

                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;


                cmd.Parameters.AddWithValue("@Tablename", Tablename);
                cmd.Parameters.AddWithValue("@DMLType", DMLType);
                cmd.Parameters.AddWithValue("@xml", xmldata);

                outp1 = new SqlParameter("@ErrorNo", SqlDbType.Int);
                outp1.Direction = ParameterDirection.Output;

                outp2 = new SqlParameter("@StatusMsg", SqlDbType.VarChar, 100);
                outp2.Direction = ParameterDirection.Output;

                cmd.Parameters.Add(outp1);
                cmd.Parameters.Add(outp2);


                da = new SqlDataAdapter(cmd);
                resds = new DataSet();
                da.Fill(resds);

            }
            catch (Exception ex)
            {
                resds = null;
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
                outp1 = null;
                outp2 = null;

            }
            return resds;
        }
        //============================================================================================================================================
        public static DataTable ExecuteNonQueryByDMLType(string ConStr, string SPName, string DMLType, string xmldata)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable resdt = null;
            try
            {
                ConStr = Str_DBCon;
                resdt = new DataTable("ResultTable");

                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DMLType", DMLType);
                cmd.Parameters.AddWithValue("@Param", xmldata);
                da = new SqlDataAdapter(cmd);
                resdt = new DataTable("ResultTable");
                da.Fill(resdt);

            }
            catch (Exception ex)
            {
                resdt = null;
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }
        public static DataTable ExecuteNonQueryforSMDeal(string ConStr, string SPName, string DMLType, string xmldata, string TblName)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataTable resdt = null;
            SqlParameter outp1 = null;
            SqlParameter outp2 = null;
            SqlParameter outp3 = null;
            try
            {
                ConStr = Str_DBCon;
                resdt = new DataTable("ResultTable");

                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Tablename", TblName);
                cmd.Parameters.AddWithValue("@DMLType", DMLType);
                cmd.Parameters.AddWithValue("@xml", xmldata);
                outp1 = new SqlParameter("@ErrorNo", SqlDbType.Int);
                outp1.Direction = ParameterDirection.Output;
                outp2 = new SqlParameter("@StatusMsg", SqlDbType.VarChar, 100);
                outp2.Direction = ParameterDirection.Output;
                outp3 = new SqlParameter("@TranID_Output", SqlDbType.VarChar, 100);
                outp3.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(outp1);
                cmd.Parameters.Add(outp2);
                cmd.Parameters.Add(outp3);
                da = new SqlDataAdapter(cmd);
                resdt = new DataTable("ResultTable");
                da.Fill(resdt);
                resdt.Columns.Add("ErrorNo");
                resdt.Columns.Add("StatusMsg");
                resdt.Columns.Add("TranID_Output");
                resdt.Rows.Add(outp1.Value.ToString(), outp2.Value.ToString(), outp3.Value.ToString());


            }
            catch (Exception ex)
            {
                resdt = null;
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }
        public static DataSet ExecuteNonQueryForAuthorized(string ConStr, string TableName, string xmldata)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataSet resdt = null;
            SqlParameter outp1 = null;
            SqlParameter outp2 = null;
            try
            {
                ConStr = Str_DBCon;
                resdt = new DataSet();
                con = new SqlConnection(ConStr);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "P_Bo_Authorised";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Param", xmldata);
                cmd.Parameters.AddWithValue("@TableName", TableName);
                outp1 = new SqlParameter("@ErrorNo", SqlDbType.Int);
                outp1.Direction = ParameterDirection.Output;
                outp2 = new SqlParameter("@StatusMsg", SqlDbType.VarChar, 100);
                outp2.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(outp1);
                cmd.Parameters.Add(outp2);
                da = new SqlDataAdapter(cmd);
                da.Fill(resdt);
            }
            catch (Exception ex)
            {
                resdt = null;
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resdt;
        }


        //=============================================================================================================================================

        /// <summary>
        /// Used to get DataSet
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <returns>Returns DataSet</returns>
        /// <remarks></remarks>
        public static DataSet ExecuteSP_GetDataSet(string ConStr, string SPName)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataSet resds = null;
            string constring = null;
            string ConnectionString = null;
            try
            {
                resds = new DataSet(Str_Tbl_ResTbl);
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resds.Tables.Add(op);
                    return resds;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resds.Tables.Add(op);
                        return resds;
                    }

                }
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
                    constring = System.Configuration.ConfigurationManager.AppSettings["123"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }


                else if (ConStr == "Reliance")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Reliance"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }

                else if (ConStr == "Karvymfsalt")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Karvymfsalt"];
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
                else if (ConStr == "axistest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["axistest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }


                con = new SqlConnection(ConnectionString);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;

                da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 500;
                da.SelectCommand.CommandTimeout = 1000;
                resds = new DataSet();
                da.Fill(resds);


                if (resds.Tables.Count > 0 && resds.Tables[0].Rows.Count > 0)
                {
                    if (resds.Tables[0].Columns.Contains(Str_Col_ErrCode))
                    {
                        if (!string.IsNullOrWhiteSpace(resds.Tables[0].Rows[0][Str_Col_ErrCode].ToString()))
                        {
                            resds.Tables[0].TableName = Str_Tbl_ErrTbl;
                        }
                        else
                        {
                            resds.DataSetName = Str_Ds_ResDs;
                        }
                    }
                    else
                    {
                        resds.DataSetName = Str_Ds_ResDs;
                    }

                    //If Not String.IsNullOrWhiteSpace(resds.Tables(0).Rows(0)(Str_Col_ErrCode)) Then
                    //    resds.Tables(0).TableName = Str_Tbl_ErrTbl
                    //Else
                    //    resds.DataSetName = Str_Ds_ResDs
                    //End If
                }
                else
                {
                    resds.DataSetName = Str_Ds_ResDs;
                }
                //If Not String.IsNullOrWhiteSpace(resds.Tables(0).Rows(0)(Str_Col_ErrCode)) Then
                //    resds.Tables(0).TableName = Str_Tbl_ErrTbl
                //Else
                //    resds.DataSetName = Str_Ds_ResDs
                //End If
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resds;
        }





        /// <summary>
        /// Used to get DataSet
        /// </summary>
        /// <param name="ConStr">Pass the DB Connection string</param>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <returns>Returns DataSet</returns>
        /// <remarks></remarks>
        public static DataSet ExecuteSP_GetDataSet(string ConStr, string SPName, string XmlParamString)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            DataSet resds = null;
            string constring = null;
            string ConnectionString = null;
            try
            {
                resds = new DataSet(Str_Tbl_ResTbl);
                dynamic op = ValidateParamDt(ConStr, "ConStr");
                if (op != null)
                {
                    resds.Tables.Add(op);
                    return resds;
                }
                else
                {
                    op = ValidateParamDt(SPName, "SPName");
                    if (op != null)
                    {
                        resds.Tables.Add(op);
                        return resds;
                    }
                    else
                    {
                        op = ValidateParamDt(XmlParamString, "XmlParamString");
                        if (op != null)
                        {
                            resds.Tables.Add(op);
                            return resds;
                        }
                    }
                }
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



                else if (ConStr == "axistest")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["axistest"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }


                con = new SqlConnection(ConnectionString);
                cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = SPName;
                cmd.CommandType = CommandType.StoredProcedure;
                if (!string.IsNullOrWhiteSpace(XmlParamString))
                {
                    cmd.Parameters.AddWithValue("@Param", XmlParamString);
                }
                da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 500;
                da.SelectCommand.CommandTimeout = 1000;
                resds = new DataSet();
                da.Fill(resds);


                if (resds.Tables.Count > 0 && resds.Tables[0].Rows.Count > 0)
                {
                    if (resds.Tables[0].Columns.Contains(Str_Col_ErrCode))
                    {
                        if (!string.IsNullOrWhiteSpace(resds.Tables[0].Rows[0][Str_Col_ErrCode].ToString()))
                        {
                            resds.Tables[0].TableName = Str_Tbl_ErrTbl;
                        }
                        else
                        {
                            resds.DataSetName = Str_Ds_ResDs;
                        }
                    }
                    else
                    {
                        resds.DataSetName = Str_Ds_ResDs;
                    }

                    //If Not String.IsNullOrWhiteSpace(resds.Tables(0).Rows(0)(Str_Col_ErrCode)) Then
                    //    resds.Tables(0).TableName = Str_Tbl_ErrTbl
                    //Else
                    //    resds.DataSetName = Str_Ds_ResDs
                    //End If
                }
                else
                {
                    resds.DataSetName = Str_Ds_ResDs;
                }
                //If Not String.IsNullOrWhiteSpace(resds.Tables(0).Rows(0)(Str_Col_ErrCode)) Then
                //    resds.Tables(0).TableName = Str_Tbl_ErrTbl
                //Else
                //    resds.DataSetName = Str_Ds_ResDs
                //End If
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
            return resds;
        }
        public static DataSet ExecuteDataSet(String SPName, List<SqlParameter> paramlist, string Dbconn, string type)
        {
            string constring = string.Empty;
            string ConnectionString = string.Empty;
            if (type != "CriticalSms")
            {
                if (Dbconn == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["RMF"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
            }
            if (type == "CriticalSms")
            {
                if (Dbconn == "RMF")
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Rmfcom"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
                else
                {
                    constring = System.Configuration.ConfigurationManager.AppSettings["Communicationlog"];
                    ConnectionString = TamperProofStringDecode(constring, "KBCONSTR");
                }
            }

            DataSet ds = null;
            try
            {
                using (SqlConnection con = new SqlConnection())
                {
                    con.ConnectionString = ConnectionString;
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = SPName;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Connection = con;
                        if (paramlist != null && paramlist.Count > 0)
                        {
                            foreach (SqlParameter p in paramlist)
                            {
                                cmd.Parameters.Add(p);

                                //cmd.Parameters.AddRange(paramlist.ToArray());
                            }
                        }
                        using (SqlDataAdapter da = new SqlDataAdapter())
                        {
                            da.SelectCommand = cmd;
                            ds = new DataSet();
                            da.Fill(ds);
                            ds.DataSetName = "Dataset1";
                            // KBSQLDataSet.returnedDataSet = ds;
                        }
                    }
                }
            }

            catch (Exception ex)
            {

            }
          
            return ds;


        }


        /// <summary>
        /// Used to get SQL string which can be executed directly in SQL Query Window
        /// </summary>
        /// <param name="SPName">Pass the SP Name</param>
        /// <param name="XmlParamString">Pass the serialized xml parameter string</param>
        /// <returns>Returns the final SQL string</returns>
        /// <remarks></remarks>

        public static string GetSqlQuery(string SPName, string XmlParamString)
        {
            string res = string.Empty;
            if (!string.IsNullOrWhiteSpace(SPName) & !string.IsNullOrWhiteSpace(XmlParamString))
            {
                dynamic temp = XmlParamString.Replace("\"", "'");
                if (temp.Contains("<?xml version='1.0' encoding='utf-16'?>"))
                {
                    temp = temp.Replace("<?xml version='1.0' encoding='utf-16'?>", "");
                    res = "EXEC [" + SPName + "] @Param = '" + temp + "'";
                }
                else
                {
                    res = "EXEC [" + SPName + "] @Param = '" + XmlParamString + "'";
                }
            }
            return res;
        }

        #endregion
        #region "Private Methods"

        private static DataTable GetTable(bool IsError = true)
        {
            DataTable dt = null;
            if (!IsError)
            {
                dt = new DataTable(Str_Tbl_ResTbl);
                dt.Columns.Add(Str_Col_ErrCode);
                dt.Columns.Add(Str_Col_PKey);
            }
            else
            {
                dt = new DataTable(Str_Tbl_ErrTbl);
                dt.Columns.Add(Str_Col_ErrCode);
                dt.Columns.Add(Str_Col_ErrMsg);
                dt.Columns.Add(Str_Col_ErrLine);
                dt.Columns.Add(Str_Col_ErrProc);
                dt.Columns.Add(Str_Col_ErrState);
                dt.Columns.Add(Str_Col_ErrSev);
                dt.Columns.Add(Str_Col_DBName);
                dt.Columns.Add(Str_Col_ServerName);
            }
            return dt;
        }

        private static DataTable ValidateParamDt(string Param, string ParamName)
        {
            DataTable dt = null;
            if (string.IsNullOrWhiteSpace(Param))
            {
                dt = GetTable();
                dt.Rows.Add("1001", "Parameter '" + ParamName + "' should not be null or empty");
            }
            return dt;
        }

        private static DataTable ValidateParamDt<T>(List<T> LstParam, string LstName)
        {
            DataTable dt = null;
            if (LstParam == null)
            {
                dt = GetTable();
                dt.Rows.Add("1002", "List '" + LstName + "' should not be null or empty");
            }
            else if (LstParam.Count <= 0)
            {
                dt = GetTable();
                dt.Rows.Add("1003", "No Items in the list parameter '" + LstName + "'");
            }
            return dt;
        }

        #endregion
        #region "NAV Uploads "
        //public static DataTable ReadExcelFile(string ConStr, string Src, string UploadFileName)
        //{
        //    DataTable dt = null;
        //    System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection();
        //    string getExcelSheetName = "";
        //    string Query = "";
        //    string connString = "";
        //    if ((UploadFileName.Substring(UploadFileName.LastIndexOf(".") + 1, UploadFileName.Length - Conversion.Val((UploadFileName.LastIndexOf(".") + 1)).ToUpper == "CSV")
        // {
        //        connString = (Convert.ToString("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=") + System.IO.Path.GetDirectoryName(Src)) + ";Extended Properties=\"Text;HDR=YES;FMT=Delimited\"";
        //    }
        //    else
        //    {
        //        connString = (Convert.ToString("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=") + Src) + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
        //    }

        //    MyConnection = new System.Data.OleDb.OleDbConnection(connString);

        //    MyConnection.Open();

        //    System.Data.DataTable dtExcelSheetName = null;
        //    dtExcelSheetName = MyConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables_Info, null);
        //    //getExcelSheetName = dtExcelSheetName.Rows(0)("Table_Name").ToString()


        //    if (UploadFileName.Substring(UploadFileName.LastIndexOf(".") + 1, UploadFileName.Length - Conversion.Val((UploadFileName.LastIndexOf(".")) + 1)).ToUpper == "CSV")
        //    {
        //        Query = string.Format("select * from [{0}]", UploadFileName);
        //    }
        //    else
        //    {
        //        if (dtExcelSheetName.Rows(0)("Table_Name").ToString().Contains('$'))
        //        {
        //            getExcelSheetName = dtExcelSheetName.Rows[0]["Table_Name"].ToString();
        //        }
        //        else
        //        {
        //            getExcelSheetName = dtExcelSheetName.Rows[1]["Table_Name"].ToString();
        //        }
        //        Query = string.Format("select * from [{0}]", getExcelSheetName);
        //    }

        //    System.Data.OleDb.OleDbDataAdapter data = new System.Data.OleDb.OleDbDataAdapter();
        //    data.SelectCommand = new System.Data.OleDb.OleDbCommand(Query, MyConnection);
        //    data.SelectCommand.CommandTimeout = 3000;
        //    dt = new DataTable();
        //    data.Fill(dt);
        //    MyConnection.Close();
        //    return dt;

        //}

        //' <summary>
        /// Used to Insert Bulk Data to Table
        /// </summary>
        /// <param name="STable">Pass the Table Name</param>
        /// <returns>returns the result string</returns>
        /// <remarks></remarks>

        public static void Execute_SQLBulkCopy(string ConStr, DataTable STable, string DTable)
        {
            SqlConnection con = null;
            SqlCommand cmd = null;
            SqlDataAdapter da = null;
            string fund = "";
            DataSet resds = null;
            string constring = null;
            ConStr = Str_DBCon;
            SqlBulkCopy bulkCopy = new SqlBulkCopy(ConStr);
            try
            {
                bulkCopy.DestinationTableName = DTable;
                foreach (DataColumn col in STable.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }
                bulkCopy.WriteToServer(STable);


            }
            catch (SqlException ex)
            {
                if (ex.Message.Contains("Received an invalid column length from the bcp client for colid"))
                {
                    string pattern = "\\d+";
                    Match match = Regex.Match(ex.Message.ToString(), pattern);
                    dynamic index = Convert.ToInt32(match.Value) - 1;

                    FieldInfo fi = typeof(SqlBulkCopy).GetField("_sortedColumnMappings", BindingFlags.NonPublic | BindingFlags.Instance);
                    dynamic sortedColumns = fi.GetValue(bulkCopy);
                    dynamic items = (Object[])sortedColumns.GetType().GetField("_items", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(sortedColumns);

                    FieldInfo itemdata = items(index).GetType().GetField("_metadata", BindingFlags.NonPublic | BindingFlags.Instance);
                    dynamic metadata = itemdata.GetValue(items(index));

                    dynamic column = metadata.GetType().GetField("column", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                    dynamic length = metadata.GetType().GetField("length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                    throw new FormatException(String.Format("Column: {0} contains data with a length greater than: {1}", column, length));
                }

                throw;



            }
            finally
            {
                if (con != null)
                {
                    if (con.State == ConnectionState.Open)
                        con.Close();
                    con.Dispose();
                    con = null;
                }
                if (cmd != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                    cmd = null;
                }
                if (da != null)
                {
                    da.Dispose();
                    da = null;
                }
            }
        }
        #endregion
    }
}
