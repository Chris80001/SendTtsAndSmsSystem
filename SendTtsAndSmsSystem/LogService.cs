using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;
using DAO = Microsoft.Office.Interop.Access.Dao;
using MySql.Data.MySqlClient;


namespace SendTtsAndSmsSystem
{
    class LogService
    {
        public static DataTable SelectLog(double StartDate, double EndDate, out string ErrorMsg)
        {
            ErrorMsg = string.Empty;
            DataTable dtResult = new DataTable();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                MySqlCommand cmd = connection.CreateCommand();
                StringBuilder commandstring = new StringBuilder();

                
                commandstring.Append("Select N_TMS, N_USERNAME, N_CALLNO, N_RET from notify2 ") ;
                commandstring.Append("Where N_TMS between @StartDate and @EndDate ");
                commandstring.Append("Order by N_TMS ");
                
                cmd.Parameters.AddWithValue("@StartDate", StartDate);
                cmd.Parameters.AddWithValue("@EndDate", EndDate);
                cmd.CommandText = commandstring.ToString();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                
                da.Fill(dtResult);
                connection.Close();
                da.Dispose();
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }

            return dtResult;
        }
    }
}
