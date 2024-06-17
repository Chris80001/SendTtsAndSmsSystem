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
using System.ComponentModel;
using System.Data.SqlClient;


namespace SendTtsAndSmsSystem
{
    class RecordingService
    {
        public static DataTable SelectRecording(out string ErrorMsg)
        {
            ErrorMsg = string.Empty;
            DataTable dtResult = new DataTable();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                MySqlCommand cmd = connection.CreateCommand();
                StringBuilder commandstring = new StringBuilder();

                commandstring.Append("Select * from phrase ") ;
                commandstring.Append("Order by ph_no ");
                
                cmd.CommandText = commandstring.ToString();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                
                da.Fill(dtResult);
                connection.Close();
                da.Dispose();
            }
            catch (MySqlException ex)
            {
                int error = ex.Number;
                if (error == 1042) { ErrorMsg = "目前無法正常連線MySQL Database\n"; }
                if (ErrorMsg == string.Empty) { ErrorMsg = ex.Message; }
            }

            return dtResult;
        }

        public static DataTable SelectRecording(string RecordId)
        {
            DataTable dtResult = new DataTable();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                MySqlCommand cmd = connection.CreateCommand();
                StringBuilder commandstring = new StringBuilder();

                commandstring.Append("Select * from phrase ");
                commandstring.Append("WHERE ph_no = @ph_no ");
                commandstring.Append("Order by ph_no ");

                cmd.Parameters.AddWithValue("@ph_no", RecordId);
                cmd.CommandText = commandstring.ToString();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);

                da.Fill(dtResult);
                connection.Close();
                da.Dispose();
            }
            catch (MySqlException ex)
            {
                string errorMsg = ex.Message;
            }

            return dtResult;
        }

        public static string InsertRecording(string ph_no, string ph_data)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                using (MySqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO phrase ");
                    sql.Append("( ph_no, ph_data2 ) ");
                    sql.Append("VALUES( @ph_no, @ph_data2 ) ");

                    cmd.Parameters.AddWithValue("@ph_no", ph_no);
                    cmd.Parameters.AddWithValue("@ph_data2", ph_data);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (MySqlException ex)
            {
                int error = ex.Number;
                if (error == 1042) { ErrorMsg = "目前無法正常連線MySQL Database\n"; }
                if (error == 1062) { ErrorMsg = "「片語代碼」不可以重複\n"; }
                if (ErrorMsg == string.Empty) { ErrorMsg = ex.Message; }
            }
            return ErrorMsg;
        }

        public static string UpdateRecording(string ph_no, string ph_data)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                using (MySqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("UPDATE phrase ");
                    sql.Append("SET ph_data = @ph_data, ph_data2 = @ph_data2 ");
                    sql.Append("WHERE ph_no = @ph_no ");

                    cmd.Parameters.AddWithValue("@ph_no", ph_no);
                    cmd.Parameters.AddWithValue("@ph_data", string.Empty);
                    cmd.Parameters.AddWithValue("@ph_data2", ph_data);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (MySqlException ex)
            {
                int error = ex.Number;
                if (error == 1042) { ErrorMsg = "目前無法正常連線MySQL Database\n"; }
                if (ErrorMsg == string.Empty) { ErrorMsg = ex.Message; }
            }
            return ErrorMsg;
        }

        public static string DeleteRecording(string ph_no)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                MySqlConnection connection = new MySqlConnection(ConfigurationManager.AppSettings["MysqlConnString"]);
                connection.Open();
                using (MySqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE FROM phrase ");
                    sql.Append("WHERE ph_no = @ph_no ");

                    cmd.Parameters.AddWithValue("@ph_no", ph_no);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (MySqlException ex)
            {
                int error = ex.Number;
                if (error == 1042) { ErrorMsg = "目前無法正常連線MySQL Database\n"; }
                if (ErrorMsg == string.Empty) { ErrorMsg = ex.Message; }
            }
            return ErrorMsg;
        }
    }
}
