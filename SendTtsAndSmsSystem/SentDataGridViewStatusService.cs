using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace SendTtsAndSmsSystem
{
    class SentDataGridViewStatusService
    {
        //public static DataTable SelectSentDataGridViewStatus()
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from [SentDataGridViewGroupStatus] ", connection);
        //        connection.Open();

        //        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //        da.Fill(dtResult);
        //        connection.Close();
        //        da.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        string ErrorMsg = ex.Message;
        //    }

        //    return dtResult;
        //}

        public static DataTable SelectSentDataGridViewStatus()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [SentDataGridViewGroupStatus] ", connection);
                connection.Open();

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dtResult);
                connection.Close();
                da.Dispose();
            }
            catch (Exception ex)
            {
                string ErrorMsg = ex.Message;
            }

            return dtResult;
        }

        //public static string InsertSentDataGridViewStatus(string GroupId, string GroupName)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("INSERT INTO  [SentDataGridViewGroupStatus] ");
        //            sql.Append("( [GroupId], [GroupName] ) ");
        //            sql.Append("VALUES( :GroupId, :GroupName ) ");

        //            cmd.Parameters.Add(":GroupId", GroupId);
        //            cmd.Parameters.Add(":GroupName", GroupName);

        //            cmd.CommandText = sql.ToString();
        //            cmd.ExecuteNonQuery();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    return ErrorMsg;
        //}

        public static string InsertSentDataGridViewStatus(string GroupId, string GroupName)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO  [SentDataGridViewGroupStatus] ");
                    sql.Append("( [GroupId], [GroupName] ) ");
                    sql.Append("VALUES( @GroupId, @GroupName ) ");

                    cmd.Parameters.AddWithValue("@GroupId", GroupId);
                    cmd.Parameters.AddWithValue("@GroupName", GroupName);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }

        //public static string InsertSentDataGridViewStatus(List<string> GroupIdList, List<string> GroupNameList)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        foreach (string GroupId in GroupIdList)
        //        {
        //            using (OleDbCommand cmd = connection.CreateCommand())
        //            {
        //                sql.Append("INSERT INTO  [SentDataGridViewGroupStatus] ");
        //                sql.Append("( [GroupId], [GroupName] ) ");
        //                sql.Append("VALUES( :GroupId, :GroupName ) ");

        //                cmd.Parameters.Add(":GroupId", GroupId);
        //                cmd.Parameters.Add(":GroupName", GroupNameList[GroupIdList.IndexOf(GroupId)]);

        //                cmd.CommandText = sql.ToString();
        //                cmd.ExecuteNonQuery();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    return ErrorMsg;
        //}

        public static string InsertSentDataGridViewStatus(List<string> GroupIdList, List<string> GroupNameList)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();

                using (SqlCommand cmd = connection.CreateCommand())
                {
                    int i = 0;
                    sql.Append("INSERT INTO  [SentDataGridViewGroupStatus] ");
                    sql.Append("( [GroupId], [GroupName] ) ");
                    sql.Append("VALUES ");
                    foreach (string GroupId in GroupIdList)
                    {
                        if (i != 0) { sql.Append(","); }
                        sql.Append("( @GroupId"+ i + ", @GroupName" + i + " ) ");

                        cmd.Parameters.AddWithValue("@GroupId" + i, GroupId);
                        cmd.Parameters.AddWithValue("@GroupName" + i, GroupNameList[GroupIdList.IndexOf(GroupId)]);
                        i++;
                    }
                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }

        //public static string DAOInsertSentDataGridViewStatus(List<string> GroupIdList, List<string> GroupNameList)
        //{
        //    string ErrorMsg = string.Empty;

        //    DAO.DBEngine dbEngine = new DAO.DBEngine();
        //    DAO.Database db = dbEngine.OpenDatabase(ConfigurationManager.AppSettings["MdbPath"]);
        //    //先將TABLE清空
        //    db.Execute("DELETE FROM [SentDataGridViewGroupStatus]");

        //    DAO.Recordset rs = db.OpenRecordset("SentDataGridViewGroupStatus");
        //    DAO.Field[] myFields = new DAO.Field[2];
        //    myFields[0] = rs.Fields["GroupId"];
        //    myFields[1] = rs.Fields["GroupName"];
        //    foreach (string GroupId in GroupIdList)
        //    {
        //        rs.AddNew();
        //        myFields[0].Value = GroupId;
        //        myFields[1].Value = GroupNameList[GroupIdList.IndexOf(GroupId)];
        //        try
        //        {
        //            rs.Update();
        //        }
        //        catch (Exception ex)
        //        {
        //            ErrorMsg = ErrorMsg + ex.Message + "\n";
        //        }

        //    }

        //    rs.Close();
        //    db.Close();


        //    return ErrorMsg;
        //}

        public static string MultipleInsertSentDataGridViewStatus(List<string> GroupIdList, List<string> GroupNameList)
        {
            string ErrorMsg = string.Empty;

            DAO.DBEngine dbEngine = new DAO.DBEngine();
            DAO.Database db = dbEngine.OpenDatabase(ConfigurationManager.AppSettings["MdbPath"]);
            //先將TABLE清空
            db.Execute("DELETE FROM [SentDataGridViewGroupStatus]");

            DAO.Recordset rs = db.OpenRecordset("SentDataGridViewGroupStatus");
            DAO.Field[] myFields = new DAO.Field[2];
            myFields[0] = rs.Fields["GroupId"];
            myFields[1] = rs.Fields["GroupName"];
            foreach (string GroupId in GroupIdList)
            {
                rs.AddNew();
                myFields[0].Value = GroupId;
                myFields[1].Value = GroupNameList[GroupIdList.IndexOf(GroupId)];
                try
                {
                    rs.Update();
                }
                catch (Exception ex)
                {
                    ErrorMsg = ErrorMsg + ex.Message + "\n";
                }

            }

            rs.Close();
            db.Close();


            return ErrorMsg;
        }

        //public static string DeleteSentDataGridViewStatus(string GroupId)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [SentDataGridViewGroupStatus] ");
        //            sql.Append("Where [GroupId] = :GroupId ");

        //            cmd.Parameters.Add(":GroupId", GroupId);

        //            cmd.CommandText = sql.ToString();
        //            cmd.ExecuteNonQuery();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    return ErrorMsg;
        //}

        public static string DeleteSentDataGridViewStatus(string GroupId)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [SentDataGridViewGroupStatus] ");
                    sql.Append("Where [GroupId] = @GroupId ");

                    cmd.Parameters.AddWithValue("@GroupId", GroupId);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }

        //public static string DeleteSentDataGridViewStatus()
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [SentDataGridViewGroupStatus] ");

        //            cmd.CommandText = sql.ToString();
        //            cmd.ExecuteNonQuery();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    return ErrorMsg;
        //}

        public static string DeleteSentDataGridViewStatus()
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [SentDataGridViewGroupStatus] ");

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }
    }
}
