using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace SendTtsAndSmsSystem
{
    class LableService
    {
        public class cbbDataList
        {
            public string cbb_No { get; set; }
            public string cbb_Data { get; set; }
        }

        //public static DataTable SelectLableAll()
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from [Lable] ", connection);
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

        public static DataTable SelectLableAll()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [Lable] ", connection);
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

        //public static DataTable GetLable(string Name, string Type)
        //{
        //    bool isMultiple = false;
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        StringBuilder str = new StringBuilder();

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);

        //        str.Append("Select * from Lable Where ");
        //        if (Name.Trim() != string.Empty)
        //        {
        //            str.Append("Name like :Name ");
        //            isMultiple = true;
        //        }

        //        if (Type.Trim() != "0")
        //        {
        //            if (isMultiple == true)
        //            {
        //                str.Append("And ");
        //            }

        //            str.Append("Type = :Type ");
        //            isMultiple = true;
        //        }

        //        OleDbCommand cmd = new OleDbCommand(str.ToString(), connection);
        //        if (Name.Trim() != string.Empty) cmd.Parameters.AddWithValue(":Name", '%'+ Name + '%');
        //        if (Type.Trim() != "0") cmd.Parameters.AddWithValue(":Type", Type);

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

        public static DataTable GetLable(string Name, string Type)
        {
            bool isMultiple = false;
            DataTable dtResult = new DataTable();
            try
            {
                StringBuilder str = new StringBuilder();

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);

                str.Append("Select * from Lable Where ");
                if (Name.Trim() != string.Empty)
                {
                    str.Append("Name like @Name ");
                    isMultiple = true;
                }

                if (Type.Trim() != "0")
                {
                    if (isMultiple == true)
                    {
                        str.Append("And ");
                    }

                    str.Append("Type = @Type ");
                    isMultiple = true;
                }

                SqlCommand cmd = new SqlCommand(str.ToString(), connection);
                if (Name.Trim() != string.Empty) cmd.Parameters.AddWithValue("@Name", '%' + Name + '%');
                if (Type.Trim() != "0") cmd.Parameters.AddWithValue("@Type", Type);

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

        //public static DataTable GetLable(string LableId)
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from [Lable] Where Id = :Id ", connection);
        //        cmd.Parameters.AddWithValue(":Id", LableId);
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

        public static DataTable GetLable(string LableId)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [Lable] Where Id = @Id ", connection);
                cmd.Parameters.AddWithValue("@Id", LableId);
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

        public static string InsertLable(string Name, string Type)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Lable] ");
                    sql.Append("( Id, Name, Type ) ");
                    sql.Append("VALUES( @Id, @Name, @Type ) ");

                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@Type", Type);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (OleDbException ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }

        //public static string UpdateLable(string Id, string Name, string Type, string ContentsSet)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();

        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("UPDATE [Lable] ");
        //            sql.Append("SET [Name] = :Name, [Type] = :Type, [ContentsSet] = :ContentsSet ");
        //            sql.Append("WHERE [Id] = :Id ");

        //            cmd.Parameters.AddWithValue(":Name", Name);
        //            cmd.Parameters.AddWithValue(":Type", Type);
        //            cmd.Parameters.AddWithValue(":ContentsSet", ContentsSet);
        //            cmd.Parameters.AddWithValue(":Id", Id);

        //            cmd.CommandText = sql.ToString();
        //            cmd.ExecuteNonQuery();
        //        }
        //    }
        //    catch (OleDbException ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    return ErrorMsg;
        //}

        public static string UpdateLable(string Id, string Name, string Type, string ContentsSet)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();

            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("UPDATE [Lable] ");
                    sql.Append("SET [Name] = @Name, [Type] = @Type, [ContentsSet] = @ContentsSet ");
                    sql.Append("WHERE [Id] = @Id ");

                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@Type", Type);
                    cmd.Parameters.AddWithValue("@ContentsSet", ContentsSet);
                    cmd.Parameters.AddWithValue("@Id", Id);

                    cmd.CommandText = sql.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (SqlException ex)
            {
                ErrorMsg = ex.Message;
            }
            return ErrorMsg;
        }

        //public static string DeleteLable(string Id)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Lable] ");
        //            sql.Append("Where Id =  :Id ");
        //            cmd.Parameters.Add(":Id", Id);

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

        public static string DeleteLable(string Id)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Lable] ");
                    sql.Append("Where Id =  @Id ");
                    cmd.Parameters.AddWithValue("@Id", Id);

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
