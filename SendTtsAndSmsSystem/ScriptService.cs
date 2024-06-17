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
    public class ScriptService
    {
        //public static DataTable SelectScriptAll()
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from Script ", connection);
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

        //    return  dtResult;
        //}

        public static DataTable SelectScriptAll()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from Script ", connection);
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

        //public static DataTable GetScript(string Id)
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from Script Where Id = :Id ", connection);
        //        cmd.Parameters.AddWithValue(":Id", Id);
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

        public static DataTable GetScript(string Id)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from Script Where Id = @Id ", connection);
                cmd.Parameters.AddWithValue("@Id", Id);
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

        //public static DataTable GetScriptByName(string Name)
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from Script Where Name like :Name ", connection);
        //        cmd.Parameters.AddWithValue(":Name", '%' + Name + '%');
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

        public static DataTable GetScriptByName(string Name)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from Script Where Name like @Name ", connection);
                cmd.Parameters.AddWithValue("@Name", '%' + Name + '%');
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

        //public static string InsertScript(string Name, string Depiction)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("INSERT INTO [Script] ");
        //            sql.Append("( Id, Name, Depiction ) ");
        //            sql.Append("VALUES( :Id, :Name, :Depiction ) ");

        //            cmd.Parameters.Add(":Id", Guid.NewGuid().ToString());
        //            cmd.Parameters.Add(":Name", Name);
        //            cmd.Parameters.Add(":Depiction", Depiction);

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

        public static string InsertScript(string Name, string Depiction)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Script] ");
                    sql.Append("( Id, Name, Depiction ) ");
                    sql.Append("VALUES( @Id, @Name, @Depiction ) ");

                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@Depiction", Depiction);

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

        //public static string InsertScript(string Name, int TextBoxCount, int LableCount, string TextBoxContentsSet, string LableContentSet, string Depiction)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("INSERT INTO [Script] ");
        //            sql.Append("( Id, Name, TextBoxCount, LableCount, TextBoxContentsSet, LableContentSet, Depiction ) ");
        //            sql.Append("VALUES( :Id, :Name, :TextBoxCount, :LableCount, :TextBoxContentsSet, :LableContentSet, :Depiction ) ");

        //            cmd.Parameters.Add(":Id", Guid.NewGuid().ToString());
        //            cmd.Parameters.Add(":Name", Name);
        //            cmd.Parameters.Add(":TextBoxCount", TextBoxCount);
        //            cmd.Parameters.Add(":LableCount", LableCount);
        //            cmd.Parameters.Add(":TextBoxContentsSet", TextBoxContentsSet);
        //            cmd.Parameters.Add(":LableContentSet", LableContentSet);
        //            cmd.Parameters.Add(":Depiction", Depiction);

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

        public static string InsertScript(string Name, int TextBoxCount, int LableCount, string TextBoxContentsSet, string LableContentSet, string Depiction)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Script] ");
                    sql.Append("( Id, Name, TextBoxCount, LableCount, TextBoxContentsSet, LableContentSet, Depiction ) ");
                    sql.Append("VALUES( @Id, @Name, @TextBoxCount, @LableCount, @TextBoxContentsSet, @LableContentSet, @Depiction ) ");

                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@TextBoxCount", TextBoxCount);
                    cmd.Parameters.AddWithValue("@LableCount", LableCount);
                    cmd.Parameters.AddWithValue("@TextBoxContentsSet", TextBoxContentsSet);
                    cmd.Parameters.AddWithValue("@LableContentSet", LableContentSet);
                    cmd.Parameters.AddWithValue("@Depiction", Depiction);

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

        //public static string UpdateScript(string Id, string Name, int TextBoxCount, int LableCount, string TextBoxContentsSet, string LableContentSet, string Depiction)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();

        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("UPDATE [Script] ");
        //            sql.Append("SET [Name] = :Name, [TextBoxCount] = :TextBoxCount, [LableCount] = :LableCount, [TextBoxContentsSet] = :TextBoxContentsSet, [LableContentSet] = :LableContentSet, [Depiction] = :Depiction ");
        //            sql.Append("WHERE [Id] = :Id ");

        //            cmd.Parameters.Add(":Name", Name);
        //            cmd.Parameters.Add(":TextBoxCount", TextBoxCount);
        //            cmd.Parameters.Add(":LableCount", LableCount);
        //            cmd.Parameters.Add(":TextBoxContentsSet", TextBoxContentsSet);
        //            cmd.Parameters.Add(":LableContentSet", LableContentSet);
        //            cmd.Parameters.Add(":Depiction", Depiction);
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

        public static string UpdateScript(string Id, string Name, int TextBoxCount, int LableCount, string TextBoxContentsSet, string LableContentSet, string Depiction)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();

            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("UPDATE [Script] ");
                    sql.Append("SET [Name] = @Name, [TextBoxCount] = @TextBoxCount, [LableCount] = @LableCount, [TextBoxContentsSet] = @TextBoxContentsSet, [LableContentSet] = @LableContentSet, [Depiction] = @Depiction ");
                    sql.Append("WHERE [Id] = @Id ");

                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@TextBoxCount", TextBoxCount);
                    cmd.Parameters.AddWithValue("@LableCount", LableCount);
                    cmd.Parameters.AddWithValue("@TextBoxContentsSet", TextBoxContentsSet);
                    cmd.Parameters.AddWithValue("@LableContentSet", LableContentSet);
                    cmd.Parameters.AddWithValue("@Depiction", Depiction);
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

        //public static string DeleteScript(string Id)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Script] ");
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

        public static string DeleteScript(string Id)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Script] ");
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
