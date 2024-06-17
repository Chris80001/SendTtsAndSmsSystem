using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace SendTtsAndSmsSystem
{
    class GroupService
    {
        //public static DataTable SelectGroupAll()
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from [Group] Order by Name ", connection);
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

        public static DataTable SelectGroupAll()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [Group] Order by Name ", connection);
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

        //public static DataTable SelectGroup(string Id)
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from [Group] Where Id = :Id Order by CreateTime, Name ", connection);

        //        cmd.Parameters.Add(":Id", Id);
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

        public static DataTable SelectGroup(string Id)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [Group] Where Id = @Id Order by CreateTime, Name ", connection);

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

        //public static DataTable DAOInsertGroup(List<string> GroupNameList)
        //{
        //    string ErrorMsg = string.Empty;
        //    //DataTable Set
        //    DataTable dtGroup = new DataTable("NewDable");
        //    DataRow row;
        //    dtGroup.Columns.Add("GroupId");
        //    dtGroup.Columns.Add("GroupName");
        //    string[] strLine;

        //    //DAO
        //    DAO.DBEngine dbEngine = new DAO.DBEngine();
        //    DAO.Database db = dbEngine.OpenDatabase(ConfigurationManager.AppSettings["MdbPath"]);
        //    //先將TABLE清空
        //    db.Execute("DELETE FROM [Group]");
        //    DAO.Recordset rs = db.OpenRecordset("Group");

        //    DAO.Field[] myFields = new DAO.Field[3];
        //    myFields[0] = rs.Fields["Id"];
        //    myFields[1] = rs.Fields["Name"];
        //    myFields[2] = rs.Fields["CreateTime"];
        //    foreach (string GroupName in GroupNameList)
        //    {
        //        rs.AddNew();
        //        string NewGuid = Guid.NewGuid().ToString();

        //        myFields[0].Value = NewGuid;
        //        myFields[1].Value = GroupName;
        //        myFields[2].Value = DateTime.Now.ToString();

        //        //帶入返回Dt
        //        strLine = new string[] { NewGuid, GroupName };
        //        row = dtGroup.NewRow();
        //        row.ItemArray = strLine;
        //        dtGroup.Rows.Add(row);

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

        //    return dtGroup;
        //}

        public static string MultipleInsertGroup(List<string> GroupNameList)
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
                    sql.Append("INSERT INTO [Group] ");
                    sql.Append("( Id, Name, CreateTime ) ");
                    sql.Append("VALUES ");
                    foreach (string GroupName in GroupNameList)
                    {
                        if (i != 0){ sql.Append(","); }
                        sql.Append("( @Id" + i + ", @Name" + i + ", @CreateTime" + i + " )");

                        cmd.Parameters.AddWithValue("@Id" + i, Guid.NewGuid().ToString());
                        cmd.Parameters.AddWithValue("@Name" + i, GroupName);
                        cmd.Parameters.AddWithValue("@CreateTime" + i, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                        i++;
                    }
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

        //public static string InsertGroup(string Name)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("INSERT INTO [Group] ");
        //            sql.Append("( Id, Name, CreateTime ) ");
        //            sql.Append("VALUES( :Id, :Name, :CreateTime ) ");

        //            cmd.Parameters.Add("Id", Guid.NewGuid().ToString());
        //            cmd.Parameters.Add("Name", Name);
        //            cmd.Parameters.Add("CreateTime", DateTime.Now.ToString());

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

        public static string InsertGroup(string Name)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Group] ");
                    sql.Append("( Id, Name, CreateTime ) ");
                    sql.Append("VALUES( @Id, @Name, @CreateTime ) ");

                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@CreateTime", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

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

        //public static string UpdateGroup(string Id, string Name)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();

        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("UPDATE [Group] ");
        //            sql.Append("SET [Name] = :Name ");
        //            sql.Append("WHERE [Id] = :Id ");

        //            cmd.Parameters.AddWithValue(":Name", Name);
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

        public static string UpdateGroup(string Id, string Name)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();

            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("UPDATE [Group] ");
                    sql.Append("SET [Name] = @Name ");
                    sql.Append("WHERE [Id] = @Id ");

                    cmd.Parameters.AddWithValue("@Name", Name);
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


        //public static string DeleteGroup()
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Group] ");

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

        public static string DeleteGroup()
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Group] ");

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

        //public static string DeleteGroup(string Id, string ErrorMsg)
        //{
        //    ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Group] ");
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
        public static string DeleteGroup(string Id, string ErrorMsg)
        {
            ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Group] ");
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
