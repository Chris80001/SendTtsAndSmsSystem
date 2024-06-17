using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendTtsAndSmsSystem
{
    class AutoSentService
    {
        public static DataTable SelectAutoSentAll()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [AutoSend] Order by Sort ", connection);
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

        public static DataTable SelectAutoSentHistory(string AutoSentId)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select Top 20 * from [AutoSentHistory] Where AutoSentId = @AutoSentId Order by AutoSentTime desc ", connection);

                cmd.Parameters.AddWithValue("@AutoSentId", AutoSentId);
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

        public static string InserAutoSentHistory(string AutoSentId, string AutoSentName, DateTime AutoSentTime)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("Insert INTO [SendTtsAndSmsSystem].[dbo].[AutoSentHistory](Id, AutoSentId, AutoSentName, AutoSentTime) ");
                    sql.Append("VALUES (@Id, @AutoSentId, @AutoSentName, @AutoSentTime) ");

                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid().ToString());
                    cmd.Parameters.AddWithValue("@AutoSentId", AutoSentId);
                    cmd.Parameters.AddWithValue("@AutoSentName", AutoSentName);
                    cmd.Parameters.AddWithValue("@AutoSentTime", AutoSentTime);

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

        public static string DeleteAutoGroup()
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [AutoGroup] ");

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

        public static string MultipleInsertAutoGroup(List<string> AutoGroupNameList)
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
                    sql.Append("INSERT INTO [AutoGroup] ");
                    sql.Append("( Id, Name, CreateTime ) ");
                    sql.Append("VALUES ");
                    foreach (string GroupName in AutoGroupNameList)
                    {
                        if (i != 0) { sql.Append(","); }
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

        public static DataTable SelectAutoGroupAll()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from [AutoGroup] Order by Name ", connection);
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

        public static string DeleteAutoEmployee()
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [AutoEmployee] ");

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

        public static string MultipleInsertAutoEmployee(DataTable CsvDb, DataTable dtAutoGroup)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [AutoEmployee] ");
                    sql.Append("( EmployeeId, GroupId, Name, PhoneNumber, GroupName ) ");
                    sql.Append("VALUES ");
                    int i = 0;
                    foreach (DataRow row in CsvDb.Rows)
                    {
                        //var GroupId = dtGroup.AsEnumerable().Where(w => w.Field<string>("Name") == row.Field<string>(1)).Select(d => d.Field<string>("Id"));
                        var GroupId = dtAutoGroup.AsEnumerable().Where(w => w.Field<string>("Name") == row.Field<string>(0)).Select(d => d.Field<string>("Id"));
                        string PhoneNumber = string.Empty;
                        if (row.Field<string>(3).Substring(0, 1) != "0")
                        { PhoneNumber = "0" + row.Field<string>(3); }
                        else { PhoneNumber = row.Field<string>(3); }

                        if (i != 0) { sql.Append(","); }
                        sql.Append("( @EmployeeId" + i + ", @GroupId" + i + ", @Name" + i + ", @PhoneNumber" + i + ", @GroupName" + i + " ) ");

                        //cmd.Parameters.AddWithValue("@EmployeeId" + i, row.Field<string>(0));
                        cmd.Parameters.AddWithValue("@EmployeeId" + i, row.Field<string>(1));
                        cmd.Parameters.AddWithValue("@GroupId" + i, GroupId.ToList()[0]);
                        cmd.Parameters.AddWithValue("@Name" + i, row.Field<string>(2));
                        cmd.Parameters.AddWithValue("@PhoneNumber" + i, PhoneNumber);
                        //cmd.Parameters.AddWithValue("@GroupName" + i, row.Field<string>(1));
                        cmd.Parameters.AddWithValue("@GroupName" + i, row.Field<string>(0));
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

        public static DataTable GetAutoEmployee()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from AutoEmployee order by GroupName, EmployeeId ", connection);

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

        public static DataTable GetAutoEmployeeForAutoGroupId(string AutoGroupId)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from AutoEmployee " +
                                                "Where [GroupId] = @GroupId " +
                                                "order by EmployeeId ", connection);

                cmd.Parameters.AddWithValue("@GroupId", AutoGroupId);

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

        public static string UpdateAutoSend(string Id, int IsSent, DateTime LastSentTime)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    //寫入IsSent值為0，表示一般復歸而已，沒有發送訊息
                    if (IsSent == 0)
                    {
                        sql.Append("update [SendTtsAndSmsSystem].[dbo].[AutoSend] ");
                        sql.Append("set IsSent = @IsSent ");
                        sql.Append("where Id = @Id ");

                        cmd.Parameters.AddWithValue("@Id", Id);
                        cmd.Parameters.AddWithValue("@IsSent", IsSent);
                    }
                    //寫入IsSent值為1，表示此次有發送訊息，須更新LastSentTime
                    if (IsSent == 1)
                    {
                        sql.Append("update [SendTtsAndSmsSystem].[dbo].[AutoSend] ");
                        sql.Append("set IsSent = @IsSent, LastSentTime = @LastSentTime ");
                        sql.Append("where Id = @Id ");

                        cmd.Parameters.AddWithValue("@Id", Id);
                        cmd.Parameters.AddWithValue("@IsSent", IsSent);
                        cmd.Parameters.AddWithValue("@LastSentTime", LastSentTime);
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

        public static string MultipleUpdateAutoSend(DataTable dtAutoGroup, DataTable dtAutoSend)
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
                    foreach (DataRow row in dtAutoSend.Rows)
                    {
                        string Name = row.Field<string>("Name");
                        DataTable dtTemp = dtAutoGroup.AsEnumerable().Where(w => w.Field<string>("Name") == Name).CopyToDataTable();
                        string AutoGroupId = string.Empty;
                        if (dtTemp.Rows.Count > 0)
                        {
                            AutoGroupId = dtTemp.Rows[0].Field<string>("Id");
                        }
                        if (AutoGroupId == string.Empty)
                        {
                            continue;
                        }

                        sql.Append("update [SendTtsAndSmsSystem].[dbo].[AutoSend] ");
                        sql.Append("set AutoGroupId = @AutoGroupId" + i + " ");
                        sql.Append("where Name = @Name" + i + " ");

                        cmd.Parameters.AddWithValue("@AutoGroupId" + i, AutoGroupId);
                        cmd.Parameters.AddWithValue("@Name" + i, Name);
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
    }
}
