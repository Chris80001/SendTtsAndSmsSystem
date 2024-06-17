using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace SendTtsAndSmsSystem
{
    class EmployeeService
    {
        //public static DataTable GetEmployee()
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from Employee order by GroupName, EmployeeId ", connection);

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

        public static DataTable GetEmployee()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from Employee order by GroupName, EmployeeId ", connection);

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

        //public static DataTable GetEmployee(string GroupId)
        //{
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        OleDbCommand cmd = new OleDbCommand("Select * from Employee Where GroupId = :GroupId order by EmployeeId ", connection);
        //        cmd.Parameters.AddWithValue(":GroupId", GroupId);
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

        public static DataTable GetEmployee(string GroupId)
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("Select * from Employee Where GroupId = @GroupId order by EmployeeId ", connection);
                cmd.Parameters.AddWithValue("@GroupId", GroupId);
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

        //public static DataTable GetEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber)
        //{
        //    bool isMultiple = false;
        //    DataTable dtResult = new DataTable();
        //    try
        //    {
        //        StringBuilder str = new StringBuilder();

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);

        //        str.Append("Select * from Employee Where ");
        //        if (EmployeeId.Trim() != string.Empty)
        //        { 
        //            str.Append("EmployeeId like :EmployeeId ");
        //            isMultiple = true;
        //        }
        //        if (GroupId.Trim() != "0")
        //        {
        //            if (isMultiple == true)
        //            {
        //                str.Append("And ");
        //            }
        //            str.Append("GroupId = :GroupId ");
        //            isMultiple = true;
        //        }
        //        if (Name.Trim() != string.Empty)
        //        {
        //            if (isMultiple == true)
        //            {
        //                str.Append("And ");
        //            }

        //            str.Append("Name like :Name ");
        //            isMultiple = true;
        //        }
        //        if (PhoneNumber.Trim() != string.Empty)
        //        {
        //            if (isMultiple == true)
        //            {
        //                str.Append("And ");
        //            }

        //            str.Append("PhoneNumber like :PhoneNumber ");
        //            isMultiple = true;
        //        }


        //        str.Append("order by EmployeeId ");

        //        OleDbCommand cmd = new OleDbCommand(str.ToString(), connection);
        //        if (EmployeeId.Trim() != string.Empty) cmd.Parameters.AddWithValue(":EmployeeId", '%' + EmployeeId + '%');
        //        if (GroupId.Trim() != "0") cmd.Parameters.AddWithValue(":GroupId", GroupId);
        //        if (Name.Trim() != string.Empty) cmd.Parameters.AddWithValue(":Name", '%' + Name + '%');
        //        if (PhoneNumber.Trim() != string.Empty) cmd.Parameters.AddWithValue(":PhoneNumber", '%' + PhoneNumber + '%');

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

        public static DataTable GetEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber)
        {
            bool isMultiple = false;
            DataTable dtResult = new DataTable();
            try
            {
                StringBuilder str = new StringBuilder();

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);

                str.Append("Select * from Employee Where ");
                if (EmployeeId.Trim() != string.Empty)
                {
                    str.Append("EmployeeId like @EmployeeId ");
                    isMultiple = true;
                }
                if (GroupId.Trim() != "0")
                {
                    if (isMultiple == true)
                    {
                        str.Append("And ");
                    }
                    str.Append("GroupId = @GroupId ");
                    isMultiple = true;
                }
                if (Name.Trim() != string.Empty)
                {
                    if (isMultiple == true)
                    {
                        str.Append("And ");
                    }

                    str.Append("Name like @Name ");
                    isMultiple = true;
                }
                if (PhoneNumber.Trim() != string.Empty)
                {
                    if (isMultiple == true)
                    {
                        str.Append("And ");
                    }

                    str.Append("PhoneNumber like @PhoneNumber ");
                    isMultiple = true;
                }


                str.Append("order by EmployeeId ");

                SqlCommand cmd = new SqlCommand(str.ToString(), connection);
                if (EmployeeId.Trim() != string.Empty) cmd.Parameters.AddWithValue("@EmployeeId", '%' + EmployeeId + '%');
                if (GroupId.Trim() != "0") cmd.Parameters.AddWithValue("@GroupId", GroupId);
                if (Name.Trim() != string.Empty) cmd.Parameters.AddWithValue("@Name", '%' + Name + '%');
                if (PhoneNumber.Trim() != string.Empty) cmd.Parameters.AddWithValue("@PhoneNumber", '%' + PhoneNumber + '%');

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

        //public static string DAOInsertEmployee(DataTable CsvDb, DataTable dtGroup)
        //{
        //    string ErrorMsg = string.Empty;

        //    DAO.DBEngine dbEngine = new DAO.DBEngine();
        //    DAO.Database db = dbEngine.OpenDatabase(ConfigurationManager.AppSettings["MdbPath"]);
        //    //先將TABLE清空
        //    db.Execute("DELETE FROM [Employee]");

        //    DAO.Recordset rs = db.OpenRecordset("Employee");
        //    DAO.Field[] myFields = new DAO.Field[5];
        //    myFields[0] = rs.Fields["EmployeeId"];
        //    myFields[1] = rs.Fields["GroupId"];
        //    myFields[2] = rs.Fields["Name"];
        //    myFields[3] = rs.Fields["PhoneNumber"];
        //    myFields[4] = rs.Fields["GroupName"];

        //    foreach (DataRow row in CsvDb.Rows)
        //    {
        //        rs.AddNew();
        //        var GroupId = dtGroup.AsEnumerable().Where(w => w.Field<string>("GroupName") == row.Field<string>(1)).Select(d => d.Field<string>("GroupId"));
        //        string PhoneNumber = string.Empty;
        //        if (row.Field<string>(3).Substring(0, 1) != "0")
        //        { PhoneNumber = "0" + row.Field<string>(3); }
        //        else { PhoneNumber = row.Field<string>(3); }

        //        myFields[0].Value = row.Field<string>(0);
        //        myFields[1].Value = GroupId.ToList()[0];
        //        myFields[2].Value = row.Field<string>(2);
        //        myFields[3].Value = PhoneNumber;
        //        myFields[4].Value = row.Field<string>(1);

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

        public static string MultipleInsertEmployee(DataTable CsvDb, DataTable dtGroup)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Employee] ");
                    sql.Append("( EmployeeId, GroupId, Name, PhoneNumber, GroupName ) ");
                    sql.Append("VALUES ");
                    int i = 0;
                    foreach (DataRow row in CsvDb.Rows)
                    {
                        //var GroupId = dtGroup.AsEnumerable().Where(w => w.Field<string>("Name") == row.Field<string>(1)).Select(d => d.Field<string>("Id"));
                        var GroupId = dtGroup.AsEnumerable().Where(w => w.Field<string>("Name") == row.Field<string>(0)).Select(d => d.Field<string>("Id"));
                        string PhoneNumber = string.Empty;
                        if (row.Field<string>(3).Substring(0, 1) != "0")
                        { PhoneNumber = "0" + row.Field<string>(3); }
                        else { PhoneNumber = row.Field<string>(3); }

                        if (i != 0) { sql.Append(","); }
                        sql.Append("( @EmployeeId"+ i + ", @GroupId" + i + ", @Name" + i + ", @PhoneNumber" + i + ", @GroupName" + i + " ) ");

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

        //public static string InsertEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber, string GroupName)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("INSERT INTO [Employee] ");
        //            sql.Append("( EmployeeId, GroupId, Name, PhoneNumber, GroupName ) ");
        //            sql.Append("VALUES( :EmployeeId, :GroupId, :Name, :PhoneNumber, :GroupName ) ");

        //            cmd.Parameters.Add("EmployeeId", EmployeeId);
        //            cmd.Parameters.Add("GroupId", GroupId);
        //            cmd.Parameters.Add("Name", Name);
        //            cmd.Parameters.Add("PhoneNumber", PhoneNumber);
        //            cmd.Parameters.Add("GroupName", GroupName);

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

        public static string InsertEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber, string GroupName)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("INSERT INTO [Employee] ");
                    sql.Append("( EmployeeId, GroupId, Name, PhoneNumber, GroupName ) ");
                    sql.Append("VALUES( @EmployeeId, @GroupId, @Name, @PhoneNumber, @GroupName ) ");

                    cmd.Parameters.AddWithValue("@EmployeeId", EmployeeId);
                    cmd.Parameters.AddWithValue("@GroupId", GroupId);
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@PhoneNumber", PhoneNumber);
                    cmd.Parameters.AddWithValue("@GroupName", GroupName);

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

        //public static string UpdateEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber, string GroupName)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();

        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("UPDATE [Employee] ");
        //            sql.Append("SET [GroupId] = :GroupId, Name = :Name, PhoneNumber = :PhoneNumber, GroupName = :GroupName ");
        //            sql.Append("WHERE [EmployeeId] = :EmployeeId ");

        //            cmd.Parameters.AddWithValue(":GroupId", GroupId);
        //            cmd.Parameters.AddWithValue(":Name", Name);
        //            cmd.Parameters.AddWithValue(":PhoneNumber", PhoneNumber);
        //            cmd.Parameters.AddWithValue(":GroupName", GroupName);
        //            cmd.Parameters.AddWithValue(":EmployeeId", EmployeeId);

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

        public static string UpdateEmployee(string EmployeeId, string GroupId, string Name, string PhoneNumber, string GroupName)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();

            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("UPDATE [Employee] ");
                    sql.Append("SET [GroupId] = @GroupId, Name = @Name, PhoneNumber = @PhoneNumber, GroupName = @GroupName ");
                    sql.Append("WHERE [EmployeeId] = @EmployeeId ");

                    cmd.Parameters.AddWithValue("@GroupId", GroupId);
                    cmd.Parameters.AddWithValue("@Name", Name);
                    cmd.Parameters.AddWithValue("@PhoneNumber", PhoneNumber);
                    cmd.Parameters.AddWithValue("@GroupName", GroupName);
                    cmd.Parameters.AddWithValue("@EmployeeId", EmployeeId);

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

        //public static string DeleteEmployee()
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Employee] ");

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

        public static string DeleteEmployee()
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Employee] ");

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

        //public static string DeleteEmployee(string EmployeeId, string GroupId)
        //{
        //    string ErrorMsg = string.Empty;
        //    StringBuilder sql = new StringBuilder();
        //    try
        //    {

        //        OleDbConnection connection = new OleDbConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
        //        connection.Open();
        //        using (OleDbCommand cmd = connection.CreateCommand())
        //        {
        //            sql.Append("DELETE From [Employee] ");
        //            sql.Append("Where EmployeeId =  :EmployeeId and GroupId = :GroupId ");
        //            cmd.Parameters.Add(":EmployeeId", EmployeeId);
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

        public static string DeleteEmployee(string EmployeeId, string GroupId)
        {
            string ErrorMsg = string.Empty;
            StringBuilder sql = new StringBuilder();
            try
            {

                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                connection.Open();
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    sql.Append("DELETE From [Employee] ");
                    sql.Append("Where EmployeeId =  @EmployeeId and GroupId = @GroupId ");
                    cmd.Parameters.AddWithValue("@EmployeeId", EmployeeId);
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
    }

}
