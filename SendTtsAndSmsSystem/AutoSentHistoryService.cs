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
    class AutoSentHistoryService
    {
        //取得所有地點最新自動發送歷史紀錄
        public static DataTable SelectLastAutoSentHistory()
        {
            DataTable dtResult = new DataTable();
            try
            {
                SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SendTtsAndSmsSystemConnString"].ConnectionString);
                SqlCommand cmd = new SqlCommand("SELECT	[SendTtsAndSmsSystem].[dbo].[AutoSentHistory].Id, " +
                                                        "maxTimeView.AutoSentId, " +
                                                        "maxTimeView.AutoSentName, " +
                                                        "[SendTtsAndSmsSystem].[dbo].[AutoSentHistory].AutoGroupId, " +
                                                        "[SendTtsAndSmsSystem].[dbo].[AutoSentHistory].AutoGroupName, " +
                                                        "maxTimeView.maxTime " +
                                                "from [SendTtsAndSmsSystem].[dbo].[AutoSentHistory], " +
                                                    "( " +
                                                        "SELECT AutoSentId, AutoSentName, max(AutoSentTime) as maxTime " +
                                                        "from [SendTtsAndSmsSystem].[dbo].[AutoSentHistory] " +
                                                        "group by AutoSentId, AutoSentName " +
                                                    ") as maxTimeView " +
                                                "where([SendTtsAndSmsSystem].[dbo].[AutoSentHistory].AutoSentId = maxTimeView.AutoSentId) and " +
                                                "([SendTtsAndSmsSystem].[dbo].[AutoSentHistory].AutoSentName = maxTimeView.AutoSentName) and " +
                                                "([SendTtsAndSmsSystem].[dbo].[AutoSentHistory].AutoSentTime = maxTimeView.maxTime) ", connection);
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
    }
}
