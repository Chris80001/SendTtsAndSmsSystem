using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using System.Windows.Forms.VisualStyles;
using System.IO;
using SendTtsAndSmsSystem.Properties;
using System.Diagnostics.Eventing.Reader;
using System.Net.Sockets;
using OMRON.Compolet.CIP;
using System.Threading.Tasks;

namespace SendTtsAndSmsSystem
{
    public partial class Form1 : Form
    {
        Thread _CheckNetwork;
        Thread _tdSent;
        Thread _tdAutoSentStatusPicture;
        bool _blAutoSentStatusPicture;
        delegate void PrintHandler(TextBox tb, string text);
        public static TextBox _textbox;
        public static TextBox _textboxAuto;
        private delegate void tbUpdate(String str, TextBox tb);
        private void updateTextBox(String str, TextBox tb)
        {
            if (this.InvokeRequired)
            {
                tbUpdate uu = new tbUpdate(updateTextBox);
                this.Invoke(uu, str, tb);
            }
            else
            {
                tb.AppendText(str);
            }
        }
        private void ClearTextBox(String str, TextBox tb)
        {
            if (this.InvokeRequired)
            {
                tbUpdate uu = new tbUpdate(ClearTextBox);
                this.Invoke(uu, str, tb);
            }
            else
            {
                tb.Clear();
            }
        }
        private delegate void cbEnble(bool b, CheckBox cb);
        private void EnbleCheckBox(bool b, CheckBox cb)
        {
            if (this.InvokeRequired)
            {
                cbEnble uu = new cbEnble(EnbleCheckBox);
                this.Invoke(uu, b, cb);
            }
            else
            {
                cb.Enabled = b;
            }
        }
        bool _isInitializeScriptLableStatus = false;
        public enum LableEnum
        {
            Number,
            General,
            Text,
            Record
        }

        public Form1()
        {
            InitializeComponent();
            InitializeControl();
            InitializeAsyncForm();
            InitializeEmpolyeeCcbGroup();
            InitializeLableCcbGroup();

            //檢查目前與Client連現狀控
            _CheckNetwork = new Thread(_CheckConnection);
            _CheckNetwork.IsBackground = true;
            _CheckNetwork.Start();

            //調整螢幕
            this.Height = this.Height * 105 / 100;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //InitializeSentDataGridViewGroupStatus();
            InitializedgvdgvAutoSendForm();

            //檢查目前與Client連現狀控
            _blAutoSentStatusPicture = true;
            _tdAutoSentStatusPicture = new Thread(AutoSentStatusPicture);
            _tdAutoSentStatusPicture.IsBackground = true;
            _tdAutoSentStatusPicture.Start();

            //程式開啟自動跳到37處無人發報頁籤
            tabControl1.SelectedTab = tabControl1.TabPages[7];
        }

        //Thread檢查目前與中華電信網路連線狀態
        private void _CheckConnection()
        {
            while (true)
            {
                string ServerIp = ConfigurationManager.AppSettings["ServerIp"];
                int ServerPort = int.Parse(ConfigurationManager.AppSettings["ServerPort"]);

                try
                {
                    using (var client = new TcpClient(ServerIp, ServerPort))
                        pbLight.Image = Resources.Aqua_Ball_Green;
                }
                catch (SocketException ex)
                {
                    pbLight.Image = Resources.Aqua_Ball_Red;
                }

                //每分鐘檢查一次
                Thread.Sleep(60000);
            }
        }

        private void InitializeControl()
        {
            //腳本
            DataTable dtScript = ScriptService.SelectScriptAll();
            DataRow dr = dtScript.NewRow();
            dr["Name"] = "請選擇腳本...";
            dr["Id"] = "0";
            dtScript.Rows.InsertAt(dr, 0);

            cbbSentScript.DataSource = dtScript;
            cbbSentScript.DisplayMember = "Name";
            cbbSentScript.ValueMember = "Id";

            //群組
            //Get data
            DataTable dtGroup = GroupService.SelectGroupAll();
            //Set datagridview
            dgvGroup.DataSource = null;
            dgvGroup.Refresh();
            dgvGroup.Columns.Clear();
            dgvGroup.DataSource = dtGroup;
            dgvGroup.Columns["Id"].Visible = false;
            dgvGroup.Columns["CreateTime"].Visible = false;
            dgvGroup.Columns["Name"].HeaderText = "群組名稱";
            DataGridViewCheckBoxColumn CheckBox = new DataGridViewCheckBoxColumn();
            dgvGroup.Columns.Insert(0, CheckBox);
            dgvGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvGroup.Columns[0].Width = 30;
            this.dgvGroup.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvGroup.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvGroup.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;

            //建立全選CheckBox
            //計算CheckBox嵌入DataGridView的位置
            //Rectangle rect = dgvGroup.GetCellDisplayRectangle(0, -1, true);
            //rect.X = rect.Location.X + rect.Width / 4 - 9 + 59;
            //rect.Y = rect.Location.Y + (rect.Height / 2 - 9) + 15;
            Rectangle rect = dgvGroup.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 9 + 66;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9) + 14;

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(13, 13);
            cbHeader.Location = rect.Location;
            //全選要設定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);
            //將 CheckBox 加入到 dataGridView
            dgvGroup.Controls.Add(cbHeader);
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            if (dgvGroup.Rows.Count > 0)
            {
                dgvGroup.CurrentRow.Cells[0].Selected = false;
                dgvGroup.CurrentRow.Cells[0].ReadOnly = true;
                foreach (DataGridViewRow dr in dgvGroup.Rows)
                {
                    dr.Cells[0].Value = ((CheckBox)dgvGroup.Controls.Find("checkboxHeader", true)[0]).Checked;
                    if (dr.Cells[0].ReadOnly == true)
                    {
                        dgvGroup.CurrentRow.Cells[0].Selected = true;
                        dgvGroup.CurrentRow.Cells[0].ReadOnly = false;
                    }
                }
            }
        }

        private void cbbSentScript_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbSentScript.SelectedIndex != 0)
            {
                DataTable dtScript = ScriptService.GetScript(cbbSentScript.SelectedValue.ToString());
                if (dtScript.Rows.Count > 0 && int.Parse(dtScript.Rows[0]["LableCount"].ToString()) > 0)
                {
                    dgvSentScript.DataSource = null;
                    dgvSentScript.Refresh();
                    dgvSentScript.Columns.Clear();
                    dgvSentScript.ColumnCount = 1;
                    dgvSentScript.Columns[0].Name = "Temp";
                    dgvSentScript.Rows.Add(new string[] { "" });

                    //腳本敘述
                    lbDepiction.Text = dtScript.Rows[0].Field<string>("Depiction");

                    //建立標籤
                    string TextBoxContentsSet = dtScript.Rows[0].Field<string>("TextBoxContentsSet");
                    string LableContentSet = dtScript.Rows[0].Field<string>("LableContentSet");
                    List<string> WordSortList = TextBoxContentsSet.Split('｜').ToList();
                    List<string> LableSortList = LableContentSet.Split('｜').ToList();
                    LableSortList = LableSortList.OrderBy(o => o.Split('※')[1]).ToList();
                    int LableCount = dtScript.Rows[0].Field<int>("LableCount");
                    /*
                    string ErrorMsg = string.Empty;
                    DataTable dtRecord = RecordingService.SelectRecording(out ErrorMsg);
                    List<string> RecordNoArray = dtRecord.AsEnumerable().Select(s => s.Field<string>("ph_no")).ToList();

                    //List<string> TempRecordDataList = dtRecord.AsEnumerable().Select(s => s.Field<string>("ph_data")).ToList();
                    List<string> TempRecordDataListTemp = dtRecord.AsEnumerable().Select(s => s.Field<string>("ph_data")).ToList();
                    List<string> TempRecordDataList = new List<string>();
                    foreach (string item in TempRecordDataListTemp)
                    {
                        if (item == null || item == string.Empty)
                        {
                            TempRecordDataList.Add(string.Empty);
                            continue;
                        }
                        byte[] unknow = Encoding.GetEncoding(28591).GetBytes(item);
                        string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                        TempRecordDataList.Add(Big5);
                    }

                    List<string> TempRecordDataList2 = dtRecord.AsEnumerable().Select(s => s.Field<string>("ph_data2")).ToList();
                    List<LableService.cbbDataList> cbbRecordDataList = new List<LableService.cbbDataList>();
                    int TempCount = 0;
                    foreach (string item in TempRecordDataList)
                    {
                        int index = TempCount;
                        if (item == null || item.Trim() == string.Empty)
                        {
                            if (TempRecordDataList2[index] == null) { TempRecordDataList2[index] = string.Empty; }
                            cbbRecordDataList.Add(new LableService.cbbDataList { cbb_No = RecordNoArray[index], cbb_Data = RecordNoArray[index] + ":" + TempRecordDataList2[index] });
                        }
                        else
                        {
                            cbbRecordDataList.Add(new LableService.cbbDataList { cbb_No = RecordNoArray[index], cbb_Data = RecordNoArray[index] + ":" + item });
                        }
                        TempCount++;
                    }*/

                    //如果其中有出錯，代表此腳本標籤有被刪除，請使用者修改腳本
                    try
                    {
                        for (int i = 0; i < LableCount; i++)
                        {
                            DataTable dtLable = LableService.GetLable(LableSortList[i].Split('※')[0]);
                            if (dtLable.Rows.Count > 0)
                            {
                                LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), dtLable.Rows[0].Field<string>("Type"));
                                switch (LableEnum)
                                {
                                    //數字
                                    case LableEnum.Number:
                                        {
                                            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
                                            string ContentsSet = dtLable.Rows[0].Field<string>("ContentsSet");
                                            string LableName = dtLable.Rows[0].Field<string>("Name");
                                            string[] Contents = ContentsSet.Split('｜');
                                            combo.DataSource = Contents;
                                            combo.HeaderText = LableName;
                                            combo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                                            combo.Width = 500;
                                            dgvSentScript.Columns.Insert(i, combo);
                                            dgvSentScript.Rows[0].Cells[i].Value = Contents[0];
                                        }
                                        break;
                                    //一般下拉選單文字
                                    case LableEnum.General:
                                        {
                                            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
                                            string ContentsSet = dtLable.Rows[0].Field<string>("ContentsSet");
                                            string LableName = dtLable.Rows[0].Field<string>("Name");
                                            string[] Contents = ContentsSet.Split('｜');
                                            combo.DataSource = Contents;
                                            combo.HeaderText = LableName;
                                            combo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                                            combo.Width = 500;
                                            dgvSentScript.Columns.Insert(i, combo);
                                            dgvSentScript.Rows[0].Cells[i].Value = Contents[0];
                                        }
                                        break;
                                    //文字
                                    case LableEnum.Text:
                                        {
                                            DataGridViewTextBoxColumn TextBox = new DataGridViewTextBoxColumn();
                                            dgvSentScript.Columns.Insert(i, TextBox);
                                            List<string> Word = WordSortList.Select(s => s).Where(w => w.Split('※')[1] == (LableSortList[i].Split('※')[1]).ToString()).ToList();
                                            if (WordSortList.Count > 0)
                                            {
                                                dgvSentScript.Rows[0].Cells[i].Value = Word[0].Split('※')[0];
                                            }
                                            else { }
                                            dgvSentScript.Columns[i].HeaderText = "文字列" + (i + 1);
                                        }
                                        break;

                                    //錄音檔
                                    case LableEnum.Record:
                                        {
                                            //data get
                                            string RecordString = string.Empty;
                                            string ph_no = dtLable.Rows[0].Field<string>("ContentsSet");
                                            DataTable dtRecord = RecordingService.SelectRecording(ph_no.Trim());
                                            if (dtRecord.Rows.Count > 0)
                                            {
                                                string ph_data = dtRecord.Rows[0].Field<string>("ph_data");
                                                string ph_data2 = dtRecord.Rows[0].Field<string>("ph_data2");
                                                if (ph_data == null || ph_data.Trim() == string.Empty)
                                                {
                                                    if (ph_data2 == null) { ph_data2 = string.Empty; }
                                                    RecordString = "【" + ph_no.Trim() + "：" + ph_data2.Trim() + "】";
                                                }
                                                else
                                                {
                                                    byte[] unknow = Encoding.GetEncoding(28591).GetBytes(ph_data);
                                                    string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                                                    RecordString = "【" + ph_no.Trim() + "：" + Big5 + "】";
                                                }
                                            }

                                            DataGridViewTextBoxColumn TextBox = new DataGridViewTextBoxColumn();
                                            dgvSentScript.Columns.Insert(i, TextBox);
                                            dgvSentScript.Rows[0].Cells[i].Value = RecordString;
                                            dgvSentScript.Columns[i].HeaderText = "錄音檔";
                                            dgvSentScript.Columns[i].ReadOnly = true;
                                            /*
                                            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
                                            combo.DataSource = cbbRecordDataList;
                                            combo.HeaderText = "錄音檔";
                                            //combo.DisplayMember = "cbb_No";
                                            combo.DisplayMember = "cbb_Data";
                                            //combo.ValueMember = "cbb_Data";
                                            combo.ValueMember = "cbb_No";
                                            combo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                                            combo.Width = 500;
                                            dgvSentScript.Columns.Insert(i, combo);
                                            DataGridViewComboBoxCell cbbc = (DataGridViewComboBoxCell)dgvSentScript.Rows[0].Cells[i];
                                            //var dsffd = cbbRecordDataList.Where(w => w.cbb_No == dtLable.Rows[0].Field<string>("ContentsSet").ToString()).ToList();
                                            var dsffd = cbbRecordDataList.Where(w => w.cbb_No == dtLable.Rows[0].Field<string>("ContentsSet").ToString()).ToList();
                                            if (dsffd.Count() > 0)
                                            {
                                                //cbbc.Value = dsffd[0].cbb_Data;
                                                cbbc.Value = dsffd[0].cbb_No;
                                            }
                                            else
                                            {
                                                //cbbc.Value = cbbRecordDataList[0].cbb_Data;
                                                cbbc.Value = cbbRecordDataList[0].cbb_No;
                                                MessageBox.Show("找不到此腳本所對應的錄音檔，請檢查腳本相關設定。", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            }*/
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("該腳本無法取得所對應的標籤，請至該腳本修改內容。", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //刪除不要的欄位
                    dgvSentScript.Columns.Remove("Temp");

                    dgvSentScript.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dgvSentScript.RowHeadersVisible = false;
                    dgvSentScript.EditMode = DataGridViewEditMode.EditOnEnter;
                }
                else
                {
                    dgvSentScript.DataSource = null;
                    dgvSentScript.Refresh();
                    dgvSentScript.Columns.Clear();
                }
            }
        }

        //換頁籤事件(重置Cobobox DataSource)
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int WorkPage = tabControl1.SelectedIndex;
            switch (WorkPage)
            {
                case 0:
                    break;
                case 1:
                    break;
                case 2:
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 6:
                    break;

            }
        }

        private void btnProduce_Click(object sender, EventArgs e)
        {
            int DataGridViewRowCount = dgvSentScript.Rows.Count;

            //檢查dgv現在是否有資料
            if (DataGridViewRowCount > 0)
            {
                int CellCount = dgvSentScript.Rows[0].Cells.Count;
                string SentMessage = string.Empty;
                for (int i = 0; i < CellCount; i++)
                {
                    if (dgvSentScript.Columns[i].HeaderText == "錄音檔")
                    {
                        //DataGridViewComboBoxCell dgvccb = (DataGridViewComboBoxCell)dgvSentScript.Rows[0].Cells[i];
                        //SentMessage = SentMessage + "【" + dgvccb.FormattedValue + dgvccb.Value + "】";
                        //SentMessage = SentMessage + "【" + dgvccb.FormattedValue.ToString().Substring(0, 6) + dgvccb.FormattedValue.ToString().Substring(7, dgvccb.FormattedValue.ToString().Length - 7) + "】";
                        string recordText = dgvSentScript.Rows[0].Cells[i].Value.ToString();
                        SentMessage = SentMessage + recordText;
                    }
                    else
                    {
                        SentMessage = SentMessage + dgvSentScript.Rows[0].Cells[i].Value;
                    }
                }

                //產生目前腳本訊息
                tbSentMessage.Text = SentMessage;
            }
        }

        private void btnSent_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("確定要發送資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult == DialogResult.Yes)
            {
                if (_tdSent == null || _tdSent.IsAlive == false)
                {
                    //檢查目前與Client連現狀控
                    _tdSent = new Thread(Sent);
                    _tdSent.IsBackground = true;
                    _tdSent.Start();
                }
                else
                {
                    MessageBox.Show("資料發送中請稍後", "Message", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
                return;
        }

        private void Sent()
        {
            //鎖住CheckBox
            EnbleCheckBox(false, cbSentSMS);
            EnbleCheckBox(false, cbSentTTS);

            //清空狀態視窗
            ClearTextBox("", tbSentMsgStatus);
            updateTextBox("-----------------------Start-----------------------" + Environment.NewLine, tbSentMsgStatus);

            //檢查是否有資料
            if (tbSentMessage.Text.Trim() != string.Empty)
            {
                //Get Group telephone numbers
                DataTable dtTelephoneNumber = new DataTable();
                foreach (DataGridViewRow row in this.dgvGroup.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == "True")
                    {
                        DataTable dtTempTelephoneNumber = new DataTable();
                        dtTempTelephoneNumber = EmployeeService.GetEmployee(row.Cells["Id"].Value.ToString());
                        dtTempTelephoneNumber.PrimaryKey = new DataColumn[] { dtTempTelephoneNumber.Columns["EmployeeId"] };
                        dtTelephoneNumber.PrimaryKey = new DataColumn[] { dtTelephoneNumber.Columns["EmployeeId"] };
                        dtTelephoneNumber.Merge(dtTempTelephoneNumber, true);
                    }
                }

                //發語音訊息
                if (cbSentTTS.Checked == true)
                {
                    updateTextBox("開始發送TTS語音訊息..." + Environment.NewLine, tbSentMsgStatus);

                    //發送給每個使用者
                    foreach (DataRow item in dtTelephoneNumber.Rows)
                    {
                        //組合Url
                        string CalloutphpUrl = GetCalloutphpUrl(item.Field<string>("PhoneNumber"), item.Field<string>("Name"), string.Empty).ToString();

                        //Get WebService發送訊息
                        _textbox = tbSentMsgStatus;
                        GetRequest(CalloutphpUrl, item.Field<string>("PhoneNumber"));
                    }
                }

                //發簡訊
                if (cbSentSMS.Checked == true)
                {
                    updateTextBox("開始發送SMS簡訊訊息..." + Environment.NewLine, tbSentMsgStatus);
                    DoSentSMS(dtTelephoneNumber);
                    updateTextBox("發送SMS簡訊訊息結束..." + Environment.NewLine, tbSentMsgStatus);
                }

                //發送結束儲存dgvGroup現在狀態
                List<string> GroupIdList = new List<string>();
                List<string> GroupNameList = new List<string>();
                foreach (DataGridViewRow row in this.dgvGroup.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[0].Value.ToString() == "True")
                    {
                        GroupIdList.Add(row.Cells["Id"].Value.ToString());
                        GroupNameList.Add(row.Cells["Name"].Value.ToString());
                    }
                }
                //使用DAO的方式insert資料，使用此方法資料能夠更快速匯入
                //SentDataGridViewStatusService.DAOInsertSentDataGridViewStatus(GroupIdList, GroupNameList);

                //已改成MS SQL
                SentDataGridViewStatusService.InsertSentDataGridViewStatus(GroupIdList, GroupNameList);
            }
            else
            {
                updateTextBox("目前沒資料發送..." + Environment.NewLine, tbSentMsgStatus);
            }
            updateTextBox("-----------------------END-----------------------" + Environment.NewLine, tbSentMsgStatus);

            //解鎖CheckBox
            EnbleCheckBox(true, cbSentSMS);
            EnbleCheckBox(true, cbSentTTS);
        }

        private StringBuilder GetCalloutphpUrl(string TelephoneNumber, string Name, string AutoMsg)
        {
            StringBuilder result = new StringBuilder();

            //網址
            string CalloutphpServerIp = ConfigurationManager.AppSettings["CalloutphpServerIp"];
            result.Append("http://" + CalloutphpServerIp + "/calloutphp/getcallout.php?callout_sn=");
            //序號年月日時分秒+001流水序
            result.Append(DateTime.Now.ToString("yyyyMMddHHmmss") + "001");
            //username
            result.Append("&username=" + Name);
            //phone number
            result.Append("&tel_no=" + TelephoneNumber);
            //內容
            result = GetTtsRecordMsg(result, AutoMsg);
            //note(接收回覆用)
            result.Append("&note=1=1,2=1,3=1,*=1&tts=utf8");

            return result;
        }

        private StringBuilder GetTtsRecordMsg(StringBuilder sb, string AutoMsg)
        {
            StringBuilder result = sb;

            if (AutoMsg == string.Empty)
            {
                //取得錄音檔ph_no
                string[] TempArray = tbSentMessage.Text.Trim().Split('【');
                List<string> RecordNoList = new List<string>();
                for (int i = 1; i < TempArray.Count(); i++)
                {
                    RecordNoList.Add(TempArray[i].Split('】')[0].Substring(0, 6));
                }
                //取得文字List
                List<string> MsgList = new List<string>();
                if (TempArray.Count() > 0)
                {
                    //濾掉所有特殊符號
                    MsgList.Add(Regex.Replace(TempArray[0].Trim(), @"[\W_]+", ""));
                    for (int i = 1; i < TempArray.Count(); i++)
                    {
                        if (TempArray[i].Split('】')[1].Trim() != string.Empty)
                        {
                            //濾掉所有特殊符號
                            MsgList.Add(Regex.Replace(TempArray[i].Split('】')[1].Trim(), @"[\W_]+", ""));
                        }
                    }
                }
                else
                {
                    //濾掉所有特殊符號
                    string NowMsg = @tbSentMessage.Text.Trim();
                    string ValidMsg = Regex.Replace(NowMsg, @"[\W_]+", "");
                    MsgList.Add(ValidMsg);
                }

                //塞入文字與錄音訊息
                result.Append("&play=");
                int Count = RecordNoList.Count + MsgList.Count;
                for (int i = 0; i < MsgList.Count; i++)
                {
                    //防止文字為空
                    if (MsgList[i] != string.Empty)
                    {
                        result.Append("$0[3]=" + MsgList[i] + ",");
                    }
                    //防止以文字結尾的話錄音次數會少一次，導致錯誤
                    if (RecordNoList.Count > i)
                    {
                        result.Append(RecordNoList[i] + ",");
                    }
                }
            }

            //加入31處無人發報功能
            if (AutoMsg != string.Empty)
            {
                result.Append("&play=");
                result.Append("$0[3]=" + AutoMsg + ",");
            }

            return result;
        }

        //用Get的方式Call Web Service
        async static void GetRequest(string url, string TelephoneNumber)
        {
            using (HttpClient Client = new HttpClient())
            {
                try
                {
                    using (HttpResponseMessage response = await Client.GetAsync(url))
                    {
                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();
                        if (responseBody == "DATAOK")
                        {
                            Print(_textbox, "發送TTS語音訊息至" + TelephoneNumber + "成功!!" + Environment.NewLine);
                        }
                        if (responseBody == "DATAERR")
                        {
                            Print(_textbox, "發送TTS語音訊息至" + TelephoneNumber + "失敗!!" + Environment.NewLine);
                        }
                    }
                }
                catch (Exception ex)
                {
                    string ErrorMsg = ex.Message;
                    Print(_textbox, "發送TTS語音訊息至" + TelephoneNumber + "失敗!!" + Environment.NewLine);
                    Print(_textbox, "失敗原因：" + ErrorMsg + Environment.NewLine);
                }

            }
        }

        public static void Print(TextBox tb, string text)
        {
            //判斷這個TextBox的物件是否在同一個執行緒上
            if (tb.InvokeRequired)
            {
                //當InvokeRequired為true時，表示在不同的執行緒上，所以進行委派的動作!!
                PrintHandler ph = new PrintHandler(Print);
                tb.Invoke(ph, tb, text);
            }
            else
            {
                //表示在同一個執行緒上了，所以可以正常的呼叫到這個TextBox物件
                tb.AppendText(text);
            }
        }

        //SentSMS
        private void DoSentSMS(DataTable dtTelephoneNumber)
        {
            //連線
            string ServerIp = string.Empty;
            string ServerPort = string.Empty;
            string UserID = string.Empty;
            string Passwd = string.Empty;
            int ret_code;
            string ret_description = string.Empty;
            string Message = string.Empty;

            //Get Appconfig data
            //Set SMS Login data
            ServerIp = ConfigurationManager.AppSettings["ServerIp"];
            ServerPort = ConfigurationManager.AppSettings["ServerPort"];
            UserID = ConfigurationManager.AppSettings["UserID"];
            Passwd = ConfigurationManager.AppSettings["Passwd"];

            //Get註冊後的HiAir.dll，並且使用動態配置的方式宣告物件使用
            dynamic dymSMS = Activator.CreateInstance(Type.GetTypeFromProgID("HiAir.HiNetSMS"));
            //連線中華電信SMS Server
            ret_code = dymSMS.StartCon(ServerIp, ServerPort, UserID, Passwd);
            //Set data
            Message = GetSmsString();

            //表示成功連上中華電信SMS Server
            if (ret_code == 0)
            {
                updateTextBox("中華電信SMS Server連線成功!!" + Environment.NewLine, tbSentMsgStatus);

                //發送給每個使用者
                foreach (DataRow item in dtTelephoneNumber.Rows)
                {
                    string Tel = string.Empty;

                    //Get data
                    Tel = item.Field<string>("PhoneNumber");

                    //發送
                    ret_code = dymSMS.SendMsg(Tel, Message.ToString());
                    if (ret_code == 0)
                    {
                        updateTextBox("發送簡訊至" + Tel + "成功!!" + Environment.NewLine, tbSentMsgStatus);
                    }
                    else
                    {
                        ret_description = dymSMS.QueryMsg();
                        updateTextBox("發送簡訊至" + Tel + "失敗!!" + Environment.NewLine, tbSentMsgStatus);
                        updateTextBox("失敗原因：" + Environment.NewLine + "", tbSentMsgStatus);
                        updateTextBox(ret_description + Environment.NewLine, tbSentMsgStatus);
                    }

                }
                //結束時關閉連線
                dymSMS.EndCon();
            }
            else
            {
                ret_description = dymSMS.Get_Message();
                updateTextBox("中華電信SMS Server連線失敗!!" + Environment.NewLine, tbSentMsgStatus);
                updateTextBox("失敗原因：" + Environment.NewLine + "", tbSentMsgStatus);
                updateTextBox(ret_description + Environment.NewLine, tbSentMsgStatus);
            }
        }

        //組合SMS所需的字串
        private string GetSmsString()
        {
            StringBuilder Message = new StringBuilder();
            string[] TempArray = tbSentMessage.Text.Trim().Split('【');
            List<string> RecordNoList = new List<string>();
            for (int i = 1; i < TempArray.Count(); i++)
            {
                string Temp = TempArray[i].Split('】')[0];
                RecordNoList.Add(Temp.Substring(6));
            }
            //取得文字List
            List<string> MsgList = new List<string>();
            if (TempArray.Count() > 0)
            {
                //塞入第一筆資料
                MsgList.Add(TempArray[0].Trim());
                for (int i = 1; i < TempArray.Count(); i++)
                {
                    if (TempArray[i].Split('】')[1].Trim() != string.Empty)
                    {
                        //塞入資料
                        MsgList.Add(TempArray[i].Split('】')[1].Trim());
                    }
                }
            }
            else
            {
                //如都沒有錄音檔直接塞入
                MsgList.Add(tbSentMessage.Text.Trim());
            }

            //塞入文字與錄音訊息
            int Count = RecordNoList.Count + MsgList.Count;
            for (int i = 0; i < MsgList.Count; i++)
            {
                //防止文字為空
                if (MsgList[i] != string.Empty)
                {
                    Message.Append(MsgList[i]);
                }
                //防止以文字結尾的話錄音次數會少一次，導致錯誤
                if (RecordNoList.Count > i)
                {
                    Message.Append(RecordNoList[i]);
                }
            }

            return Message.ToString();
        }

        //讀取資料庫最後一次發送的群組並且更新現在群組選取狀況
        private void InitializeSentDataGridViewGroupStatus()
        {
            //Set Gata
            DataTable SentDataGridViewStatus = SentDataGridViewStatusService.SelectSentDataGridViewStatus();
            if (SentDataGridViewStatus.Rows.Count > 0)
            {
                foreach (DataRow item in SentDataGridViewStatus.Rows)
                {
                    string NowGroupId = item["GroupId"].ToString();

                    foreach (DataGridViewRow row in this.dgvGroup.Rows)
                    {
                        if ((row.Cells["Id"].Value != null) && (row.Cells["Id"].Value.ToString() == NowGroupId))
                        {
                            row.Cells[0].Value = true;
                            break;
                        }
                    }
                }
            }
        }
        //-------------------------------------------Employee------------------------------------------------------------
        private void btnEmployeeFileImport_Click(object sender, EventArgs e)
        {
            //Select CSV File
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "csv files (*.*)|*.csv";
            string ErrorMsg = string.Empty;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                AsyncTemplate.DoWorkAsync(
               () =>
               {
                   ErrorMsg = doImportCsv(dialog.FileName);
               },
               () =>
               {
                   //MessageBox.Show("Success, Result is " + result.ToString());
                   if (ErrorMsg == string.Empty)
                   {
                       MessageBox.Show("匯入完成", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   }
                   else
                   {
                       MessageBox.Show(ErrorMsg, "匯入失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                   }

                   //刪除狀態
                   SentDataGridViewStatusService.DeleteSentDataGridViewStatus();
               },
               (exception) =>
               {
                   //MessageBox.Show(exception.Message);
                   MessageBox.Show("失敗原因；" + exception.Message, "匯出失敗", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   //error handling
               });
            }

            //初始化畫面
            InitializeControl();
            InitializeEmpolyeeCcbGroup();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //Get data
                string Path = dialog.SelectedPath;
                DataTable dtEmployee = EmployeeService.GetEmployee();
                //dtEmployee.Columns["GroupName"].SetOrdinal(1);
                dtEmployee.Columns["GroupName"].SetOrdinal(0);
                dtEmployee.Columns.Remove("GroupId");
                dtEmployee.Columns["EmployeeId"].ColumnName = "員工編號";
                dtEmployee.Columns["GroupName"].ColumnName = "所屬群組";
                dtEmployee.Columns["Name"].ColumnName = "姓名";
                dtEmployee.Columns["PhoneNumber"].ColumnName = "電話號碼";

                //Set csv
                if (dtEmployee.Rows.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();

                    IEnumerable<string> columnNames = dtEmployee.Columns.Cast<DataColumn>().
                                                      Select(column => column.ColumnName);
                    sb.AppendLine(string.Join(",", columnNames));

                    foreach (DataRow row in dtEmployee.Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }

                    File.WriteAllText(Path + "\\語音系統群組聯絡人匯出" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv", sb.ToString(), System.Text.Encoding.UTF8);

                    MessageBox.Show("匯出完成", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    MessageBox.Show("目前群組聯絡人無資料", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }


        }

        private DataTable GetCsvDb(string CsvPath)
        {
            DataTable dt = new DataTable("NewTable");
            DataRow row;
            using (TextFieldParser parser = new TextFieldParser(CsvPath))
            {

                int i = 0;
                string[] strLine;
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    try
                    {
                        //Processing row
                        strLine = parser.ReadFields();
                        if (i == 0)
                        {
                            for (int j = 0; j < strLine.Length; j++)
                            {
                                dt.Columns.Add(strLine[j]);
                            }
                        }
                        else
                        {
                            row = dt.NewRow();
                            row.ItemArray = strLine;
                            dt.Rows.Add(row);
                        }
                        i++;
                    }
                    catch (Exception ex)
                    {
                        //ex.Message;
                    }
                }
            }
            return dt;
        }

        private void InitializeAsyncForm()
        {
            MyWaitForm MyWaitForm = new MyWaitForm();

            AsyncTemplate.OnInvokeStarting =
                () =>
                {
                    MyWaitForm.ShowDialog();
                };

            AsyncTemplate.OnInvokeEnding =
                () =>
                {
                    if (MyWaitForm.InvokeRequired)
                    {
                        MyWaitForm.Invoke(new MethodInvoker(
                            ()
                            =>
                            {
                                MyWaitForm.Close();
                            })
                        );
                    }
                    else
                    {
                        MyWaitForm.Close();
                    }
                };
        }

        private string doImportCsv(string FileName)
        {
            string ErrorMsg = string.Empty;
            string CrvPath = FileName;
            //Get data
            DataTable CsvDb = GetCsvDb(CrvPath);
            if (CsvDb.Rows.Count > 0)
            {
                //Set data
                //List<string> NewGroupList = CsvDb.AsEnumerable().Select(r => r.Field<string>(1)).Distinct().ToList();
                List<string> NewGroupList = CsvDb.AsEnumerable().Select(r => r.Field<string>(0)).Distinct().ToList();

                //Insert data
                //DataTable dtGroup = GroupService.DAOInsertGroup((List<string>)NewGroupList);
                GroupService.DeleteGroup();
                ErrorMsg = GroupService.MultipleInsertGroup((List<string>)NewGroupList);
                if (ErrorMsg == string.Empty)
                {
                    DataTable dtGroup = GroupService.SelectGroupAll();
                    EmployeeService.DeleteEmployee();
                    //Insert Employees
                    int csvCount = CsvDb.Rows.Count;
                    for (int i = 0; i < (csvCount / 400) + 1; i++)
                    {
                        DataTable rows = CsvDb.AsEnumerable().Skip(i * 400).Take(400).CopyToDataTable();
                        ErrorMsg = EmployeeService.MultipleInsertEmployee(rows, dtGroup);
                        if (ErrorMsg != string.Empty)
                        {
                            return ErrorMsg;
                        }
                    }
                }
            }
            return ErrorMsg;
        }

        //-------------------------------------------Log------------------------------------------------------------
        private void btnLogSelect_Click(object sender, EventArgs e)
        {
            //Get data
            DateTime SelectStartDate = dtpLogStartDate.Value.Date;
            DateTime SelectEndDate = dtpLogEndDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);
            double StartDate = DateParseLong(SelectStartDate);
            double EndDate = DateParseLong(SelectEndDate);
            string ErrorMsg = string.Empty;
            //Select data
            DataTable dtLogSelect = LogService.SelectLog(StartDate, EndDate, out ErrorMsg);

            //Show data
            if (dtLogSelect.Rows.Count > 0)
            {
                //新增姓名欄
                DataColumn column = new DataColumn("姓名");
                column.DataType = System.Type.GetType("System.String");
                column.AllowDBNull = true;
                column.Caption = "姓名";
                dtLogSelect.Columns.Add(column);

                //新增時間欄
                DataColumn column1 = new DataColumn("撥出日期時間");
                column1.DataType = System.Type.GetType("System.String");
                column1.AllowDBNull = true;
                column1.Caption = "撥出日期時間";
                dtLogSelect.Columns.Add(column1);

                //新增結果欄
                DataColumn column2 = new DataColumn("外撥結果");
                column2.DataType = System.Type.GetType("System.String");
                column2.AllowDBNull = true;
                column2.Caption = "外撥結果";
                dtLogSelect.Columns.Add(column2);

                foreach (DataRow row in dtLogSelect.Rows)
                {
                    //姓名
                    string dsf = row.Field<string>("N_USERNAME");
                    byte[] unknow = Encoding.GetEncoding(28591).GetBytes(dsf);
                    string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                    row["姓名"] = Big5;

                    //時間
                    long value = long.Parse(row["N_TMS"].ToString() + "0000000");
                    TimeSpan ts = new TimeSpan(value);
                    DateTime dt = new DateTime(1970, 1, 1).AddHours(8);
                    string dtResult = (dt + ts).ToString("yyyy/MM/dd HH:mm:ss");
                    row["撥出日期時間"] = dtResult;

                    //結果
                    string ret = row["N_RET"].ToString();
                    switch (ret)
                    {
                        case "0":
                            row["外撥結果"] = "未通報";
                            break;
                        case "1":
                            row["外撥結果"] = "通報中";
                            break;
                        case "2":
                            row["外撥結果"] = "成功";
                            break;
                        case "3":
                            row["外撥結果"] = "失敗(忙線)";
                            break;
                        case "4":
                            row["外撥結果"] = "失敗(無人接聽)";
                            break;
                        case "5":
                            row["外撥結果"] = "失敗(線路異常)";
                            break;
                        case "6":
                            row["外撥結果"] = "失敗(資料異常)";
                            break;
                        case "7":
                            row["外撥結果"] = "失敗(超出時段)";
                            break;
                        case "8":
                            row["外撥結果"] = "失敗(無效電話)";
                            break;
                        case "9":
                            row["外撥結果"] = "失敗(強制中斷)";
                            break;
                        default:
                            break;
                    }

                }
                dgvLogResult.DataSource = null;
                dgvLogResult.Refresh();
                dgvLogResult.Columns.Clear();
                dgvLogResult.DataSource = dtLogSelect;
                dgvLogResult.AutoResizeColumns();
                dgvLogResult.Columns["撥出日期時間"].DisplayIndex = 0;
                dgvLogResult.Columns["姓名"].DisplayIndex = 1;
                dgvLogResult.Columns["N_CALLNO"].DisplayIndex = 2;
                dgvLogResult.Columns["N_CALLNO"].HeaderText = "電話號碼";
                dgvLogResult.Columns["外撥結果"].DisplayIndex = 3;
                dgvLogResult.Columns["N_TMS"].Visible = false;
                dgvLogResult.Columns["N_USERNAME"].Visible = false;
                dgvLogResult.Columns["N_RET"].Visible = false;
                dgvLogResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvLogResult.ReadOnly = true;
                dgvLogResult.AllowUserToAddRows = false;

                lblLogTotleCount.Text = "共" + dtLogSelect.Rows.Count + "筆資料";
            }
            else
            {

                //沒有資料
                MessageBox.Show("此區間無任何資料", "搜尋成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                dgvLogResult.DataSource = null;
                dgvLogResult.Refresh();
                dgvLogResult.Columns.Clear();
            }

            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private double DateParseLong(DateTime dt)
        {
            double result = 0;
            result = (TimeZoneInfo.ConvertTimeToUtc(dt) - new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc)).TotalSeconds;

            return result;
        }

        private void btnLogOutput_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //Get data
                string FileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "LogQuery.csv";
                string Path = dialog.SelectedPath + '\\' + FileName;
                DataGridView dgv = dgvLogResult;

                try
                {
                    dgv.Columns.Remove("N_TMS");
                    dgv.Columns.Remove("N_USERNAME");
                    dgv.Columns.Remove("N_RET");
                }
                catch (Exception ex) { }

                writeCSV(dgv, @Path);
            }
        }

        public void writeCSV(DataGridView gridIn, string outputFile)
        {
            //test to see if the DataGridView has any rows
            if (gridIn.RowCount > 0)
            {
                string value = "";
                DataGridViewRow dr = new DataGridViewRow();
                StreamWriter swOut = new StreamWriter(outputFile, true, System.Text.Encoding.GetEncoding("UTF-8"));

                //write header rows to csv
                for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
                {
                    if (i > 0)
                    {
                        swOut.Write(",");
                    }
                    swOut.Write(gridIn.Columns[i].HeaderText);
                }

                swOut.WriteLine();

                //write DataGridView rows to csv
                for (int j = 0; j <= gridIn.Rows.Count - 1; j++)
                {
                    if (j > 0)
                    {
                        swOut.WriteLine();
                    }

                    dr = gridIn.Rows[j];

                    for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
                    {
                        if (i > 0)
                        {
                            swOut.Write(",");
                        }

                        value = dr.Cells[i].Value.ToString();
                        //replace comma's with spaces
                        value = value.Replace(',', ' ');
                        //replace embedded newlines with spaces
                        value = value.Replace(Environment.NewLine, " ");

                        swOut.Write(value);
                    }
                }
                swOut.Close();
                MessageBox.Show("匯出完成", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("目前無查詢資料", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        //------------------------------------------Recording----------------------------------------------------------
        private void InitializeRecordingForm()
        {
            tbRecordPhraseNo.Text = string.Empty;
            tbRecordPhraseExcerpt.Text = string.Empty;

            //Get data
            string ErrorMsg = string.Empty;
            DataTable dtRecording = RecordingService.SelectRecording(out ErrorMsg);

            //Data change
            if (dtRecording.Rows.Count > 0)
            {
                //新增片語摘要欄
                DataColumn column = new DataColumn("片語摘要");
                column.DataType = System.Type.GetType("System.String");
                column.AllowDBNull = true;
                column.Caption = "片語摘要";
                dtRecording.Columns.Add(column);

                foreach (DataRow row in dtRecording.Rows)
                {
                    //姓名
                    string ph_data = row.Field<string>("ph_data");
                    string ph_data2 = row.Field<string>("ph_data2");
                    if (ph_data == null || ph_data.Trim() == string.Empty)
                    {
                        if (ph_data2 == null) { ph_data2 = string.Empty; }
                        row["片語摘要"] = ph_data2;
                    }
                    else
                    {
                        byte[] unknow = Encoding.GetEncoding(28591).GetBytes(ph_data);
                        string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                        row["片語摘要"] = Big5;
                    }
                }
            }
            //Set data
            if (dtRecording.Rows.Count > 0)
            {
                dgvRecordContent.DataSource = null;
                dgvRecordContent.Refresh();
                dgvRecordContent.Columns.Clear();
                dgvRecordContent.DataSource = dtRecording;

                //新增試聽欄
                DataGridViewDisableButtonColumn btncListen = new DataGridViewDisableButtonColumn();
                btncListen.HeaderText = "試聽";
                btncListen.Name = "試聽";
                btncListen.Text = "播放";
                btncListen.UseColumnTextForButtonValue = true;
                dgvRecordContent.Columns.Add(btncListen);

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvRecordContent.Columns.Add(btncDelete);

                dgvRecordContent.Columns["ph_no"].HeaderText = "代碼";
                dgvRecordContent.Columns["ph_data"].Visible = false;
                dgvRecordContent.Columns["ph_data2"].Visible = false;
                dgvRecordContent.AutoResizeColumns();
                dgvRecordContent.Columns["ph_no"].Width = 100;
                dgvRecordContent.Columns["片語摘要"].Width = 720;
                dgvRecordContent.Columns["試聽"].Width = 60;
                dgvRecordContent.Columns["功能"].Width = 60;
                dgvRecordContent.ReadOnly = true;
                dgvRecordContent.RowTemplate.Height = 30;
                dgvRecordContent.AllowUserToAddRows = false;
            }

            //Message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //檢查錄音檔
            foreach (DataGridViewRow row in this.dgvRecordContent.Rows)
            {
                if (row.Cells["ph_no"].Value != null)
                {
                    string url = ConfigurationManager.AppSettings["RecordingPath"] + row.Cells["ph_no"].Value.ToString() + ".1";
                    if (!System.IO.File.Exists(url))
                    {
                        ((DataGridViewDisableButtonCell)row.Cells["試聽"]).Enabled = false;
                    }
                }
            }
            dgvRecordContent.Refresh();

        }

        private void dgvRecordContent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            //試聽
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 4)
            {
                DataGridViewDisableButtonCell buttonCell = (DataGridViewDisableButtonCell)dgvRecordContent.Rows[e.RowIndex].Cells["試聽"];
                if (buttonCell.Enabled == true)
                {
                    string url = ConfigurationManager.AppSettings["RecordingPath"] + dgvRecordContent.Rows[e.RowIndex].Cells["ph_no"].Value.ToString() + ".1";
                    MediaPlayer MediaPlayerForm = new MediaPlayer();
                    MediaPlayerForm.FilePath = url;
                    MediaPlayerForm.ShowDialog();
                }
            }

            //功能刪除
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 5)
            {
                var confirmResult = MessageBox.Show("確定要刪除此片語資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string ErrorMsg = string.Empty;
                if (confirmResult == DialogResult.Yes)
                {
                    //Get data
                    string RecordPhraseNo = dgvRecordContent.Rows[e.RowIndex].Cells["ph_no"].Value.ToString();
                    //Delete data
                    if (RecordPhraseNo.Trim() != null && RecordPhraseNo.Trim() != string.Empty)
                    {
                        ErrorMsg = RecordingService.DeleteRecording(RecordPhraseNo);
                    }
                }
                else
                    return;

                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //刪除錄音檔
                    string url = ConfigurationManager.AppSettings["RecordingPath"] + dgvRecordContent.Rows[e.RowIndex].Cells["ph_no"].Value.ToString() + ".1";
                    if (System.IO.File.Exists(url))
                    {
                        File.Delete(url);
                    }
                    MessageBox.Show("刪除成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializeRecordingForm();
                }
            }
        }

        private void btnRecordingReset_Click(object sender, EventArgs e)
        {
            InitializeRecordingForm();
        }

        private void dgvRecordContent_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            tbRecordPhraseNo.Text = e.Row.Cells["ph_no"].Value.ToString();
            tbRecordPhraseExcerpt.Text = e.Row.Cells["片語摘要"].Value.ToString();
        }


        private void btnRecordAdd_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string RecordPhraseNo = tbRecordPhraseNo.Text.Trim();
            string RecordPhraseExcerpt = tbRecordPhraseExcerpt.Text.Trim().Replace("｜", "").Replace("※", "").Replace("【", "").Replace("】", "");

            //Check data
            //是否空白
            if (RecordPhraseNo == string.Empty || RecordPhraseExcerpt == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」與「片語摘要」不能空白\n";
            }
            //檢查是否為數字
            int result = 0;
            if (int.TryParse(RecordPhraseNo, out result) == false)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」只能為6位的數字\n";
            }
            //片語代碼6位數
            if (RecordPhraseNo.Length != 6)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」只限6位數\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = RecordingService.InsertRecording(RecordPhraseNo, RecordPhraseExcerpt);
            }


            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializeRecordingForm();
            }

        }

        private void btnRecordUpdate_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string RecordPhraseNo = tbRecordPhraseNo.Text.Trim();
            string RecordPhraseExcerpt = tbRecordPhraseExcerpt.Text.Trim().Replace("｜", "").Replace("※", "").Replace("【", "").Replace("】", "");

            //Check data
            //是否空白
            if (RecordPhraseNo == string.Empty || RecordPhraseExcerpt == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」與「片語摘要」不能空白\n";
            }
            //檢查是否為數字
            int result = 0;
            if (int.TryParse(RecordPhraseNo, out result) == false)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」只能為6位的數字\n";
            }

            //片語代碼6位數
            if (RecordPhraseNo.Length != 6)
            {
                ErrorMsg = ErrorMsg + "「片語代碼」只能為6位的數字\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = RecordingService.UpdateRecording(RecordPhraseNo, RecordPhraseExcerpt);
            }


            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("更新成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializeRecordingForm();
                tbRecordPhraseNo.Text = string.Empty;
                tbRecordPhraseExcerpt.Text = string.Empty;
            }
        }
        //------------------------------------------Group----------------------------------------------------------
        private void btnGroupSelect_Click(object sender, EventArgs e)
        {
            InitializedgvGroupContentForm();
        }

        private void InitializedgvGroupContentForm()
        {
            //Get data
            DataTable dtGroup = GroupService.SelectGroupAll();

            //Set data
            if (dtGroup.Rows.Count > 0)
            {
                dgvGroupContent.DataSource = null;
                dgvGroupContent.Refresh();
                dgvGroupContent.Columns.Clear();
                dgvGroupContent.DataSource = dtGroup;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvGroupContent.Columns.Add(btncDelete);

                dgvGroupContent.Columns["Id"].Visible = false;
                dgvGroupContent.Columns["CreateTime"].Visible = false;
                dgvGroupContent.Columns["Name"].HeaderText = "群組名稱";
                dgvGroupContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvGroupContent.Columns["功能"].Width = 60;
                dgvGroupContent.AutoResizeColumns();
                dgvGroupContent.ReadOnly = true;
                dgvGroupContent.RowTemplate.Height = 30;
                dgvGroupContent.AllowUserToAddRows = false;

                tbGroupId.Text = string.Empty;
                tbGroupName.Text = string.Empty;
            }
            else
            {
                dgvGroupContent.DataSource = null;
                dgvGroupContent.Refresh();
                dgvGroupContent.Columns.Clear();
            }
        }

        private void dgvGroupContent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            //功能刪除
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 3)
            {
                var confirmResult = MessageBox.Show("確定要刪除此筆資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string ErrorMsg = string.Empty;
                if (confirmResult == DialogResult.Yes)
                {
                    //Get data
                    string GroupId = dgvGroupContent.Rows[e.RowIndex].Cells["Id"].Value.ToString();
                    //Delete data
                    if (GroupId.Trim() != null && GroupId.Trim() != string.Empty)
                    {
                        ErrorMsg = GroupService.DeleteGroup(GroupId, ErrorMsg);
                    }
                }
                else
                    return;

                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("刪除成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializedgvGroupContentForm();
                    InitializeControl();
                    //InitializeSentDataGridViewGroupStatus();
                    InitializeEmpolyeeCcbGroup();
                }
            }
        }

        private void dgvGroupContent_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            tbGroupId.Text = e.Row.Cells["Id"].Value.ToString();
            tbGroupName.Text = e.Row.Cells["Name"].Value.ToString();
        }

        private void btnGroupAdd_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string GroupName = tbGroupName.Text.Trim();

            //Check data
            //是否空白
            if (GroupName == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「群組名稱」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = GroupService.InsertGroup(GroupName);
            }


            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvGroupContentForm();
                InitializeControl();
                //InitializeSentDataGridViewGroupStatus();
                InitializeEmpolyeeCcbGroup();
            }
        }

        private void btnGroupUpdate_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;

            //Get data
            string GroupName = tbGroupName.Text.Trim();
            string GroupId = tbGroupId.Text.Trim();

            //Check data
            //是否空白
            if (GroupName == string.Empty || GroupId == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「群組編號」與「群組名稱」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {

                DataTable SelectDb = GroupService.SelectGroup(GroupId);
                string SelectName = SelectDb.Rows[0]["Name"].ToString();
                string SelectId = SelectDb.Rows[0]["Id"].ToString();
                var confirmResult = MessageBox.Show("確定要將「" + SelectName + "」更改成「" + GroupName + "」嗎?\n",
                                         "Check",
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmResult == DialogResult.Yes)
                {
                    ErrorMsg = GroupService.UpdateGroup(GroupId, GroupName);
                }
                else
                    return;
            }

            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("更新成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvGroupContentForm();
                InitializeControl();
                //InitializeSentDataGridViewGroupStatus();
                InitializeEmpolyeeCcbGroup();
            }

        }
        //---------------------------------------------------Empolyee------------------------------------------------------
        private void btnEmployeeSelect_Click(object sender, EventArgs e)
        {
            //Get data
            string EmployeeId = tbSelectEmployeeId.Text.Trim();
            string EmployeeName = tbSelectEmployeeName.Text.Trim();
            string GroupId = cbbSelectEmployeeGroupName.SelectedValue.ToString();
            string PhoneNumber = tbSelectEmployeePhoneNumber.Text.Trim();
            DataTable dtEmployee = new DataTable();

            if (EmployeeId.Trim() == string.Empty && EmployeeName.Trim() == string.Empty && GroupId.Trim() == "0" && PhoneNumber.Trim() == string.Empty)
                dtEmployee = EmployeeService.GetEmployee();
            else
                dtEmployee = EmployeeService.GetEmployee(EmployeeId, GroupId, EmployeeName, PhoneNumber);

            //Set data
            if (dtEmployee.Rows.Count > 0)
            {
                dgvEmployeeContent.DataSource = null;
                dgvEmployeeContent.Refresh();
                dgvEmployeeContent.Columns.Clear();
                dgvEmployeeContent.DataSource = dtEmployee;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvEmployeeContent.Columns.Add(btncDelete);

                dgvEmployeeContent.Columns["GroupId"].Visible = false;
                dgvEmployeeContent.Columns["Name"].HeaderText = "姓名";
                dgvEmployeeContent.Columns["EmployeeId"].HeaderText = "員工編號";
                dgvEmployeeContent.Columns["PhoneNumber"].HeaderText = "電話號碼";
                dgvEmployeeContent.Columns["GroupName"].HeaderText = "群組名稱";
                dgvEmployeeContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvEmployeeContent.Columns["功能"].Width = 60;

                dgvEmployeeContent.AutoResizeColumns();
                dgvEmployeeContent.ReadOnly = true;
                dgvEmployeeContent.RowTemplate.Height = 30;
                dgvEmployeeContent.AllowUserToAddRows = false;

                tbEmployeeId.Text = string.Empty;
                tbEmployeeName.Text = string.Empty;
                cbbEmployeeGroupName.SelectedIndex = 0;
                tbEmployeePhoneNumber.Text = string.Empty;

                lblEmployeeTotleCount.Text = "共" + dtEmployee.Rows.Count + "筆資料";
            }
            else
            {
                dgvEmployeeContent.DataSource = null;
                dgvEmployeeContent.Refresh();
                dgvEmployeeContent.Columns.Clear();
                cbbEmployeeGroupName.SelectedIndex = 0;

                lblEmployeeTotleCount.Text = string.Empty;
            }
        }

        private void InitializeEmpolyeeCcbGroup()
        {
            //Group data set
            DataTable dtGroup = GroupService.SelectGroupAll();
            DataTable dtGroup2 = GroupService.SelectGroupAll();

            if (dtGroup.Rows.Count > 0)
            {
                DataRow dr = dtGroup.NewRow();
                dr["Name"] = "---請選擇群組---";
                dr["Id"] = "0";
                dtGroup.Rows.InsertAt(dr, 0);

                DataRow dr2 = dtGroup2.NewRow();
                dr2["Name"] = "---請選擇群組---";
                dr2["Id"] = "0";
                dtGroup2.Rows.InsertAt(dr2, 0);

                //搜尋的ccb
                cbbSelectEmployeeGroupName.DataSource = null;
                cbbSelectEmployeeGroupName.Items.Clear();
                cbbSelectEmployeeGroupName.DisplayMember = "Name";
                cbbSelectEmployeeGroupName.ValueMember = "Id";
                cbbSelectEmployeeGroupName.DataSource = dtGroup;

                //功能的ccb

                cbbEmployeeGroupName.DataSource = null;
                cbbEmployeeGroupName.Items.Clear();
                cbbEmployeeGroupName.DisplayMember = "Name";
                cbbEmployeeGroupName.ValueMember = "Id";
                cbbEmployeeGroupName.DataSource = dtGroup2;
            }
            else
            {
                //搜尋的ccb
                cbbSelectEmployeeGroupName.DataSource = null;
                cbbSelectEmployeeGroupName.Items.Clear();

                //功能的ccb
                cbbEmployeeGroupName.DataSource = null;
                cbbEmployeeGroupName.Items.Clear();
            }
        }

        private void InitializedgvEmpolyeeForm()
        {
            DataTable dtEmployee = EmployeeService.GetEmployee();
            if (dtEmployee.Rows.Count > 0)
            {
                dgvEmployeeContent.DataSource = null;
                dgvEmployeeContent.Refresh();
                dgvEmployeeContent.Columns.Clear();
                dgvEmployeeContent.DataSource = dtEmployee;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvEmployeeContent.Columns.Add(btncDelete);

                dgvEmployeeContent.Columns["GroupId"].Visible = false;
                dgvEmployeeContent.Columns["Name"].HeaderText = "姓名";
                dgvEmployeeContent.Columns["EmployeeId"].HeaderText = "員工編號";
                dgvEmployeeContent.Columns["PhoneNumber"].HeaderText = "電話號碼";
                dgvEmployeeContent.Columns["GroupName"].HeaderText = "群組名稱";
                dgvEmployeeContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvEmployeeContent.Columns["功能"].Width = 60;

                dgvEmployeeContent.AutoResizeColumns();
                dgvEmployeeContent.ReadOnly = true;
                dgvEmployeeContent.RowTemplate.Height = 30;
                dgvEmployeeContent.AllowUserToAddRows = false;

                tbEmployeeId.Text = string.Empty;
                tbEmployeeName.Text = string.Empty;
                cbbEmployeeGroupName.SelectedIndex = 0;
                tbEmployeePhoneNumber.Text = string.Empty;
            }
            else
            {
                dgvEmployeeContent.DataSource = null;
                dgvEmployeeContent.Refresh();
                dgvEmployeeContent.Columns.Clear();
                cbbEmployeeGroupName.SelectedIndex = 0;

            }
        }

        private void dgvEmployeeContent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            //功能刪除
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 5)
            {
                var confirmResult = MessageBox.Show("確定要刪除此筆資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string ErrorMsg = string.Empty;
                if (confirmResult == DialogResult.Yes)
                {
                    //Get data
                    string EmployeeId = dgvEmployeeContent.Rows[e.RowIndex].Cells["EmployeeId"].Value.ToString();
                    string GroupId = dgvEmployeeContent.Rows[e.RowIndex].Cells["GroupId"].Value.ToString();
                    //Delete data
                    if (EmployeeId.Trim() != null && EmployeeId.Trim() != string.Empty)
                    {
                        ErrorMsg = EmployeeService.DeleteEmployee(EmployeeId, GroupId);
                    }
                }
                else
                    return;

                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("刪除成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializedgvEmpolyeeForm();
                }
            }
        }

        private void dgvEmployeeContent_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            tbEmployeeId.Text = e.Row.Cells["EmployeeId"].Value.ToString();
            tbEmployeeName.Text = e.Row.Cells["Name"].Value.ToString();
            cbbEmployeeGroupName.SelectedValue = e.Row.Cells["GroupId"].Value.ToString();
            tbEmployeePhoneNumber.Text = e.Row.Cells["PhoneNumber"].Value.ToString();
        }

        private void btnEmployeeAdd_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string EmployeeId = tbEmployeeId.Text.Trim();
            string EmployeeName = tbEmployeeName.Text.Trim();
            string EmployeeGroupId = cbbEmployeeGroupName.SelectedValue.ToString().Trim();
            string EmployeeGroupName = cbbEmployeeGroupName.Text.Trim();
            string EmployeePhoneNumber = tbEmployeePhoneNumber.Text.Trim();

            //Check data
            //是否空白
            if (EmployeeId == string.Empty || EmployeeName == string.Empty || EmployeeGroupId == string.Empty || EmployeeGroupId == "0" || EmployeePhoneNumber == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「員工編號」、「姓名」、「群組名稱」、「電話號碼」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = EmployeeService.InsertEmployee(EmployeeId, EmployeeGroupId, EmployeeName, EmployeePhoneNumber, EmployeeGroupName);
            }


            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvEmpolyeeForm();
            }
        }

        private void btnEmployeeUpdate_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;

            //Get data
            //Get data
            string EmployeeId = tbEmployeeId.Text.Trim();
            string EmployeeName = tbEmployeeName.Text.Trim();
            string EmployeeGroupId = cbbEmployeeGroupName.SelectedValue.ToString().Trim();
            string EmployeeGroupName = cbbEmployeeGroupName.Text.Trim();
            string EmployeePhoneNumber = tbEmployeePhoneNumber.Text.Trim();

            //Check data
            //是否空白
            if (EmployeeId == string.Empty && EmployeeName == string.Empty || EmployeeGroupId == string.Empty || EmployeeGroupId == "0" || EmployeePhoneNumber == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「員工編號」、「姓名」、「群組名稱」、「電話號碼」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {

                var confirmResult = MessageBox.Show("確定要修改資料嗎?\n",
                                         "Check",
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmResult == DialogResult.Yes)
                {
                    ErrorMsg = EmployeeService.UpdateEmployee(EmployeeId, EmployeeGroupId, EmployeeName, EmployeePhoneNumber, EmployeeGroupName);
                }
                else
                    return;
            }

            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("更新成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvEmpolyeeForm();
            }
        }
        //---------------------------------------------Lable----------------------------------------------------------------
        private void btnLableSelect_Click(object sender, EventArgs e)
        {
            //Get data
            string LableName = tbLableSelectName.Text.Trim();
            string LableType = ccbLableSelectType.SelectedValue.ToString();

            DataTable dtLable = new DataTable();

            if (LableName.Trim() == string.Empty && LableType.Trim() == "0")
                dtLable = LableService.SelectLableAll();
            else
                dtLable = LableService.GetLable(LableName, LableType);

            //Set data
            if (dtLable.Rows.Count > 0)
            {
                dgvLableSelectContent.DataSource = null;
                dgvLableSelectContent.Refresh();
                dgvLableSelectContent.Columns.Clear();
                dgvLableSelectContent.DataSource = dtLable;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvLableSelectContent.Columns.Add(btncDelete);

                dgvLableSelectContent.Columns["Id"].Visible = false;
                dgvLableSelectContent.Columns["Name"].HeaderText = "姓名";
                dgvLableSelectContent.Columns["Type"].HeaderText = "種類";
                dgvLableSelectContent.Columns["ContentsSet"].HeaderText = "內容";
                dgvLableSelectContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvLableSelectContent.Columns["功能"].Width = 60;

                dgvLableSelectContent.AutoResizeColumns();
                dgvLableSelectContent.ReadOnly = true;
                dgvLableSelectContent.RowTemplate.Height = 30;
                dgvLableSelectContent.AllowUserToAddRows = false;

                tbLableId.Text = string.Empty;
                tbLableName.Text = string.Empty;
                ccbLableType.SelectedIndex = 0;
                dgvLableContent.DataSource = null;
                dgvLableContent.Refresh();
                dgvLableContent.Columns.Clear();
            }
            else
            {
                dgvLableSelectContent.DataSource = null;
                dgvLableSelectContent.Refresh();
                dgvLableSelectContent.Columns.Clear();
                ccbLableType.SelectedIndex = 0;
            }

        }

        private void InitializedgvLableForm()
        {
            DataTable dtLable = dtLable = LableService.SelectLableAll();

            //Set data
            if (dtLable.Rows.Count > 0)
            {
                dgvLableSelectContent.DataSource = null;
                dgvLableSelectContent.Refresh();
                dgvLableSelectContent.Columns.Clear();
                dgvLableSelectContent.DataSource = dtLable;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvLableSelectContent.Columns.Add(btncDelete);

                dgvLableSelectContent.Columns["Id"].Visible = false;
                dgvLableSelectContent.Columns["Name"].HeaderText = "姓名";
                dgvLableSelectContent.Columns["Type"].HeaderText = "種類";
                dgvLableSelectContent.Columns["ContentsSet"].HeaderText = "內容";
                dgvLableSelectContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvLableSelectContent.Columns["功能"].Width = 60;

                dgvLableSelectContent.AutoResizeColumns();
                dgvLableSelectContent.ReadOnly = true;
                dgvLableSelectContent.RowTemplate.Height = 30;
                dgvLableSelectContent.AllowUserToAddRows = false;

                tbLableId.Text = string.Empty;
                tbLableName.Text = string.Empty;
                ccbLableType.SelectedIndex = 0;
                dgvLableContent.DataSource = null;
                dgvLableContent.Refresh();
                dgvLableContent.Columns.Clear();
            }
            else
            {
                dgvLableSelectContent.DataSource = null;
                dgvLableSelectContent.Refresh();
                dgvLableSelectContent.Columns.Clear();
                ccbLableType.SelectedIndex = 0;
            }
        }

        private void InitializeLableCcbGroup()
        {
            DataTable dtNew = new DataTable();
            dtNew.Columns.Add(new DataColumn("Name"));
            dtNew.Columns.Add(new DataColumn("Value"));

            DataRow dr = dtNew.NewRow();
            dr["Name"] = "---請選擇種類---";
            dr["Value"] = "0";
            dtNew.Rows.InsertAt(dr, 0);

            dr = dtNew.NewRow();
            dr["Name"] = "文字列";
            dr["Value"] = "Text";
            dtNew.Rows.InsertAt(dr, 1);

            dr = dtNew.NewRow();
            dr["Name"] = "下拉選單(文字)";
            dr["Value"] = "General";
            dtNew.Rows.InsertAt(dr, 2);

            dr = dtNew.NewRow();
            dr["Name"] = "下拉選單(數字)";
            dr["Value"] = "Number";
            dtNew.Rows.InsertAt(dr, 3);

            dr = dtNew.NewRow();
            dr["Name"] = "錄音檔";
            dr["Value"] = "Record";
            dtNew.Rows.InsertAt(dr, 4);

            DataTable dtNew2 = new DataTable();
            dtNew2.Columns.Add(new DataColumn("Name"));
            dtNew2.Columns.Add(new DataColumn("Value"));

            DataRow dr2 = dtNew2.NewRow();
            dr2["Name"] = "---請選擇種類---";
            dr2["Value"] = "0";
            dtNew2.Rows.InsertAt(dr2, 0);

            dr2 = dtNew2.NewRow();
            dr2["Name"] = "文字列";
            dr2["Value"] = "Text";
            dtNew2.Rows.InsertAt(dr2, 1);

            dr2 = dtNew2.NewRow();
            dr2["Name"] = "下拉選單(文字)";
            dr2["Value"] = "General";
            dtNew2.Rows.InsertAt(dr2, 2);

            dr2 = dtNew2.NewRow();
            dr2["Name"] = "下拉選單(數字)";
            dr2["Value"] = "Number";
            dtNew2.Rows.InsertAt(dr2, 3);

            dr2 = dtNew2.NewRow();
            dr2["Name"] = "錄音檔";
            dr2["Value"] = "Record";
            dtNew2.Rows.InsertAt(dr2, 4);

            //搜尋的ccb
            ccbLableSelectType.DataSource = null;
            ccbLableSelectType.Items.Clear();
            ccbLableSelectType.DisplayMember = "Name";
            ccbLableSelectType.ValueMember = "Value";
            ccbLableSelectType.DataSource = dtNew;

            //功能的ccb
            ccbLableType.DataSource = null;
            ccbLableType.Items.Clear();
            ccbLableType.DisplayMember = "Name";
            ccbLableType.ValueMember = "Value";
            ccbLableType.DataSource = dtNew2;

        }

        private void dgvLableSelectContent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            //功能刪除
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 4)
            {
                var confirmResult = MessageBox.Show("確定要刪除此筆資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string ErrorMsg = string.Empty;
                if (confirmResult == DialogResult.Yes)
                {
                    //Get data
                    string LableId = dgvLableSelectContent.Rows[e.RowIndex].Cells["Id"].Value.ToString();
                    //Delete data
                    if (LableId.Trim() != null && LableId.Trim() != string.Empty)
                    {
                        ErrorMsg = LableService.DeleteLable(LableId);
                    }
                }
                else
                    return;

                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("刪除成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializedgvLableForm();
                }
            }
        }

        private void dgvLableSelectContent_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            tbLableId.Text = e.Row.Cells["Id"].Value.ToString();
            tbLableName.Text = e.Row.Cells["Name"].Value.ToString();
            ccbLableType.SelectedValue = e.Row.Cells["Type"].Value.ToString();

            //依照不同Type帶入
            LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), e.Row.Cells["Type"].Value.ToString());
            switch (LableEnum)
            {
                //數字
                case LableEnum.Number:
                case LableEnum.General:
                    {
                        dgvLableContent.DataSource = null;
                        dgvLableContent.Refresh();
                        dgvLableContent.Columns.Clear();
                        string Content = e.Row.Cells["ContentsSet"].Value.ToString();
                        List<string> Contents = Content.Split('｜').ToList();
                        DataTable dtNew = new DataTable();
                        dtNew.Columns.Add(new DataColumn("內容"));
                        foreach (string item in Contents)
                        {
                            dtNew.Rows.Add(item);
                        }
                        dgvLableContent.DataSource = dtNew;
                        dgvLableContent.AllowUserToAddRows = true;
                        dgvLableContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    }
                    break;
                //文字
                case LableEnum.Text:
                    {
                        dgvLableContent.DataSource = null;
                        dgvLableContent.Refresh();
                        dgvLableContent.Columns.Clear();
                        dgvLableContent.AllowUserToAddRows = false;
                    }
                    break;

                //錄音檔
                case LableEnum.Record:
                    {
                        dgvLableContent.DataSource = null;
                        dgvLableContent.Refresh();
                        dgvLableContent.Columns.Clear();
                        dgvLableContent.AllowUserToAddRows = false;
                        dgvLableContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                        //加入一行數值
                        DataTable dtNew = new DataTable();
                        dtNew.Columns.Add(new DataColumn("內容"));
                        dtNew.Rows.Add("");
                        dgvLableContent.DataSource = dtNew;
                        dgvLableContent.Columns["內容"].Visible = false;

                        //加入下拉選單
                        string Content = e.Row.Cells["ContentsSet"].Value.ToString();
                        string ErrorMsg = string.Empty;
                        DataTable dtRecord = RecordingService.SelectRecording(out ErrorMsg);
                        dtRecord.Columns.Remove("ph_data");
                        if (dtRecord.Rows.Count > 0)
                        {
                            //加入下拉選單欄位
                            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
                            combo.DataSource = dtRecord;
                            combo.DisplayMember = "ph_no";
                            combo.ValueMember = "ph_no";
                            combo.HeaderText = "錄音檔ID";
                            combo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            combo.Width = 50;
                            dgvLableContent.Columns.Insert(0, combo);
                        }

                        //下拉選單帶入數值
                        if (e.Row.Cells["ContentsSet"].Value.ToString().Trim() != string.Empty)
                        {
                            //檢查內容
                            string RecordId = e.Row.Cells["ContentsSet"].Value.ToString();
                            var NowRecord = dtRecord.AsEnumerable().Where(w => w.Field<string>("ph_no") == RecordId);
                            if (NowRecord.Count() > 0)
                            {
                                dgvLableContent.Rows[0].Cells[0].Value = RecordId;
                            }

                        }
                    }
                    break;
                default:
                    break;
            }
        }

        private void btnLableAdd_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string LableName = tbLableName.Text.Trim();
            string LableType = ccbLableType.SelectedValue.ToString().Trim();

            //Check data
            //是否空白
            if (LableName == string.Empty || LableType == "0")
            {
                ErrorMsg = ErrorMsg + "「名稱」、「種類」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = LableService.InsertLable(LableName, LableType);
            }

            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvLableForm();
            }
        }

        private void btnLableUpdate_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;

            //Get data
            string LableId = tbLableId.Text.Trim();
            string LableName = tbLableName.Text.Trim();
            string LableType = ccbLableType.SelectedValue.ToString().Trim();
            string ContentsSet = string.Empty;
            if (dgvLableContent.Rows.Count > 0)
            {
                LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), LableType);
                switch (LableEnum)
                {
                    //數字或一般下拉
                    case LableEnum.Number:
                    case LableEnum.General:
                        {
                            foreach (DataGridViewRow row in dgvLableContent.Rows)
                            {
                                if (row.Cells["內容"].Value != null)
                                {
                                    if (dgvLableContent.Rows.IndexOf(row) != dgvLableContent.Rows.Count - 2)
                                        ContentsSet = ContentsSet + row.Cells["內容"].Value.ToString().Trim().Replace("｜", "").Replace("※", "").Replace("【", "").Replace("】", "") + '｜';
                                    else
                                        ContentsSet = ContentsSet + row.Cells["內容"].Value.ToString().Trim().Replace("｜", "").Replace("※", "").Replace("【", "").Replace("】", "");
                                }
                            }
                        }
                        break;

                    //文字
                    case LableEnum.Text:
                        {

                        }
                        break;
                    //錄音檔
                    case LableEnum.Record:
                        {
                            if (dgvLableContent.Rows[0].Cells[0].Value != null)
                                ContentsSet = dgvLableContent.Rows[0].Cells[0].Value.ToString();
                            else
                                ErrorMsg = ErrorMsg + "請選擇「錄音檔ID」內容\n";
                        }
                        break;
                    default:
                        break;
                }
            }

            //Check data
            //是否空白
            if (LableId == string.Empty || LableName == string.Empty || LableType == "0")
            {
                ErrorMsg = ErrorMsg + "「編號」、「名稱」、「種類」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {

                var confirmResult = MessageBox.Show("確定要修改資料嗎?\n",
                                         "Check",
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmResult == DialogResult.Yes)
                {
                    ErrorMsg = LableService.UpdateLable(LableId, LableName, LableType, ContentsSet);
                }
                else
                    return;
            }

            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("更新成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvLableForm();
            }
        }
        //----------------------------------------------腳本---------------------------------------------------
        private void btnScriptSelect_Click(object sender, EventArgs e)
        {
            //初始化datagridview
            InitializedgvScriptForm();
        }

        private void InitializedgvScriptForm()
        {
            //Get data
            string ScriptName = tbScriptSelectName.Text.Trim();
            DataTable dtScript = new DataTable();

            if (ScriptName == string.Empty) { dtScript = ScriptService.SelectScriptAll(); }
            else { dtScript = ScriptService.GetScriptByName(ScriptName); }

            //Set data
            if (dtScript.Rows.Count > 0)
            {
                dgvScriptContent.DataSource = null;
                dgvScriptContent.Refresh();
                dgvScriptContent.Columns.Clear();
                dgvScriptContent.DataSource = dtScript;

                //新增刪除欄
                DataGridViewButtonColumn btncDelete = new DataGridViewButtonColumn();
                btncDelete.HeaderText = "功能";
                btncDelete.Name = "功能";
                btncDelete.Text = "刪除";
                btncDelete.UseColumnTextForButtonValue = true;
                dgvScriptContent.Columns.Add(btncDelete);

                dgvScriptContent.Columns["Id"].Visible = false;
                dgvScriptContent.Columns["TextBoxCount"].Visible = false;
                dgvScriptContent.Columns["LableCount"].Visible = false;
                dgvScriptContent.Columns["TextBoxContentsSet"].Visible = false;
                dgvScriptContent.Columns["LableContentSet"].Visible = false;
                dgvScriptContent.Columns["Name"].HeaderText = "腳本名稱";
                dgvScriptContent.Columns["Depiction"].HeaderText = "腳本敘述";
                dgvScriptContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvScriptContent.Columns["功能"].Width = 60;
                dgvScriptContent.AutoResizeColumns();
                dgvScriptContent.ReadOnly = true;
                dgvScriptContent.RowTemplate.Height = 30;
                dgvScriptContent.AllowUserToAddRows = false;

                //清除功能裡的欄位資料
                tbScriptId.Text = string.Empty;
                tbScriptName.Text = string.Empty;
                tbScriptDepiction.Text = string.Empty;

                dgvScriptLableContent.DataSource = null;
                dgvScriptLableContent.Refresh();
                dgvScriptLableContent.Columns.Clear();
            }
            else
            {
                dgvScriptContent.DataSource = null;
                dgvScriptContent.Refresh();
                dgvScriptContent.Columns.Clear();
            }
        }

        private void dgvScriptContent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            //功能刪除
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0 && e.ColumnIndex == 7)
            {
                var confirmResult = MessageBox.Show("確定要刪除此筆資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string ErrorMsg = string.Empty;
                if (confirmResult == DialogResult.Yes)
                {
                    //Get data
                    string ScriptId = dgvScriptContent.Rows[e.RowIndex].Cells["Id"].Value.ToString();
                    //Delete data
                    if (ScriptId.Trim() != null && ScriptId.Trim() != string.Empty)
                    {
                        ErrorMsg = ScriptService.DeleteScript(ScriptId);
                    }
                }
                else
                    return;

                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("刪除成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializedgvScriptForm();
                    InitializeControl();
                }
            }
        }

        private void dgvScriptContent_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            tbScriptId.Text = e.Row.Cells["Id"].Value.ToString();
            tbScriptName.Text = e.Row.Cells["Name"].Value.ToString();
            tbScriptDepiction.Text = e.Row.Cells["Depiction"].Value.ToString();

            dgvScriptLableContent.DataSource = null;
            dgvScriptLableContent.Refresh();
            dgvScriptLableContent.Columns.Clear();
            string LableContentSet = e.Row.Cells["LableContentSet"].Value.ToString();
            List<string> ContentsId = LableContentSet.Split('｜').ToList();
            if (ContentsId[0].Trim() != string.Empty)
            {
                //排序
                ContentsId = ContentsId.OrderBy(o => o.Split('※')[1]).ToList();
            }
            DataTable dtNew = new DataTable();
            dtNew.Columns.Add(new DataColumn("內容"));
            dtNew.Columns.Add(new DataColumn("排序"));

            DataTable dtLable = LableService.SelectLableAll();
            //加入下拉選單欄位
            DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
            combo.DataSource = dtLable;
            combo.DisplayMember = "Name";
            combo.ValueMember = "Id";
            combo.HeaderText = "名稱";
            combo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            combo.Width = 50;
            dgvScriptLableContent.Columns.Insert(0, combo);

            if (ContentsId[0].Trim() != string.Empty)
                foreach (string id in ContentsId)
                {
                    var dtSelect = dtLable.AsEnumerable().Where(w => w.Field<string>("id") == id.Split('※')[0]);

                    if (dtSelect.Count() > 0)
                    {
                        DataTable tbLable = dtSelect.CopyToDataTable();
                        DataRow NewRow;
                        NewRow = dtNew.NewRow();

                        NewRow["內容"] = tbLable.Rows[0]["ContentsSet"];
                        NewRow["排序"] = id.Split('※')[1];
                        dtNew.Rows.Add(NewRow);
                    }
                    else
                    {
                        DataRow NewRow;
                        NewRow = dtNew.NewRow();

                        NewRow["內容"] = "該原本標籤已經被刪除";
                        NewRow["排序"] = "該原本標籤已經被刪除";
                        dtNew.Rows.Add(NewRow);
                    }
                }

            dgvScriptLableContent.DataSource = dtNew;
            dgvScriptLableContent.Columns["內容"].ReadOnly = true;
            dgvScriptLableContent.Columns[0].Width = 120;
            dgvScriptLableContent.Columns["內容"].Width = 681;
            dgvScriptLableContent.Columns["排序"].Width = 38;
            //dgvScriptLableContent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _isInitializeScriptLableStatus = false;
            //取得Text資料
            string TextContentSet = e.Row.Cells["TextBoxContentsSet"].Value.ToString();
            List<string> TextList = TextContentSet.Split('｜').ToList();
            //更新ccb狀態
            foreach (DataGridViewRow row in dgvScriptLableContent.Rows)
            {
                int index = dgvScriptLableContent.Rows.IndexOf(row);
                if (index != dgvScriptLableContent.Rows.Count - 1)
                {
                    string LableId = ContentsId[index].Split('※')[0];

                    var dtSelect = dtLable.AsEnumerable().Where(w => w.Field<string>("id") == LableId);
                    //如果有找把值帶入，如果沒找到代表此標籤已被刪除
                    if (dtSelect.Count() > 0)
                    {
                        row.Cells[0].Value = LableId;
                    }

                    //依照不同的Lable Type變更Value與設定值
                    if (dtSelect.Count() > 0)
                    {
                        DataTable tbLable = dtSelect.CopyToDataTable();
                        LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), tbLable.Rows[0].Field<string>("Type"));
                        switch (LableEnum)
                        {
                            //文字
                            case LableEnum.Text:
                                {
                                    if (TextList[0].Trim() != string.Empty)
                                    {
                                        //排序
                                        TextList = TextList.OrderBy(o => o.Split('※')[1]).ToList();
                                    }
                                    //內容照順序往下塞，塞完刪除
                                    var NowContent = TextList.Where(w => w.Split('※')[1] == row.Cells["排序"].Value.ToString()).ToList();
                                    if (NowContent.Count() > 0)
                                    {
                                        row.Cells["內容"].Value = NowContent[0].Split('※')[0];
                                    }
                                    row.Cells["內容"].ReadOnly = false;
                                }
                                break;
                            //錄音檔
                            case LableEnum.Record:
                                {
                                    string RecordId = tbLable.Rows[0]["ContentsSet"].ToString();
                                    DataTable dtRecord = RecordingService.SelectRecording(RecordId);
                                    if (dtRecord.Rows.Count > 0)
                                    {
                                        string Content = dtRecord.Rows[0]["ph_data"].ToString();
                                        string ph_data2 = dtRecord.Rows[0]["ph_data2"].ToString();
                                        if (Content == null || Content.Trim() == string.Empty)
                                        {
                                            if (ph_data2 == null) { ph_data2 = string.Empty; }
                                            row.Cells["內容"].Value = ph_data2;
                                        }
                                        else
                                        {
                                            //編碼
                                            byte[] unknow = Encoding.GetEncoding(28591).GetBytes(Content);
                                            string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                                            row.Cells["內容"].Value = Big5;
                                        }
                                    }
                                    else
                                        row.Cells["內容"].Value = "此標籤目前沒有對應的錄音檔";
                                    row.Cells["內容"].ReadOnly = true;
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            _isInitializeScriptLableStatus = true;

            //表頭不排序
            foreach (DataGridViewColumn col in dgvScriptLableContent.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Add the events to listen for
            dgvScriptLableContent.CellValueChanged += new DataGridViewCellEventHandler(dgvScriptLableContent_CellValueChanged);
            dgvScriptLableContent.CurrentCellDirtyStateChanged += new EventHandler(dgvScriptLableContent_CurrentCellDirtyStateChanged);

            //dgvScriptLableContent.AllowUserToAddRows = false;
            //dgvScriptLableContent.ReadOnly = true;
        }

        private void btnScriptAdd_Click(object sender, EventArgs e)
        {
            string ErrorMsg = string.Empty;
            //Get data
            string ScriptName = tbScriptName.Text.Trim();
            string ScriptDepiction = tbScriptDepiction.Text.Trim();

            //Check data
            //是否空白
            if (ScriptName == string.Empty || ScriptDepiction == string.Empty)
            {
                ErrorMsg = ErrorMsg + "「腳本名稱」、「腳本敘述」不能空白\n";
            }

            //Set data
            if (ErrorMsg == string.Empty)
            {
                ErrorMsg = ScriptService.InsertScript(ScriptName, ScriptDepiction);
            }

            //Show message
            if (ErrorMsg != string.Empty)
            {
                MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //初始化datagridview
                InitializedgvScriptForm();
                InitializeControl();
            }
        }

        private void btnScriptUpdate_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("確定要修改此筆資料嗎?\n",
                                     "Check",
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult == DialogResult.Yes)
            {
                int TempTextBoxCount = 0;
                int TempNotTextBoxCount = 0;
                int TempLableCount = 0;
                string TempTextBoxContentsSet = string.Empty;
                string TempLableContentSet = string.Empty;
                bool IsSortOk = false;
                string ErrorMsg = string.Empty;
                //Get data

                //DataGridView
                foreach (DataGridViewRow item in dgvScriptLableContent.Rows)
                {
                    int index = dgvScriptLableContent.Rows.IndexOf(item);

                    //濾掉沒有新增的新增項
                    if (item.Cells[0].Value != null)
                    {
                        string LableId = item.Cells[0].Value.ToString();
                        DataTable dtLable = LableService.GetLable(LableId);
                        if (dtLable.Rows.Count > 0)
                        {
                            LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), dtLable.Rows[0].Field<string>("Type"));
                            //一般Text標籤
                            if (LableEnum == LableEnum.Text)
                            {
                                //判斷是否為第一個TextBoxContent
                                if (TempTextBoxCount != 0)
                                {
                                    TempTextBoxContentsSet = TempTextBoxContentsSet + '｜';
                                }
                                TempTextBoxContentsSet = TempTextBoxContentsSet + item.Cells["內容"].Value.ToString().Trim().Replace("｜", "").Replace("※", "").Replace("【", "").Replace("】", "") + "※" + item.Cells["排序"].Value.ToString();

                                TempTextBoxCount++;
                            }

                            //LableContentSet
                            //判斷是否為第一個LableContent
                            if (TempNotTextBoxCount != 0)
                            {
                                TempLableContentSet = TempLableContentSet + '｜';
                            }
                            TempLableContentSet = TempLableContentSet + item.Cells[0].Value.ToString() + "※" + item.Cells["排序"].Value.ToString();
                            TempNotTextBoxCount++;

                            //LableCount
                            TempLableCount++;

                            //判斷排序資料是否正確
                            int resiut = 0;
                            if (int.TryParse(item.Cells["排序"].Value.ToString(), out resiut))
                                IsSortOk = true;
                            else
                                IsSortOk = false;
                        }
                    }
                }

                //ID
                string ScriptId = tbScriptId.Text;
                //Name
                string ScriptName = tbScriptName.Text;
                //TextBoxCount
                int TextBoxCount = TempTextBoxCount;
                //LableCount
                int LableCount = TempLableCount;
                //TextBoxContentsSet
                string TextBoxContentsSet = TempTextBoxContentsSet;
                //LableContentSet
                string LableContentSet = TempLableContentSet;
                //Depiction
                string Depiction = tbScriptDepiction.Text;

                //Check data
                //是否空白
                if (ScriptId.Trim() == string.Empty || ScriptName.Trim() == string.Empty || Depiction.Trim() == string.Empty)
                {
                    ErrorMsg = ErrorMsg + "「腳本編號」、「腳本名稱」、「腳本敘述」不能空白\n";
                }
                //檢查排序
                if (IsSortOk == false)
                {
                    ErrorMsg = ErrorMsg + "「排序」不可空白並且必須為數字\n";
                }

                //Set data
                if (ErrorMsg == string.Empty)
                {
                    ErrorMsg = ScriptService.UpdateScript(ScriptId, ScriptName, TempTextBoxCount, LableCount, TextBoxContentsSet, LableContentSet, Depiction);
                }

                //Show message
                if (ErrorMsg != string.Empty)
                {
                    MessageBox.Show(ErrorMsg, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("新增成功", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //初始化datagridview
                    InitializedgvScriptForm();
                    InitializeControl();
                }
            }
        }


        // This event handler manually raises the CellValueChanged event 
        // by calling the CommitEdit method. 
        void dgvScriptLableContent_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvScriptLableContent.IsCurrentCellDirty)
            {
                // This fires the cell value changed handler below
                dgvScriptLableContent.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgvScriptLableContent_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // My combobox column is the second one so I hard coded a 0, flavor to taste
            DataGridViewComboBoxCell cb = (DataGridViewComboBoxCell)dgvScriptLableContent.Rows[e.RowIndex].Cells[0];

            if (cb.Value != null && _isInitializeScriptLableStatus && e.ColumnIndex == 0)
            {
                //Get data
                string LableId = dgvScriptLableContent.Rows[e.RowIndex].Cells[0].Value.ToString();
                DataTable dtLable = LableService.GetLable(LableId);
                if (dtLable.Rows.Count > 0)
                {
                    LableEnum LableEnum = (LableEnum)Enum.Parse(typeof(LableEnum), dtLable.Rows[0].Field<string>("Type"));
                    switch (LableEnum)
                    {
                        //數字或一般下拉
                        case LableEnum.Number:
                        case LableEnum.General:
                            {
                                dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].Value = dtLable.Rows[0]["ContentsSet"];
                                dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].ReadOnly = true;
                            }
                            break;

                        //文字
                        case LableEnum.Text:
                            {
                                dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].ReadOnly = false;
                            }
                            break;
                        //錄音檔
                        case LableEnum.Record:
                            {
                                string RecordId = dtLable.Rows[0]["ContentsSet"].ToString();
                                DataTable dtRecord = RecordingService.SelectRecording(RecordId);
                                if (dtRecord.Rows.Count > 0)
                                {
                                    string Content = dtRecord.Rows[0]["ph_data"].ToString();
                                    string ph_data2 = dtRecord.Rows[0]["ph_data2"].ToString();
                                    if (Content == null || Content.Trim() == string.Empty)
                                    {
                                        if (ph_data2 == null) { ph_data2 = string.Empty; }
                                        dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].Value = ph_data2;
                                    }
                                    else
                                    {
                                        byte[] unknow = Encoding.GetEncoding(28591).GetBytes(Content);
                                        string Big5 = Encoding.GetEncoding(950).GetString(unknow);
                                        dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].Value = Big5;
                                    }
                                }
                                else
                                {
                                    dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].Value = "此標籤目前沒有對應的錄音檔";
                                }
                                dgvScriptLableContent.Rows[e.RowIndex].Cells["內容"].ReadOnly = true;
                            }
                            break;
                        default:
                            break;
                    }
                }

                //do stuff
                dgvScriptLableContent.Invalidate();
            }
        }

        //----------------------------------------------31處無人發報---------------------------------------------------
        private void InitializedgvdgvAutoSendForm()
        {
            //Get data
            DataTable dtAutoSent = AutoSentService.SelectAutoSentAll();

            //Set data
            if (dtAutoSent.Rows.Count > 0)
            {
                dgvAutoSend.DataSource = null;
                dgvAutoSend.Refresh();
                dgvAutoSend.Columns.Clear();
                dgvAutoSend.DataSource = dtAutoSent;

                //新增圖片狀態欄
                Bitmap img;
                string path = System.Windows.Forms.Application.StartupPath;
                img = new Bitmap(@path + @"\Resources\Button-Blank-Gray-icon.png");
                DataGridViewImageColumn igStatus = new DataGridViewImageColumn();
                igStatus.HeaderText = "狀態";
                igStatus.Name = "Status";
                igStatus.Image = img;
                igStatus.ImageLayout = DataGridViewImageCellLayout.Zoom;
                dgvAutoSend.Columns.Insert(2, igStatus);

                //新增Tag狀態欄
                DataGridViewTextBoxColumn tbTagValue = new DataGridViewTextBoxColumn();
                tbTagValue.HeaderText = "Tag數值";
                tbTagValue.Name = "Tag";
                dgvAutoSend.Columns.Insert(4, tbTagValue);
                dgvAutoSend.Columns["Tag"].SortMode = DataGridViewColumnSortMode.NotSortable;

                dgvAutoSend.Columns["Id"].Visible = false;
                dgvAutoSend.Columns["AutoGroupId"].Visible = false;
                dgvAutoSend.Columns["Sort"].Visible = false;
                dgvAutoSend.Columns["TagName"].Visible = false;
                dgvAutoSend.Columns["IsSent"].Visible = false;
                dgvAutoSend.Columns["Name"].HeaderText = "單位";
                dgvAutoSend.Columns["Name"].SortMode = DataGridViewColumnSortMode.NotSortable;
                //dgvAutoSend.Columns["AutoGroupId"].HeaderText = "發送群組";
                dgvAutoSend.Columns["LastSentTime"].HeaderText = "最後發送時間";
                dgvAutoSend.Columns["LastSentTime"].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgvAutoSend.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvAutoSend.Columns["Status"].Width = 60;
                dgvAutoSend.Columns["Tag"].Width = 100;
                //dgvAutoSend.Columns["LastSentTime"].Width = 100;
                //dgvAutoSend.AutoResizeColumns();
                dgvAutoSend.ReadOnly = true;
                //dgvAutoSend.RowTemplate.Height = 30;
                dgvAutoSend.AllowUserToAddRows = false;

                tbGroupId.Text = string.Empty;
                tbGroupName.Text = string.Empty;
            }
            else
            {
                dgvAutoSend.DataSource = null;
                dgvAutoSend.Refresh();
                dgvAutoSend.Columns.Clear();
            }
        }

        private void dgvAutoSend_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            // For any other operation except, StateChanged, do nothing
            if (e.StateChanged != DataGridViewElementStates.Selected) return;

            string AutoSentId = e.Row.Cells["Id"].Value.ToString();

            //Get data
            DataTable dtAutoSentHistory = AutoSentService.SelectAutoSentHistory(AutoSentId);
            //Set data
            if (dtAutoSentHistory.Rows.Count > 0)
            {
                dgvAutoSentHistory.DataSource = null;
                dgvAutoSentHistory.Refresh();
                dgvAutoSentHistory.Columns.Clear();
                dgvAutoSentHistory.DataSource = dtAutoSentHistory;

                dgvAutoSentHistory.Columns["Id"].Visible = false;
                dgvAutoSentHistory.Columns["AutoSentId"].Visible = false;
                dgvAutoSentHistory.Columns["AutoSentName"].Visible = false;
                dgvAutoSentHistory.Columns["AutoSentTime"].HeaderText = "發送時間";
                dgvAutoSentHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            {
                dgvAutoSentHistory.DataSource = null;
                dgvAutoSentHistory.Refresh();
                dgvAutoSentHistory.Columns.Clear();
            }
        }

        private async void AutoSentStatusPicture()
        {
            //data declare
            string NJCompoletPeerAddress = ConfigurationManager.AppSettings["NJCompoletPeerAddress"];
            int NJCompoletLocalPort = int.Parse(ConfigurationManager.AppSettings["NJCompoletLocalPort"]);
            string NJCompoletTagName = ConfigurationManager.AppSettings["NJCompoletTagName"];

            Bitmap imgGreen;
            Bitmap imgBlue;
            Bitmap imgYellow;
            Bitmap imgRed;
            string path = System.Windows.Forms.Application.StartupPath;
            imgGreen = new Bitmap(@path + @"\Resources\Button-Blank-Green-icon.png");
            imgBlue = new Bitmap(@path + @"\Resources\Button-Blank-Blue-icon.png");
            imgYellow = new Bitmap(@path + @"\Resources\Button-Blank-Yellow-icon.png");
            imgRed = new Bitmap(@path + @"\Resources\Button-Blank-Red-icon.png");
            while (_blAutoSentStatusPicture == true)
            {
                //Omrom NJ PLC Sata Get
                int colorNum = 0;
                bool[] VariableArray = new bool[0];
                try
                {
                    NJCompolet NJCompolet = new NJCompolet();
                    NJCompolet.UseRoutePath = false;
                    NJCompolet.PeerAddress = NJCompoletPeerAddress;
                    NJCompolet.LocalPort = NJCompoletLocalPort;
                    NJCompolet.Active = true;
                    VariableArray = (bool[])NJCompolet.ReadVariable(NJCompoletTagName);
                    NJCompolet.Active = false;
                    NJCompolet.Dispose();
                }
                catch (Exception ex)
                {
                    colorNum = 4;
                }
                //
                DataTable dtAutoSent = AutoSentService.SelectAutoSentAll();
                //顯示每站圖片狀態
                for (int i = 0; i < dgvAutoSend.Rows.Count; i++)
                {
                    DateTime LastSentTime = new DateTime();
                    //date check
                    if (colorNum != 4)
                    {
                        //db data get
                        //string GridViewAutoSentId = (string)dgvAutoSend.Rows[i].Cells["Id"].Value;
                        int GridViewIsSent = dtAutoSent.Rows[i].Field<int>("IsSent");
                        try
                        {
                            LastSentTime = dtAutoSent.Rows[i].Field<DateTime>("LastSentTime");
                        }
                        catch (Exception ex) { }

                        //發送訊息
                        //確認目前此PLC Tag被觸發並且尚未發送過訊息
                        if (VariableArray[i] == true && GridViewIsSent == 0)
                        {
                            //清空狀態視窗
                            ClearTextBox("", tbAutoSentMsgStatus);
                            updateTextBox("-----------------------Start-----------------------" + Environment.NewLine, tbAutoSentMsgStatus);
                            //Get AutoGroup telephone numbers
                            DataTable dtTelephoneNumber = new DataTable();
                            string AutoGroupId = (string)dgvAutoSend.Rows[i].Cells["AutoGroupId"].Value;
                            string Name = (string)dgvAutoSend.Rows[i].Cells["Name"].Value;
                            dtTelephoneNumber = AutoSentService.GetAutoEmployeeForAutoGroupId(AutoGroupId);
                            if (dtTelephoneNumber.Rows.Count > 0)
                            {
                                DateTime SentTime = DateTime.Now;
                                //發送簡訊
                                updateTextBox("開始發送SMS簡訊訊息..." + Environment.NewLine, tbAutoSentMsgStatus);
                                DoSentSutoSMS(dtTelephoneNumber, Name);
                                updateTextBox("發送SMS簡訊訊息結束..." + Environment.NewLine, tbAutoSentMsgStatus);

                                //發送TTS
                                updateTextBox("開始發送TTS語音訊息..." + Environment.NewLine, tbAutoSentMsgStatus);
                                //發送給每個使用者
                                foreach (DataRow item in dtTelephoneNumber.Rows)
                                {
                                    string AutoMsg = string.Empty;
                                    string IsTest = ConfigurationManager.AppSettings["IsTest"];
                                    if (IsTest == "0")
                                    {
                                        AutoMsg = Name + "發生火災，請相關人員回廠支援。";
                                    }
                                    else
                                    {
                                        AutoMsg = "測試 測試 " + Name + "發生火災，請相關人員回廠支援。";
                                    }

                                    //組合Url
                                    /**/
                                    string CalloutphpUrl = GetCalloutphpUrl(item.Field<string>("PhoneNumber"), item.Field<string>("Name"), AutoMsg).ToString();
                                    //Get WebService發送訊息
                                    _textboxAuto = tbSentMsgStatus;
                                    GetAutoRequest(CalloutphpUrl, item.Field<string>("PhoneNumber"));
                                }
                                updateTextBox("發送TTS語音訊息結束..." + Environment.NewLine, tbAutoSentMsgStatus);

                                //db寫入紀錄
                                string DbUpdateErrorMsg = string.Empty;
                                string AutoSendId = (string)dgvAutoSend.Rows[i].Cells["Id"].Value;
                                DbUpdateErrorMsg = DbUpdateErrorMsg + AutoSentService.UpdateAutoSend(AutoSendId, 1, SentTime);
                                DbUpdateErrorMsg = DbUpdateErrorMsg + AutoSentService.InserAutoSentHistory(AutoSendId, Name, SentTime);
                                //db紀錄寫入訊息回饋
                                if (DbUpdateErrorMsg == string.Empty)
                                {
                                    GridViewIsSent = 1;
                                    updateTextBox(Name + "IsSent寫入成功" + Environment.NewLine, tbAutoSentMsgStatus);
                                    LastSentTime = SentTime;
                                }
                                else
                                {
                                    updateTextBox(Name + "IsSent寫入失敗：" + DbUpdateErrorMsg + Environment.NewLine, tbAutoSentMsgStatus);
                                }
                            }
                            else
                            {
                                updateTextBox("無人發報錯誤：發報的單位內沒有任何人員" + Environment.NewLine, tbAutoSentMsgStatus);
                            }
                        }

                        //復歸 Db IsSent欄位
                        if (VariableArray[i] == false && GridViewIsSent == 1)
                        {
                            string DbUpdateErrorMsg = string.Empty;
                            string Id = (string)dgvAutoSend.Rows[i].Cells["Id"].Value;
                            string Name = (string)dgvAutoSend.Rows[i].Cells["Name"].Value;
                            DbUpdateErrorMsg = AutoSentService.UpdateAutoSend(Id, 0, new DateTime());
                            //db紀錄寫入訊息回饋
                            if (DbUpdateErrorMsg == string.Empty)
                            {
                                updateTextBox(Name + "復歸成功。" + Environment.NewLine, tbAutoSentMsgStatus);
                            }
                            else
                            {
                                updateTextBox(Name + "復歸失敗：" + DbUpdateErrorMsg + Environment.NewLine, tbAutoSentMsgStatus);
                            }
                        }

                        //狀態燈號確認
                        //Green
                        if (VariableArray[i] == false && ((DateTime.Now - LastSentTime).TotalHours >= 24 || LastSentTime == new DateTime()))
                        {
                            colorNum = 1;
                        }
                        //Blue
                        if (VariableArray[i] == true && GridViewIsSent == 1)
                        {
                            colorNum = 2;
                            //如果觸發警報後消防沒有復歸，過24小時後會再發送一次警報
                            if ((DateTime.Now - LastSentTime).TotalHours >= 24 || LastSentTime == new DateTime())
                            {
                                //復歸 Db IsSent欄位
                                if (VariableArray[i] == true && GridViewIsSent == 1)
                                {
                                    string DbUpdateErrorMsg = string.Empty;
                                    string Id = (string)dgvAutoSend.Rows[i].Cells["Id"].Value;
                                    string Name = (string)dgvAutoSend.Rows[i].Cells["Name"].Value;
                                    DbUpdateErrorMsg = AutoSentService.UpdateAutoSend(Id, 0, new DateTime());
                                    //db紀錄寫入訊息回饋
                                    if (DbUpdateErrorMsg == string.Empty)
                                    {
                                        updateTextBox(Name + "復歸成功。" + Environment.NewLine, tbAutoSentMsgStatus);
                                    }
                                    else
                                    {
                                        updateTextBox(Name + "復歸失敗：" + DbUpdateErrorMsg + Environment.NewLine, tbAutoSentMsgStatus);
                                    }
                                }
                            }
                        }
                        //Yellow
                        if (VariableArray[i] == false && ((DateTime.Now - LastSentTime).TotalHours < 24 && LastSentTime != new DateTime()))
                        {
                            colorNum = 3;
                        }
                    }

                    //data set
                    //Status
                    switch (colorNum)
                    {
                        case 1:
                            ((DataGridViewImageCell)dgvAutoSend.Rows[i].Cells["Status"]).Value = imgGreen;
                            break;
                        case 2:
                            ((DataGridViewImageCell)dgvAutoSend.Rows[i].Cells["Status"]).Value = imgBlue;
                            break;
                        case 3:
                            ((DataGridViewImageCell)dgvAutoSend.Rows[i].Cells["Status"]).Value = imgYellow;
                            break;
                        case 4:
                            ((DataGridViewImageCell)dgvAutoSend.Rows[i].Cells["Status"]).Value = imgRed;
                            break;
                        default:
                            break;
                    }
                    try
                    {
                        //TagValue
                        dgvAutoSend.Rows[i].Cells["Tag"].Value = VariableArray[i].ToString();
                    }
                    catch (Exception ex)
                    {
                        //表示目前讀取PLC內容失敗
                        //dgvAutoSend.Rows[i].Cells["Tag"].Value = "X";
                    }

                    //((DataGridViewTextBoxColumn)dgvAutoSend.Rows[i].Cells["Tag"]). = VariableArray[i].ToString();
                    //LastSentTime
                    if (LastSentTime != new DateTime())
                    {
                        dgvAutoSend.Rows[i].Cells["LastSentTime"].Value = LastSentTime.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                }
                Thread.Sleep(5 * 1000);
            }
        }

        private void btnAutoEmployeeFileImport_Click(object sender, EventArgs e)
        {
            //Select CSV File
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "csv files (*.*)|*.csv";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                AsyncTemplate.DoWorkAsync(
               () =>
               {
                   doImportAutoSentCsv(dialog.FileName);
               },
               () =>
               {
                   //MessageBox.Show("Success, Result is " + result.ToString());
                   MessageBox.Show("匯入完成", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

               },
               (exception) =>
               {
                   //MessageBox.Show(exception.Message);
                   MessageBox.Show("失敗原因；" + exception.Message, "匯出失敗", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   //error handling
               });
            }
        }

        private void doImportAutoSentCsv(string fileName)
        {
            string CrvPath = fileName;
            //Get data
            DataTable CsvDb = GetCsvDb(CrvPath);
            if (CsvDb.Rows.Count > 0)
            {
                //Set data
                //List<string> NewGroupList = CsvDb.AsEnumerable().Select(r => r.Field<string>(1)).Distinct().ToList();
                List<string> NewGroupList = CsvDb.AsEnumerable().Select(r => r.Field<string>(0)).Distinct().ToList();

                //Insert data
                AutoSentService.DeleteAutoGroup();
                //AutoGroup Insert
                string ErrorMsg = AutoSentService.MultipleInsertAutoGroup((List<string>)NewGroupList);
                if (ErrorMsg == string.Empty)
                {
                    DataTable dtAutoGroup = AutoSentService.SelectAutoGroupAll();
                    AutoSentService.DeleteAutoEmployee();
                    //AutoEmployee Insert
                    int csvCount = CsvDb.Rows.Count;
                    for (int i = 0; i < (csvCount / 400) + 1; i++)
                    {
                        DataTable rows = CsvDb.AsEnumerable().Skip(i * 400).Take(400).CopyToDataTable();
                        AutoSentService.MultipleInsertAutoEmployee(rows, dtAutoGroup);
                    }
                    //AutoSend Update
                    DataTable dtAutoSend = AutoSentService.SelectAutoSentAll();
                    if (dtAutoGroup.Rows.Count > 0)
                    {
                        ErrorMsg = AutoSentService.MultipleUpdateAutoSend(dtAutoGroup, dtAutoSend);
                    }
                }
            }
        }

        private void btnAutoExport_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //Get data
                string Path = dialog.SelectedPath;
                DataTable dtAutoEmployee = AutoSentService.GetAutoEmployee();
                //dtEmployee.Columns["GroupName"].SetOrdinal(1);
                dtAutoEmployee.Columns["GroupName"].SetOrdinal(0);
                dtAutoEmployee.Columns.Remove("GroupId");
                dtAutoEmployee.Columns["EmployeeId"].ColumnName = "員工編號";
                dtAutoEmployee.Columns["GroupName"].ColumnName = "發送群組";
                dtAutoEmployee.Columns["Name"].ColumnName = "姓名";
                dtAutoEmployee.Columns["PhoneNumber"].ColumnName = "電話號碼";

                //Set csv
                if (dtAutoEmployee.Rows.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();

                    IEnumerable<string> columnNames = dtAutoEmployee.Columns.Cast<DataColumn>().
                                                      Select(column => column.ColumnName);
                    sb.AppendLine(string.Join(",", columnNames));

                    foreach (DataRow row in dtAutoEmployee.Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }

                    File.WriteAllText(Path + "\\37處無人發報群組聯絡人匯出" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv", sb.ToString(), System.Text.Encoding.UTF8);

                    MessageBox.Show("匯出完成", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    MessageBox.Show("目前群組聯絡人無資料", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
        //SentAutoSMS
        private void DoSentSutoSMS(DataTable dtTelephoneNumber, string Name)
        {
            //連線
            string ServerIp = string.Empty;
            string ServerPort = string.Empty;
            string UserID = string.Empty;
            string Passwd = string.Empty;
            string IsTest = string.Empty;
            int ret_code;
            string ret_description = string.Empty;
            string Message = string.Empty;

            //Get Appconfig data
            //Set SMS Login data
            ServerIp = ConfigurationManager.AppSettings["ServerIp"];
            ServerPort = ConfigurationManager.AppSettings["ServerPort"];
            UserID = ConfigurationManager.AppSettings["UserID"];
            Passwd = ConfigurationManager.AppSettings["Passwd"];
            IsTest = ConfigurationManager.AppSettings["IsTest"];

            //Get註冊後的HiAir.dll，並且使用動態配置的方式宣告物件使用
            dynamic dymSMS = Activator.CreateInstance(Type.GetTypeFromProgID("HiAir.HiNetSMS"));
            //連線中華電信SMS Server
            ret_code = dymSMS.StartCon(ServerIp, ServerPort, UserID, Passwd);
            //Set data
            if (IsTest == "0")
                Message = Name + "發生火災，請相關人員回廠支援。";
            else
                Message = "測試 測試 " + Name + "發生火災，請相關人員回廠支援。";

            //表示成功連上中華電信SMS Server
            if (ret_code == 0)
            {
                //updateTextBox("中華電信SMS Server連線成功!!" + Environment.NewLine, tbAutoSentMsgStatus);

                //發送給每個使用者
                foreach (DataRow item in dtTelephoneNumber.Rows)
                {
                    string Tel = string.Empty;

                    //Get data
                    Tel = item.Field<string>("PhoneNumber");

                    ///*測試用*/
                    //Tel = "";
                    //發送
                    ret_code = dymSMS.SendMsg(Tel, Message.ToString());
                    if (ret_code == 0)
                    {
                        updateTextBox("發送簡訊至" + Tel + "成功!!" + Environment.NewLine, tbAutoSentMsgStatus);
                    }
                    else
                    {
                        try
                        {
                            ret_description = dymSMS.QueryMsg();
                            updateTextBox("發送簡訊至" + Tel + "失敗!!" + Environment.NewLine, tbAutoSentMsgStatus);
                            updateTextBox("失敗原因：" + Environment.NewLine + "", tbAutoSentMsgStatus);
                            updateTextBox(ret_description + Environment.NewLine, tbAutoSentMsgStatus);
                        }
                        catch (Exception ex)
                        {
                            updateTextBox("發送簡訊至" + Tel + "失敗!!" + Environment.NewLine, tbAutoSentMsgStatus);
                            updateTextBox("失敗原因：" + Environment.NewLine + "", tbAutoSentMsgStatus);
                            updateTextBox(ex.Message + Environment.NewLine, tbAutoSentMsgStatus);
                        }

                        //updateTextBox("發送簡訊至" + Tel + "失敗!!" + Environment.NewLine, tbSentMsgStatus);
                        //updateTextBox("失敗原因：" + Environment.NewLine + "", tbSentMsgStatus);
                        //updateTextBox(ret_description + Environment.NewLine, tbSentMsgStatus);
                    }
                }
                //結束時關閉連線
                dymSMS.EndCon();
            }
            else
            {
                ret_description = dymSMS.Get_Message();
                updateTextBox("中華電信SMS Server連線失敗!!" + Environment.NewLine, tbAutoSentMsgStatus);
                updateTextBox("失敗原因：" + Environment.NewLine + "", tbAutoSentMsgStatus);
                updateTextBox(ret_description + Environment.NewLine, tbAutoSentMsgStatus);
            }
        }

        //用Get的方式Call Web Service
        static async void GetAutoRequest(string url, string TelephoneNumber)
        {
            using (HttpClient Client = new HttpClient())
            {
                try
                {
                    using (HttpResponseMessage response = await Client.GetAsync(url))
                    {
                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();
                        if (responseBody == "DATAOK")
                        {
                            Print(_textboxAuto, "發送TTS語音訊息至" + TelephoneNumber + "成功!!" + Environment.NewLine);
                        }
                        if (responseBody == "DATAERR")
                        {
                            Print(_textboxAuto, "發送TTS語音訊息至" + TelephoneNumber + "失敗!!" + Environment.NewLine);
                        }
                    }
                }
                catch (Exception ex)
                {
                    string ErrorMsg = ex.Message;
                    Print(_textboxAuto, "發送TTS語音訊息至" + TelephoneNumber + "失敗!!" + Environment.NewLine);
                    Print(_textboxAuto, "失敗原因：" + ErrorMsg + Environment.NewLine);
                }

            }
        }

        //群組點及觸發事件
        private void dgvGroup_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                string nowClickGroup = dgvGroup.Rows[e.RowIndex].Cells[2].Value.ToString();
                string AllGroup = ConfigurationManager.AppSettings["AllGroup"];
                if (nowClickGroup.Contains(AllGroup) == false)
                {
                    return;
                }

                var senderGrid = (DataGridView)sender; 
                senderGrid.EndEdit();
                DataGridViewCheckBoxCell chkchecking = dgvGroup.Rows[e.RowIndex].Cells[0] as DataGridViewCheckBoxCell;
                if (Convert.ToBoolean(chkchecking.Value) == true)
                {
                    cbSentTTS.Checked = false;
                }
            }
        }
    }
}
