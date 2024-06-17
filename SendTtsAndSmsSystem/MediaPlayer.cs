using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WMPLib;

namespace SendTtsAndSmsSystem
{
    public partial class MediaPlayer : Form
    {
        public MediaPlayer()
        {
            InitializeComponent();
        }

        private void MediaPlayer_Load(object sender, EventArgs e)
        {
            //axWindowsMediaPlayer1.settings.autoStart = true;
            axWindowsMediaPlayer1.URL = this.FilePath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void axWindowsMediaPlayer1_PlayStateChange(object sender, AxWMPLib._WMPOCXEvents_PlayStateChangeEvent e)
        {
            if (axWindowsMediaPlayer1.status == "就緒")
            {
                axWindowsMediaPlayer1.Ctlcontrols.play();
            }
        }
    }
}
