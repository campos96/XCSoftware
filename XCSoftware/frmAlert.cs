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

namespace XCSoftware
{
    public partial class frmAlert : Form
    {
        private XCDeviceChannel xCDeviceChannel;

        public bool AlertReloaded { get; set; }

        public frmAlert(XCDeviceChannel xCDeviceChannel)
        {
            this.xCDeviceChannel = xCDeviceChannel;
            InitializeComponent();
        }

        private void frmAlert_Load(object sender, EventArgs e)
        {

            foreach (var scrn in Screen.AllScreens)
            {
                if (scrn.Bounds.Contains(this.Location))
                {
                    this.Location = new Point(scrn.Bounds.Right - this.Width - 50 + 100, scrn.Bounds.Top + 50);
                    break;
                }
            }

            bwStartAlert.RunWorkerAsync();
            UpdateAlert();
        }

        private void UpdateAlert()
        {
            this.Opacity = 1;

            if (xCDeviceChannel.MMDevice == null)
            {
                if (xCDeviceChannel.Value == 0)
                {
                    pictureBox1.Image = Image.FromFile("Icons/volume_off_white.png");
                }
                else
                {
                    pictureBox1.Image = Image.FromFile("Icons/volume_up_white.png");
                }
            }
            else
            {
                if (xCDeviceChannel.Value == 0)
                {
                    if (xCDeviceChannel.MMDevice.DataFlow.ToString() == "Render")
                    {
                        pictureBox1.Image = Image.FromFile("Icons/volume_off_white.png");
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile("Icons/mic_off_white.png");
                    }
                }
                else
                {
                    if (xCDeviceChannel.MMDevice.DataFlow.ToString() == "Capture")
                    {
                        pictureBox1.Image = Image.FromFile("Icons/mic_white.png");
                    }
                    else
                    {
                        pictureBox1.Image = Image.FromFile("Icons/volume_up_white.png");
                    }
                }
            }


            

            panel1.Width = xCDeviceChannel.Value * this.Width / 100;
            label1.Text = xCDeviceChannel.Name;
            label2.Text = $"{xCDeviceChannel.Value}%";
            label3.Text = xCDeviceChannel.MMDevice != null ? xCDeviceChannel.MMDevice.FriendlyName : "N/A";

            if (!bwTimer.IsBusy)
            {
                bwTimer.RunWorkerAsync();
            }


        }

        private void bwTimer_DoWork(object sender, DoWorkEventArgs e)
        {
            AlertReloaded = false;

            for (int i = 0; i < 30; i++)
            {
                if (bwTimer.CancellationPending)
                {
                    AlertReloaded = true;
                    break;
                }
                Thread.Sleep(100);
            }

            if (!AlertReloaded)
            {
                for (int i = 0; i <= 20; i++)
                {
                    if (bwTimer.CancellationPending)
                    {
                        AlertReloaded = true;
                        break;
                    }
                    this.Invoke(new MethodInvoker(delegate { this.Opacity = 1.0 - ((float)i / 20); }));
                    Thread.Sleep(50);
                }
            }
        }

        private void bwTimer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void bwTimer_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (AlertReloaded)
            {
                bwTimer.RunWorkerAsync();
            }
            else
            {
                this.Close();
            }
        }

        internal void Reload(XCDeviceChannel xCDeviceChannel)
        {
            this.xCDeviceChannel = xCDeviceChannel;
            bwTimer.CancelAsync();
            UpdateAlert();
        }

        private void bwStartAlert_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i <= 10; i++)
            {
                this.Invoke(new MethodInvoker(delegate { this.Opacity = ((float)i / 10); }));
                this.Invoke(new MethodInvoker(delegate { this.Location = new Point(this.Location.X - 10, this.Location.Y); }));
                Thread.Sleep(1);
            }

        }
    }
}
