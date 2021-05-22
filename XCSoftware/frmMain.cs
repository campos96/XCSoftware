using NAudio.CoreAudioApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XCSoftware
{
    public partial class frmMain : Form
    {
        private MMDeviceEnumerator mMDeviceEnumerator;
        private List<MMDevice> MMDevices;
        private XCDevice xCDevice { get; set; }
        SerialPort Port;
        bool PortClosed = false;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd,
                         int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        public frmMain()
        {
            InitializeComponent();

            mMDeviceEnumerator = new MMDeviceEnumerator();
            xCDevice = new XCDevice { ID = 1, Name = "XC-01", SerialNumber = "XCD202102050001" };
            xCDevice.XCDeviceChannels.Add(new XCDeviceChannel { ID = 1, Name = "Channel 1", Value = 25 });
            xCDevice.XCDeviceChannels.Add(new XCDeviceChannel { ID = 2, Name = "Channel 2", Value = 50 });
            xCDevice.XCDeviceChannels.Add(new XCDeviceChannel { ID = 3, Name = "Channel 3", Value = 75 });
            xCDevice.XCDeviceChannels.Add(new XCDeviceChannel { ID = 4, Name = "Channel 4", Value = 100 });
            this.AutoSize = true;

            try
            {
                Port = new System.IO.Ports.SerialPort();
                Port.PortName = "COM3";
                Port.BaudRate = 9600;
                Port.ReadTimeout = 500;
                Port.Open();
            }
            catch
            {
                MessageBox.Show("Unable to connect with device...");
                this.Close();
            }

            ReadConfigFile();
        }

        private void ReadConfigFile()
        {
            try
            {
                if (!File.Exists("config_backup.txt"))
                {
                    File.Create("config_backup.txt").Close();
                    using (StreamWriter sw = File.CreateText("config_backup.txt"))
                    {
                        string jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(xCDevice);
                        sw.WriteLine(jsonString);
                    }
                }
                else
                {
                    string jsonString = System.IO.File.ReadAllText("config_backup.txt");
                    var _xCDevice = Newtonsoft.Json.JsonConvert.DeserializeObject<XCDevice>(jsonString);
                    if (_xCDevice != null)
                    {
                        foreach (var xCDeviceChannel in _xCDevice.XCDeviceChannels)
                        {
                            MMDevices = GetDevices();
                            MMDevice _mMDevice = MMDevices.Where(x => xCDeviceChannel.MMDeviceID != null ? x.ID == xCDeviceChannel.MMDeviceID : false).FirstOrDefault();
                            if (_mMDevice != null)
                            {
                                var _xCDeviceChannel = xCDevice.XCDeviceChannels.Where(x => x.ID == xCDeviceChannel.ID).FirstOrDefault();
                                if (_xCDeviceChannel != null)
                                {
                                    _xCDeviceChannel.MMDevice = _mMDevice;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot read config file.");
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            LoadXCDeviceChannels();
            Thread thread = new Thread(ListenSerial);
            thread.Start();
        }

        private void LoadXCDeviceChannels()
        {
            flpContainer.Controls.Clear();

            foreach (var xCDeviceChannel in xCDevice.XCDeviceChannels)
            {
                Panel pnlXCDeviceChannel = new Panel();
                pnlXCDeviceChannel.Name = $"pnlXCDeviceChannel{xCDeviceChannel.ID}";
                pnlXCDeviceChannel.Margin = new Padding(10);
                pnlXCDeviceChannel.Size = new Size(flpContainer.Width - 20, 100);
                pnlXCDeviceChannel.Anchor = AnchorStyles.Right;
                pnlXCDeviceChannel.BackColor = Color.Black;
                //pnlXCDeviceChannel.BorderStyle = BorderStyle.FixedSingle;

                PictureBox pbVolumeImage = new PictureBox();
                pbVolumeImage.Name = $"pbVolumeImage{xCDeviceChannel.ID}";
                pbVolumeImage.Size = new Size(pnlXCDeviceChannel.Height - 20, pnlXCDeviceChannel.Height - 20);
                pbVolumeImage.Location = new Point(10, 10);
                pbVolumeImage.Anchor = AnchorStyles.Right;
                pbVolumeImage.SizeMode = PictureBoxSizeMode.Zoom;
                pnlXCDeviceChannel.Controls.Add(pbVolumeImage);

                Label lblXCDeviceChannelMMDeviceName = new Label();
                lblXCDeviceChannelMMDeviceName.Name = $"lblXCDeviceChannelMMDeviceName{xCDeviceChannel.ID}";
                lblXCDeviceChannelMMDeviceName.Font = new Font("Arial", 14, FontStyle.Bold);
                lblXCDeviceChannelMMDeviceName.Text = $"{(xCDeviceChannel.MMDevice != null ? (string)xCDeviceChannel.MMDevice.FriendlyName : "N/A")}";
                lblXCDeviceChannelMMDeviceName.Location = new Point(pbVolumeImage.Location.X + pbVolumeImage.Width + 20, 20);
                lblXCDeviceChannelMMDeviceName.AutoSize = true;
                lblXCDeviceChannelMMDeviceName.ForeColor = Color.White;
                pnlXCDeviceChannel.Controls.Add(lblXCDeviceChannelMMDeviceName);

                Label lblXCDeviceChannelName = new Label();
                lblXCDeviceChannelName.Name = $"lblXCDeviceChannelName{xCDeviceChannel.ID}";
                lblXCDeviceChannelName.Font = new Font("Arial", 12, FontStyle.Regular);
                lblXCDeviceChannelName.Text = $"{xCDeviceChannel.Name}";
                lblXCDeviceChannelName.Location = new Point(pbVolumeImage.Location.X + pbVolumeImage.Width + 20, 53);
                lblXCDeviceChannelName.AutoSize = true;
                lblXCDeviceChannelName.ForeColor = Color.DarkGray;
                pnlXCDeviceChannel.Controls.Add(lblXCDeviceChannelName);

                Label lblXCDeviceChannelVolume = new Label();
                lblXCDeviceChannelVolume.Name = $"lblXCDeviceChannelVolume{xCDeviceChannel.ID}";
                lblXCDeviceChannelVolume.Font = new Font("Arial", 24, FontStyle.Bold);
                lblXCDeviceChannelVolume.Text = $"{xCDeviceChannel.Value}%";
                lblXCDeviceChannelVolume.AutoSize = true;
                lblXCDeviceChannelVolume.TextAlign = ContentAlignment.MiddleCenter;
                lblXCDeviceChannelVolume.ForeColor = Color.White;
                lblXCDeviceChannelVolume.AutoSize = false;
                lblXCDeviceChannelVolume.Size = new Size(pnlXCDeviceChannel.Height, pnlXCDeviceChannel.Height - 20);
                lblXCDeviceChannelVolume.Location = new Point((pnlXCDeviceChannel.Width - lblXCDeviceChannelVolume.Width) - 10, 10);
                pnlXCDeviceChannel.Controls.Add(lblXCDeviceChannelVolume);


                Button btnXCDeviceChannelAssign = new Button();
                btnXCDeviceChannelAssign.Name = $"btnXCDeviceChannelAssign{xCDeviceChannel.ID}";
                btnXCDeviceChannelAssign.Font = new Font("Arial", 10, FontStyle.Regular);
                btnXCDeviceChannelAssign.Size = new Size(130, 25);
                btnXCDeviceChannelAssign.FlatStyle = FlatStyle.Flat;
                btnXCDeviceChannelAssign.FlatAppearance.BorderSize = 0;
                btnXCDeviceChannelAssign.BackColor = Color.FromArgb(0, 192, 0);
                btnXCDeviceChannelAssign.ForeColor = Color.Black;
                btnXCDeviceChannelAssign.TextAlign = ContentAlignment.MiddleCenter;
                btnXCDeviceChannelAssign.Location = new Point(lblXCDeviceChannelName.Location.X + lblXCDeviceChannelName.Width + 5, 50);

                if (xCDeviceChannel.MMDevice != null)
                {
                    btnXCDeviceChannelAssign.Text = $"Change Device";

                    Button btnXCDeviceChannelRemove = new Button();
                    btnXCDeviceChannelRemove.Name = $"btnXCDeviceChannelRemove{xCDeviceChannel.ID}";
                    btnXCDeviceChannelRemove.Text = $"Remove Device";
                    btnXCDeviceChannelRemove.Font = new Font("Arial", 10, FontStyle.Regular);
                    btnXCDeviceChannelRemove.Size = new Size(130, 25);
                    btnXCDeviceChannelRemove.FlatStyle = FlatStyle.Flat;
                    btnXCDeviceChannelRemove.FlatAppearance.BorderSize = 0;
                    btnXCDeviceChannelRemove.BackColor = Color.FromArgb(0, 192, 0);
                    btnXCDeviceChannelRemove.ForeColor = Color.Black;
                    btnXCDeviceChannelRemove.TextAlign = ContentAlignment.MiddleCenter;
                    btnXCDeviceChannelRemove.Location = new Point(btnXCDeviceChannelAssign.Location.X + btnXCDeviceChannelAssign.Width + 5, 50);
                    btnXCDeviceChannelRemove.Click += new EventHandler((s, ev) => btnXCDeviceChannelRemove_Click(s, ev, xCDeviceChannel.ID));
                    pnlXCDeviceChannel.Controls.Add(btnXCDeviceChannelRemove);
                }
                else
                {
                    btnXCDeviceChannelAssign.Text = $"Assing Device";
                }

                btnXCDeviceChannelAssign.Click += new EventHandler((s, ev) => btnXCDeviceChannelAssign_Click(s, ev, xCDeviceChannel.ID));
                pnlXCDeviceChannel.Controls.Add(btnXCDeviceChannelAssign);

                Panel pnlXCDeviceChannelVolumeBar = new Panel();
                pnlXCDeviceChannelVolumeBar.Name = $"pnlXCDeviceChannelVolumeBar{xCDeviceChannel.ID}";
                pnlXCDeviceChannelVolumeBar.Margin = new Padding(10);
                pnlXCDeviceChannelVolumeBar.Size = new Size(pnlXCDeviceChannel.Width, 5);
                pnlXCDeviceChannelVolumeBar.BackColor = Color.FromArgb(0, 192, 0);
                pnlXCDeviceChannelVolumeBar.Location = new Point(0, pnlXCDeviceChannel.Height - 10);
                pnlXCDeviceChannelVolumeBar.Anchor = AnchorStyles.Right;
                pnlXCDeviceChannelVolumeBar.BorderStyle = BorderStyle.FixedSingle;
                pnlXCDeviceChannel.Controls.Add(pnlXCDeviceChannelVolumeBar);

                TrackBar tbXCDeviceChannelValue = new TrackBar();
                tbXCDeviceChannelValue.Name = $"tbXCDeviceChannelValue{xCDeviceChannel.ID}";
                tbXCDeviceChannelValue.Size = new Size(flpContainer.Width - 40, 0);
                tbXCDeviceChannelValue.Location = new Point(10, 35);
                tbXCDeviceChannelValue.Minimum = 0;
                tbXCDeviceChannelValue.Maximum = 100;
                tbXCDeviceChannelValue.SmallChange = 1;
                tbXCDeviceChannelValue.Visible = false;
                tbXCDeviceChannelValue.LargeChange = 10;
                tbXCDeviceChannelValue.Value = xCDeviceChannel.Value;
                tbXCDeviceChannelValue.ValueChanged += new EventHandler((s, ev) => tbXCDeviceChannelValue_ValueChanged(s, ev, xCDeviceChannel, tbXCDeviceChannelValue.Value));

                if (xCDeviceChannel.MMDevice == null)
                {
                    tbXCDeviceChannelValue.Enabled = false;
                }
                pnlXCDeviceChannel.Controls.Add(tbXCDeviceChannelValue);


                flpContainer.Controls.Add(pnlXCDeviceChannel);

                UpdatePanelView(xCDeviceChannel);
            }

        }

        private void btnXCDeviceChannelRemove_Click(object s, EventArgs ev, int xCDeviceChannelID)
        {
            var xCDeviceChannel = xCDevice.XCDeviceChannels.Where(c => c.ID == xCDeviceChannelID).FirstOrDefault();
            if (xCDeviceChannel != null)
            {
                xCDeviceChannel.MMDevice = null;
                LoadXCDeviceChannels();
                Alert(xCDeviceChannel);
            }
        }

        private void tbXCDeviceChannelValue_ValueChanged(object s, EventArgs ev, XCDeviceChannel xCDeviceChannel, int volume)
        {
            SetVolume(xCDeviceChannel, volume);
        }

        private void SetVolume(XCDeviceChannel xCDeviceChannel, int volume)
        {
            MMDevices = GetDevices();
            MMDevice mMDevice = MMDevices.Where(od => xCDeviceChannel.MMDevice != null ? od.ID == xCDeviceChannel.MMDevice.ID : false).FirstOrDefault();

            xCDeviceChannel.Value = volume;
            var scaledVolume = (float)Math.Max(Math.Min(xCDeviceChannel.Value, 100), 0) / (float)100;

            UpdatePanelView(xCDeviceChannel);


            Alert(xCDeviceChannel);


            if (mMDevice != null)
            {
                if (mMDevice.State == DeviceState.Active)
                {
                    try
                    {
                        mMDevice.AudioEndpointVolume.MasterVolumeLevelScalar = scaledVolume;

                    }
                    catch { }
                }

            }
        }

        private void Alert(XCDeviceChannel xCDeviceChannel)
        {

            frmAlert frmAlert = Application.OpenForms.OfType<frmAlert>().FirstOrDefault();

            if (frmAlert != null)
            {
                frmAlert.Reload(xCDeviceChannel);
            }
            else
            {
                frmAlert = new frmAlert(xCDeviceChannel);
                frmAlert.Show();
            }

            UpdateConfigFile();
        }

        private void UpdateConfigFile()
        {
            try
            {
                using (StreamWriter sw = File.CreateText("config_backup.txt"))
                {
                    var _xcDevice = new XCDevice
                    {
                        ID = xCDevice.ID,
                        Name = xCDevice.Name,
                        SerialNumber = xCDevice.SerialNumber
                    };

                    foreach (var xCDeviceChannel in xCDevice.XCDeviceChannels)
                    {
                        _xcDevice.XCDeviceChannels.Add(new XCDeviceChannel
                        {
                            ID = xCDeviceChannel.ID,
                            Name = xCDeviceChannel.Name,
                            Value = xCDeviceChannel.Value,
                            MMDeviceID = xCDeviceChannel.MMDevice != null ? xCDeviceChannel.MMDevice.ID : null
                        });
                    }

                    string jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(_xcDevice);
                    sw.WriteLine(jsonString);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot update config file.");
            }
        }

        private void UpdatePanelView(XCDeviceChannel xCDeviceChannel)
        {
            Control lblXCDeviceChannelMMDeviceName = flpContainer.Controls.Find($"lblXCDeviceChannelMMDeviceName{xCDeviceChannel.ID}", true).FirstOrDefault();
            if (lblXCDeviceChannelMMDeviceName != null)
            {
                lblXCDeviceChannelMMDeviceName.Text = $"{(xCDeviceChannel.MMDevice != null ? (string)xCDeviceChannel.MMDevice.FriendlyName : "N/A")}";

            }

            Control lblXCDeviceChannelName = flpContainer.Controls.Find($"lblXCDeviceChannelName{xCDeviceChannel.ID}", true).FirstOrDefault();
            if (lblXCDeviceChannelName != null)
            {
                lblXCDeviceChannelName.Text = $"{xCDeviceChannel.Name}";

            }

            Control lblXCDeviceChannelVolume = flpContainer.Controls.Find($"lblXCDeviceChannelVolume{xCDeviceChannel.ID}", true).FirstOrDefault();
            if (lblXCDeviceChannelVolume != null)
            {
                lblXCDeviceChannelVolume.Text = $"{xCDeviceChannel.Value}%";

            }

            PictureBox pbVolumeImage = (PictureBox)flpContainer.Controls.Find($"pbVolumeImage{xCDeviceChannel.ID}", true).FirstOrDefault();
            if (pbVolumeImage != null)
            {
                if (xCDeviceChannel.MMDevice == null)
                {
                    if (xCDeviceChannel.Value == 0)
                    {
                        pbVolumeImage.Image = Image.FromFile("Icons/volume_off_white.png");
                    }
                    else
                    {
                        pbVolumeImage.Image = Image.FromFile("Icons/volume_up_white.png");
                    }
                }
                else
                {
                    if (xCDeviceChannel.Value == 0)
                    {
                        if (xCDeviceChannel.MMDevice.DataFlow.ToString() == "Render")
                        {
                            pbVolumeImage.Image = Image.FromFile("Icons/volume_off_white.png");
                        }
                        else
                        {
                            pbVolumeImage.Image = Image.FromFile("Icons/mic_off_white.png");
                        }
                    }
                    else
                    {
                        if (xCDeviceChannel.MMDevice.DataFlow.ToString() == "Capture")
                        {
                            pbVolumeImage.Image = Image.FromFile("Icons/mic_white.png");
                        }
                        else
                        {
                            pbVolumeImage.Image = Image.FromFile("Icons/volume_up_white.png");
                        }
                    }
                }

            }

            Control pnlXCDeviceChannelVolumeBar = flpContainer.Controls.Find($"pnlXCDeviceChannelVolumeBar{xCDeviceChannel.ID}", true).FirstOrDefault();
            if (pnlXCDeviceChannelVolumeBar != null)
            {
                pnlXCDeviceChannelVolumeBar.Width = xCDeviceChannel.Value * flpContainer.Width / 100;

            }

        }

        private void btnXCDeviceChannelAssign_Click(object s, EventArgs ev, int xCDeviceChannelID)
        {
            frmMMDeviceList frmMMDeviceList = new frmMMDeviceList();
            frmMMDeviceList.ShowDialog();
            if (frmMMDeviceList.MMDevice != null)
            {
                //check if is already assigned
                if (xCDevice.XCDeviceChannels.Where(c => c.MMDevice != null ? c.MMDevice.FriendlyName == frmMMDeviceList.MMDevice.FriendlyName : false).Count() > 0)
                {
                    MessageBox.Show($"{frmMMDeviceList.MMDevice.FriendlyName} is already assigned");
                }
                else
                {
                    var xCDeviceChannel = xCDevice.XCDeviceChannels.Where(c => c.ID == xCDeviceChannelID).FirstOrDefault();
                    if (xCDeviceChannel != null)
                    {
                        xCDeviceChannel.MMDevice = frmMMDeviceList.MMDevice;
                        LoadXCDeviceChannels();
                        Alert(xCDeviceChannel);
                    }
                }
            }
        }

        private List<MMDevice> GetDevices()
        {
            return mMDeviceEnumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active).ToList();
        }

        private void ListenSerial()
        {
            while (!PortClosed)
            {
                try
                {
                    //read to data from arduino
                    string responseString = Port.ReadLine();

                    var _XCDevice = Newtonsoft.Json.JsonConvert.DeserializeObject<XCDevice>(responseString);

                    if (_XCDevice != null)
                    {
                        foreach (var xCDeviceChannel in xCDevice.XCDeviceChannels)
                        {
                            var _xCDeviceChannel = _XCDevice.XCDeviceChannels.Where(x => x.ID == xCDeviceChannel.ID).FirstOrDefault();

                            if (_xCDeviceChannel != null)
                            {
                                if (_xCDeviceChannel.Value != xCDeviceChannel.Value)
                                {
                                    flpContainer.Invoke(new MethodInvoker(
                                        delegate
                                        {
                                            TrackBar tbXCDeviceChannelValue = (TrackBar)flpContainer.Controls.Find($"tbXCDeviceChannelValue{xCDeviceChannel.ID}", true).FirstOrDefault();
                                            if (tbXCDeviceChannelValue != null)
                                            {
                                                tbXCDeviceChannelValue.Value = _xCDeviceChannel.Value;
                                            }
                                        }
                                        ));

                                }
                            }
                        }
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }

                Thread.Sleep(100);
            }
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            PortClosed = true;
            if (Port.IsOpen)
            {
                Port.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
    }
}
