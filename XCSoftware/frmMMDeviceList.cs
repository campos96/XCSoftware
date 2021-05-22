using NAudio.CoreAudioApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XCSoftware
{
    public partial class frmMMDeviceList : Form
    {
        private MMDeviceEnumerator mMDeviceEnumerator;
        private List<MMDevice> MMDevices;
        public MMDevice MMDevice { get; set; }
        public frmMMDeviceList()
        {
            InitializeComponent();
            mMDeviceEnumerator = new MMDeviceEnumerator();
            MMDevices = GetDevices();
        }

        private List<MMDevice> GetDevices()
        {
            return mMDeviceEnumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active).ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            MMDevice = MMDevices.Where(d => d.ID == (string)listBox1.SelectedValue).FirstOrDefault();
            this.Close();
        }

        private void frmMMDeviceList_Load(object sender, EventArgs e)
        {
            listBox1.DataSource = MMDevices;
            listBox1.DisplayMember = "FriendlyName";
            listBox1.ValueMember = "ID";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void listBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            //if the item state is selected them change the back color 
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                e = new DrawItemEventArgs(e.Graphics,
                                          e.Font,
                                          e.Bounds,
                                          e.Index,
                                          e.State ^ DrawItemState.Selected,
                                          e.ForeColor,
                                          Color.FromArgb(0, 192, 0));//Choose the color

            // Draw the background of the ListBox control for each item.
            e.DrawBackground();
            // Draw the current item text
            e.Graphics.DrawString(listBox1.Items[e.Index].ToString(), e.Font, Brushes.White, e.Bounds, StringFormat.GenericDefault);
            // If the ListBox has focus, draw a focus rectangle around the selected item.
            e.DrawFocusRectangle();
        }
    }
}
