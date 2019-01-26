using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

namespace _2015ck_1
{
    public partial class check : Form
    {
        public check()
        {
            InitializeComponent();
        }

        int count = 0;
        System.Windows.Forms.RadioButton[] radioButtons;

        protected override void WndProc(ref Message m)
        {
            UInt32 WM_DEVICECHANGE = 0x0219;
            UInt32 DEV_DEVTUP_VOLUME = 0x02;
            UInt32 DBT_DEVICEARRIVAL = 0x8000;
            UInt32 DBT_DEVICEREMOVECOMPLETE = 0x8004;

            if ((m.Msg == WM_DEVICECHANGE) &&
                (m.WParam.ToInt32() == DBT_DEVICEARRIVAL))
            {
                int devType = Marshal.ReadInt32(m.LParam, 4);
                if (devType == DEV_DEVTUP_VOLUME)
                {
                    for (int rd = 0; rd < count; rd++)
                    {
                        this.Controls.Remove(radioButtons[rd]);
                    }
                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is RadioButton)
                        {
                            this.Controls.Remove(ctrl);
                        }
                    }
                    count = 0;
                    RefreshDevice();
                }
            }
            if ((m.Msg == WM_DEVICECHANGE) &&
                (m.WParam.ToInt32() == DBT_DEVICEREMOVECOMPLETE))
            {
                int devType = Marshal.ReadInt32(m.LParam, 4);
                if (devType == DEV_DEVTUP_VOLUME)
                {
                    for (int rd = 0; rd < count; rd++)
                    {
                        this.Controls.Remove(radioButtons[rd]);
                    } 
                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is RadioButton)
                        {
                            this.Controls.Remove(ctrl);
                        }
                    }
                    count = 0;
                    RefreshDevice();
                }
            }

            base.WndProc(ref m);
        }

        public void RefreshDevice()
        {
            try
            {
                DriveInfo[] allDrives = DriveInfo.GetDrives();
                
                foreach (DriveInfo d in allDrives)
                {
                    if (d.IsReady == true)
                        count++;
                }
                
                string[] strArr = new string[count];
                string[] strAddress = new string[count];
                int x = 0;
                foreach (DriveInfo d in allDrives)
                {
                    if (d.IsReady == true)
                    {
                        strArr[x] = d.DriveType.ToString() + " : \"" + d.Name + "\", File System : " + d.DriveFormat.ToString();
                        strAddress[x] = d.RootDirectory.ToString();
                        x++;
                        if (x == count)
                            break;
                    }
                }

                radioButtons = new System.Windows.Forms.RadioButton[count];
                for (int i = 0; i < count; i++)
                {
                    radioButtons[i] = new RadioButton();
                    radioButtons[i].Width = 1000;
                    radioButtons[i].Text = strArr[i];
                    radioButtons[i].Location = new System.Drawing.Point(30, 50 + i * 30);
                    this.Controls.Add(radioButtons[i]);
                    radioButtons[i].CheckedChanged += new EventHandler(radioBtn_CheckedChanged);
                }
            }
            catch(Exception)
            {   
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RefreshDevice();
        }

        private void radioBtn_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton R = sender as RadioButton;
            textBox_Dir.Clear();
            string[] dir = R.Text.Split('\"');
            textBox_Dir.Text = dir[1];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ((Form1)(this.Owner)).walkDir = textBox_Dir.Text;
            ((Form1)(this.Owner)).label3.Text = textBox_Dir.Text;
            this.Close();
        }

        private void textBox_Dir_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, e);
        }
    }
}
