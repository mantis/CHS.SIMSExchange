using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CHS.SIMSExchange.Service
{
    public partial class Tester : Form
    {
        public Tester()
        {
            InitializeComponent();
            Cover = new CoverThread();
            Worker = new Thread(new ThreadStart(Working));
            Worker.IsBackground = true;
        }

        protected CoverThread Cover;
        protected StaffThread Staff;
        protected RoomsThread Rooms;
        bool stop = false;
        protected Thread Worker;

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Start")
            {
                if (Properties.Settings.Default.Cover) Cover.Bind();
                if (Properties.Settings.Default.Staff) Staff.Bind();
                if (Properties.Settings.Default.Rooms) Rooms.Bind();
                button1.Text = "Stop";
                listBox1.Items.Clear();
                Worker.Start();
            }
            else
            {
                if (Properties.Settings.Default.Cover) Cover.Unbind();
                if (Properties.Settings.Default.Staff) Staff.Unbind();
                if (Properties.Settings.Default.Rooms) Rooms.Unbind();
                button1.Text = "Start";
                stop = true;
                Worker.Abort();
            }
        }

        public void Working()
        {
            while (!stop)
            {
                Thread.Sleep(100);
                if (Cover.Progress != null && Properties.Settings.Default.Cover)
                {
                    if (listBox1.Items.Count > 0)
                    {
                        if (((string)listBox1.Items[listBox1.Items.Count - 1]) != Cover.Progress.Value) listBox1.Invoke(new Action(() => { listBox1.Items.Add(Cover.Progress.Value); }));
                    }
                    else listBox1.Invoke(new Action(() => { listBox1.Items.Add(Cover.Progress.Value); }));
                }
                if (Properties.Settings.Default.Staff && Staff.Progress != null)
                {
                    if (listBox1.Items.Count > 0)
                    {
                        if (((string)listBox1.Items[listBox1.Items.Count - 1]) != Staff.Progress.Value) listBox1.Invoke(new Action(() => { listBox1.Items.Add(Staff.Progress.Value); }));
                    }
                    else listBox1.Invoke(new Action(() => { listBox1.Items.Add(Staff.Progress.Value); }));
                }
                if (Properties.Settings.Default.Rooms && Rooms.Progress != null)
                {
                    if (listBox1.Items.Count > 0)
                    {
                        if (((string)listBox1.Items[listBox1.Items.Count - 1]) != Rooms.Progress.Value) listBox1.Invoke(new Action(() => { listBox1.Items.Add(Rooms.Progress.Value); }));
                    }
                    else listBox1.Invoke(new Action(() => { listBox1.Items.Add(Rooms.Progress.Value); }));
                }
            }
        }

    }
}
