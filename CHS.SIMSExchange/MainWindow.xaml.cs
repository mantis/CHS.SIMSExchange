using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CHS.SIMSExchange
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            thread = new Thread();
            thread.Initialized += new ThreadInitialized(thread_Initialized);
            thread.Updated += new ThreadUpdated(thread_Updated);
            thread.Done += new ThreadUpdated(thread_Done);
            Thread = new System.Threading.Thread(new System.Threading.ThreadStart(thread.Start));
            thread2 = new StaffThread();
            thread2.Initialized += new ThreadInitialized(thread_Initialized);
            thread2.Updated += new ThreadUpdated(thread_Updated);
            thread2.Done += new ThreadUpdated(thread2_Done);
            Thread2 = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(thread2.Start));
            thread3 = new RoomThread();
            thread3.Initialized += new ThreadInitialized(thread3_Initialized);
            thread3.Updated += new ThreadUpdated(thread3_Updated);
            thread3.Done += new ThreadUpdated(thread3_Done);
            Thread3 = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(thread3.Start));
            thread4 = new Thread4();
            thread4.Initialized += new ThreadInitialized(thread4_Initialized);
            thread4.Updated += new ThreadUpdated(thread4_Updated);
            thread4.Done += new ThreadUpdated(thread4_Done);
            Thread4 = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(thread4.Start));
            if (Properties.Settings.Default.Silent) update_Click(this, new RoutedEventArgs());
        }

        private void thread4_Done()
        {
            Dispatcher.BeginInvoke(new Action(thread4_Done1));
        }

        void thread4_Done1()
        {
            message4.Text = "Cover Done";
            if (message.Text == "Done, you can now exit" && Properties.Settings.Default.Silent) Close();
        }

        private void thread4_Updated()
        {
            Dispatcher.BeginInvoke(new ThreadUpdated(thread_Updated4));
        }

        void thread_Updated4()
        {
            progress4.Value = thread4.Progress;
            message4.Text = thread4.Progress + "/" + progress4.Maximum + " " + thread4.Current;
        }

        private void thread4_Initialized(int Max)
        {
            Dispatcher.BeginInvoke(new ThreadInitialized(thread_Initialized4), Max);
        }

        private void thread_Initialized4(int Max)
        {
            progress4.Maximum = Convert.ToDouble(Max);
            message4.Text = "0/" + Max + "";
        }

        private void thread3_Done()
        {
            Dispatcher.BeginInvoke(new Action(thread3_Done1));
        }

        void thread3_Done1()
        {
            message3.Text = "Done, you can now exit";
            if (!Properties.Settings.Default.Staff) message.Text = message3.Text;
            if (message.Text == "Done, you can now exit" && Properties.Settings.Default.Silent) Close();
        }

        private void thread3_Updated()
        {
            Dispatcher.BeginInvoke(new ThreadUpdated(thread_Updated3));
        }

        void thread_Updated3()
        {
            progress3.Value = thread3.Progress;
            message3.Text = thread3.Progress + "/" + progress3.Maximum + " " + thread3.Current;
        }

        private void thread3_Initialized(int Max)
        {
            Dispatcher.BeginInvoke(new ThreadInitialized(thread_Initialized3), Max);
        }

        private void thread_Initialized3(int Max)
        {
            progress3.Maximum = Convert.ToDouble(Max);
            message3.Text = "0/" + Max + "";
        }

        void thread_Done()
        {
            Dispatcher.BeginInvoke(new Action(thread_Done1));
        }

        void thread_Done1()
        {
            if (Thread2.ThreadState != System.Threading.ThreadState.Running && Properties.Settings.Default.Staff) Thread2.Start(thread.staff);
            if (Thread3.ThreadState != System.Threading.ThreadState.Running && Properties.Settings.Default.Rooms) Thread3.Start(thread.room);
            job.Text = "Step 2: Export to Exchange";
            message.Text = "Connecting to Exchange";
            message2.Text = message3.Text = "";
        }

        void thread2_Done()
        {
            Dispatcher.BeginInvoke(new Action(thread2_Done1));
        }

        void thread2_Done1()
        {
            message.Text = "Done, you can now exit";
            message3.Text = "";
            message2.Text = "";
            if (message.Text == "Done, you can now exit" && Properties.Settings.Default.Silent) Close();
        }

        void thread_Updated()
        {
            Dispatcher.BeginInvoke(new ThreadUpdated(thread_Updated1));
        }

        void thread_Updated1()
        {
            if (progress.Value == progress.Maximum)
            {
                progress2.Value = thread2.Progress;
                message.Text = thread2.Progress + "/" + progress2.Maximum;
                message2.Text = thread2.Current;
            }
            else
            {
                progress.Value = thread.Progress;
                message.Text = thread.Progress + "/" + progress.Maximum;
                message2.Text = thread.Current;
            }
        }

        void thread_Initialized(int Max)
        {
            Dispatcher.BeginInvoke(new ThreadInitialized(thread_Initialized1), Max);
        }

        void thread_Initialized1(int Max)
        {
            if (progress.Value == progress.Maximum)
                progress2.Maximum = Convert.ToDouble(Max);
            else
                progress.Maximum = Convert.ToDouble(Max);
            message.Text = "0/" + Max;
            message2.Text = "";
        }

        private Thread thread { get; set; }
        private StaffThread thread2 { get; set; }
        private RoomThread thread3 { get; set; }
        private Thread4 thread4 { get; set; }
        private System.Threading.Thread Thread { get; set; }
        private System.Threading.Thread Thread2 { get; set; }
        private System.Threading.Thread Thread3 { get; set; }
        private System.Threading.Thread Thread4 { get; set; }

        private void update_Click(object sender, RoutedEventArgs e)
        {
            job.Text = "Step 1 - Importing from SIMS";
            message.Text = "Starting, Connecting to SIMS.net...";
            message2.Text = message3.Text = "";
            progress.IsEnabled = true;
            update.IsEnabled = false;
            if (Properties.Settings.Default.Staff || Properties.Settings.Default.Rooms) Thread.Start();
            else message.Text = "Done, you can now exit";
            if (Properties.Settings.Default.Cover) Thread4.Start();
        }
    }
}
