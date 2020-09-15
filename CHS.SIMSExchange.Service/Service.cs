using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;

namespace CHS.SIMSExchange.Service
{
    [ComVisible(true)]
    partial class Service : ServiceBase
    {
        public Service()
        {
            InitializeComponent();
            Cover = new CoverThread();
            Rooms = new RoomsThread();
            Staff = new StaffThread();
        }

        protected CoverThread Cover;
        protected StaffThread Staff;
        protected RoomsThread Rooms;
        protected override void OnStart(string[] args)
        {
            if (Properties.Settings.Default.Cover) Cover.Bind();
            if (Properties.Settings.Default.Staff) Staff.Bind();
            if (Properties.Settings.Default.Rooms) Rooms.Bind();
        }

        protected override void OnStop()
        {
            if (Properties.Settings.Default.Cover) Cover.Unbind();
            if (Properties.Settings.Default.Staff) Staff.Unbind();
            if (Properties.Settings.Default.Rooms) Rooms.Unbind();
        }
    }
}
