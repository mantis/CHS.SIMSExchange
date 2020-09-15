using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml;

namespace CHS.SIMSExchange.Service
{
    public class CoverThread
    {
        public Progress Progress { get; set; }
        System.IO.FileSystemWatcher fws;
        Thread thread;
        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return true;
        }

        private bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        public void Bind()
        {
            thread = new Thread(new ThreadStart(Process));
            Progress = new SIMSExchange.Service.Progress { Finished = true, Value = "" };
            fws = new System.IO.FileSystemWatcher(Properties.Settings.Default.CoverUNC);
            fws.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.CreationTime;
            fws.Changed += fws_Changed;
            fws.Created += fws_Changed;
            fws.EnableRaisingEvents = true;
        }

        public void Unbind()
        {
            fws.Changed -= fws_Changed;
            fws.Created -= fws_Changed;
            fws.EnableRaisingEvents = false;
        }

        void fws_Changed(object sender, System.IO.FileSystemEventArgs e)
        {
            if (this.Progress.Finished && !thread.IsAlive)
            {
                thread = new Thread(new ThreadStart(Process));
                thread.IsBackground = true;
                thread.Start();
            }
        }

        public void Process()
        {
            Progress = new SIMSExchange.Service.Progress() { Finished = false, Value = "" };
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Covererror.log");
            File.WriteAllText(logpath, "");
            StreamWriter sr = new StreamWriter(File.Open(logpath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read));
            sr.AutoFlush = true;
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Wait 5s before start for SIMS to complete it's writing");
            Thread.Sleep(new TimeSpan(0, 0, 5));
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Starting Cover Processing");
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword);
                if (string.IsNullOrEmpty(Properties.Settings.Default.ExchangeUri))
                    service.AutodiscoverUrl(Properties.Settings.Default.EXIMPUser, RedirectionUrlValidationCallback);
                else service.Url = new Uri(Properties.Settings.Default.ExchangeUri + "/ews/exchange.asmx");

                List<Appointment> Appointments = new List<Appointment>();
                int count = 0;
                for (var i = 0; i <= Properties.Settings.Default.AdditionalCoverDays; i++)
                {
                    XmlDocument doc = new XmlDocument();
                    //try
                    //{
                    DateTime dt = DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Saturday ? DateTime.Now.AddDays(i + 2) : DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Sunday ? DateTime.Now.AddDays(i + 1) : DateTime.Now.AddDays(i);
                    doc.Load(Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                    XmlDocument doc1 = new XmlDocument();
                    doc1.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "staffmapping.xml"));
                    count += doc.SelectNodes("/html/body/table/tr[@bgcolor='#f1f1f1']").Count;
                    sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Found " + doc.SelectNodes("/html/body/table/tr[@bgcolor='#f1f1f1']").Count + " cover entries in " + Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                }
                Progress.Total = count;
                Progress.Current = 0;
                sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Found " + count + " cover entries in total to process");
                for (var i = 0; i <= Properties.Settings.Default.AdditionalCoverDays; i++)
                {
                    XmlDocument doc = new XmlDocument();
                    //try
                    //{
                    DateTime dt = DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Saturday ? DateTime.Now.AddDays(i + 2) : DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Sunday ? DateTime.Now.AddDays(i + 1) : DateTime.Now.AddDays(i);
                    doc.Load(Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                    sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Processing " + Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                    XmlDocument doc1 = new XmlDocument();
                    doc1.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "staffmapping.xml"));
                    List<string> clearedaccounts = new List<string>();
                    foreach (XmlNode n in doc.SelectNodes("/html/body/table/tr[@bgcolor='#f1f1f1']"))
                    {
                        sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Processing entry " + Progress.Current + " - Current File: " + Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                        try
                        {
                            CoverRow r = CoverRow.Parse(n);
                            if (r.Cover != "No Cover Reqd")
                            {
                                string email = "";
                                foreach (XmlNode c in doc1.SelectNodes("/staffmappings/staff[@last=\"" + r.Cover.Split(new char[] { ',' })[0] + "\"]"))
                                {
                                    if (r.Cover.ToLower().EndsWith(c.Attributes["first"].Value.ToLower().ToCharArray()[0].ToString()))
                                    {
                                        email = c.Attributes["email"].Value;
                                        sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Found matched staff mapping " + email);
                                        break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(email))
                                {
                                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, email);
                                    if (!clearedaccounts.Contains(email))
                                    {
                                        SearchFilter.SearchFilterCollection searchFilter = new SearchFilter.SearchFilterCollection();
                                        searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, dt.Date));
                                        searchFilter.Add(new SearchFilter.ContainsSubstring(AppointmentSchema.Subject, "Cover:"));
                                        ItemView view = new ItemView(999);
                                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.AppointmentType);
                                        FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Calendar, searchFilter, view);
                                        sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Deleting any existing cover");
                                        foreach (Item item in findResults.Items) ((Appointment)item).Delete(DeleteMode.HardDelete);
                                        clearedaccounts.Add(email);
                                    }
                                    sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Creating cover Cover: " + r.Class);
                                    Appointment a = CreateApp(r, dt, service);
                                    if (a != null) a.Save();
                                }
                            }
                            Progress.Current++;
                        }
                        catch (Exception ex)
                        {
                            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Error - " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Error - " + e.Message);
            }
            finally
            {
                sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Cover Done");
                sr.Close();
                this.Progress.Finished = true;
            }
        }

        public Appointment CreateApp(CoverRow row, DateTime dt, ExchangeService service)
        {
            Dictionary<string, string[]> times = new Dictionary<string, string[]>();
            foreach (string s in Properties.Settings.Default.LessonTimes.Split(new char[] { ':' }))
                if (!string.IsNullOrEmpty(s)) times.Add(s.Split(new char[] { '=' })[0], s.Split(new char[] { '=' })[1].Split(new char[] { '-' }));
            if (row.Period.Trim() == "---" && row.Time.Trim() == "---") return null;
            if (!times.ContainsKey(row.Period.Trim().Split(new char[] { ':' })[1].Trim()) && row.Time.Trim() == "---") return null;
            Appointment appointment = new Appointment(service);
            appointment.Subject = appointment.Body = row.Class;
            appointment.Body += "For Staff: " + row.Absent + "<br />Auto Created: " + DateTime.Now.ToString();
            appointment.Subject = "Cover: " + appointment.Subject;
            if (row.Time == "---")
            {
                string[] time = times[row.Period.Trim().Split(new char[] { ':' })[1].Trim()];
                appointment.Start = new DateTime(dt.Year, dt.Month, dt.Day, int.Parse(time[0].Split(new char[] { '.' })[0]), int.Parse(time[0].Split(new char[] { '.' })[1]), 0);
                appointment.End = new DateTime(dt.Year, dt.Month, dt.Day, int.Parse(time[1].Split(new char[] { '.' })[0]), int.Parse(time[1].Split(new char[] { '.' })[1]), 0);
            }
            else
            {
                appointment.Start = new DateTime(dt.Year, dt.Month, dt.Day, int.Parse(row.Time.Split(new char[] { '-' })[0].Trim().Split(new char[] { ':' })[0]), int.Parse(row.Time.Split(new char[] { '-' })[0].Trim().Split(new char[] { ':' })[1]), 0);
                appointment.End = new DateTime(dt.Year, dt.Month, dt.Day, int.Parse(row.Time.Split(new char[] { '-' })[1].Trim().Split(new char[] { ':' })[0]), int.Parse(row.Time.Split(new char[] { '-' })[1].Trim().Split(new char[] { ':' })[1]), 0);
            }
            appointment.IsReminderSet = false;
            return appointment;
        }


    }
}
