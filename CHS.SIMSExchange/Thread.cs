using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Windows;

namespace CHS.SIMSExchange
{
    public class Thread
    {
        public void Start()
        {
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SIMSerror.log");
            if (File.Exists(logpath)) File.Delete(logpath);
            StreamWriter sr = new StreamWriter(logpath);
            try
            {
                XmlDocument doc = new XmlDocument();
                if (Properties.Settings.Default.AutoSIMS)
                {
                    ProcessStartInfo startinfo;
                    Process process = null;
                    StreamReader stdoutreader;
                    try
                    {
                        startinfo = new ProcessStartInfo();

                        // Set the Parameters for the report

                        startinfo.FileName = @"c:\program files\sims\sims .net\CommandReporter.exe";
                        if (!File.Exists(startinfo.FileName)) startinfo.FileName = @"c:\Program Files (x86)\sims\sims .net\CommandReporter.exe";
                        //string param = "<ReportParameters><Parameter id='EffectiveDate' subreportfilter='FALSE' bypass='TRUE'><Name>EffectiveDate</Name><Type>Date</Type><Values><Date>" + DateTime.Now.ToString("dd/MM/yyyy") + " 00:00:00</Date></Values></Parameter></ReportParameters>";
                        startinfo.Arguments = String.Format("/USER:{0} /PASSWORD:{1} /REPORT:\"{2}\" /QUIET"/* /PARAMS:\"{3}\" */, Properties.Settings.Default.SIMSUser, Properties.Settings.Default.SIMSPassword, "SIMSExchange"/*, param*/);
                        startinfo.UseShellExecute = false;
                        startinfo.RedirectStandardOutput = true;
                        startinfo.CreateNoWindow = true;
                        process = Process.Start(startinfo);
                        stdoutreader = process.StandardOutput;
                        doc = new XmlDocument();
                        doc.Load(stdoutreader);
                        doc.Save(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "report.xml"));
                        process.WaitForExit();
                        stdoutreader.Close();
                        stdoutreader = null;
                    }
                    catch { }
                    finally
                    {
                        if (process != null)
                        {
                            process.Close();
                        }
                        process = null;
                        startinfo = null;
                    }
                }

                if (Properties.Settings.Default.Staff)
                {

                    if (File.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "report.xml"))) doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "report.xml"));
                    this.Progress = 0;
                    XmlNodeList nodes = doc.SelectNodes("/SuperStarReport/Record");
                    staff = new List<Staff>();

                    foreach (XmlNode node in nodes)
                    {
                        this.Progress++;
                        try
                        {
                            if (node.SelectSingleNode("Class") != null || (node.SelectSingleNode("ShortName") != null))
                            {
                                Staff cs = null;
                                foreach (Staff s in staff) if (s.FirstName == (node.SelectSingleNode("Preferred_x0020_Forename") == null ? node.SelectSingleNode("ChosenName") : node.SelectSingleNode("Preferred_x0020_Forename")).InnerText && s.Surname == (node.SelectSingleNode("Legal_x0020_Surname") == null ? node.SelectSingleNode("LegalSurname") : node.SelectSingleNode("Legal_x0020_Surname")).InnerText) { cs = s; break; }
                                if (cs == null) { cs = new Staff(node); staff.Add(cs); }
                                cs.Lessons.Add(new Lesson(node));

                                this.Current = cs.Title + " " + cs.FirstName + " " + cs.Surname + " " + cs.Lessons.Last().Day + " " + cs.Lessons.Last().Start + " " + cs.Lessons.Last().Class;
                            }
                        }
                        catch { }
                        if (Updated != null) Updated();
                    }
                }

                if (Properties.Settings.Default.Rooms)
                {
                    doc = new XmlDocument();
                    if (File.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "rooms.xml"))) doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "rooms.xml"));
                    this.Progress = 0;
                    XmlNodeList nodes = doc.SelectNodes("/SuperStarReport/Record");
                    room = new List<Room>();

                    foreach (XmlNode node in nodes)
                    {
                        this.Progress++;
                        if ((node.SelectSingleNode("Class") != null || node.SelectSingleNode("ShortName") != null) && (node.SelectSingleNode("Name") != null))
                        {
                            Room cr = null;
                            foreach (Room s in room) if (s.Name == node.SelectSingleNode("Name").InnerText) { cr = s; break; }
                            if (cr == null) { cr = new Room(node.SelectSingleNode("Name").InnerText); room.Add(cr); }
                            cr.Add(new Lesson(node));

                            this.Current = cr.Name + " " + cr.Last().Day + " " + cr.Last().Start + " " + cr.Last().Class;
                        }
                        if (Updated != null) Updated();
                    }
                }
            }
            catch (Exception e)
            {
                sr.WriteLine(e.Message);
                sr.WriteLine(e.Source);
                sr.WriteLine(e);
            }
            finally
            {
                sr.Close();
            }
            if (Done != null) Done();
        }

        public event ThreadUpdated Updated;
        public event ThreadUpdated Done;
        public event ThreadInitialized Initialized;
        public int Progress { get; private set; }
        public List<Staff> staff { get; private set; }
        public List<Room> room { get; private set; }
        public string Current { get; private set; }

    }


    public class Thread4
    {
        public void Start(object o)
        {
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Covererror.log");
            if (File.Exists(logpath)) File.Delete(logpath);
            StreamWriter sr = new StreamWriter(logpath);
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                if (string.IsNullOrEmpty(Properties.Settings.Default.Domain))
                    service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword);
                else
                    service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword, Properties.Settings.Default.Domain);
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
                }
                if (Initialized != null) Initialized(count);
                for (var i = 0; i <= Properties.Settings.Default.AdditionalCoverDays; i++)
                {
                    XmlDocument doc = new XmlDocument();
                    //try
                    //{
                    DateTime dt = DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Saturday ? DateTime.Now.AddDays(i + 2) : DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Sunday ? DateTime.Now.AddDays(i + 1) : DateTime.Now.AddDays(i);
                    doc.Load(Path.Combine(Properties.Settings.Default.CoverUNC, "CV" + dt.ToString("ddMMyy") + ".htm"));
                    XmlDocument doc1 = new XmlDocument();
                    doc1.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "staffmapping.xml"));
                    List<string> clearedaccounts = new List<string>();
                    foreach (XmlNode n in doc.SelectNodes("/html/body/table/tr[@bgcolor='#f1f1f1']"))
                    {
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
                                        foreach (Item item in findResults.Items)
                                            ((Appointment)item).Delete(DeleteMode.HardDelete);
                                        clearedaccounts.Add(email);
                                    }
                                    Appointment a = CreateApp(r, dt, service);
                                    if (a != null) a.Save();
                                }
                            }
                            this.Progress++;
                            if (Updated != null) Updated();
                        }
                        catch (Exception ex)
                        {
                            sr.WriteLine(ex.Message);
                            sr.WriteLine(ex.Source);
                            sr.WriteLine(ex);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                sr.WriteLine(e.Message);
                sr.WriteLine(e.Source);
                sr.WriteLine(e);
            }
            finally
            {
                sr.Close();
            }

            if (Done != null) Done();
        }

        public Appointment CreateApp(CoverRow row, DateTime dt, ExchangeService service)
        {
            Dictionary<string, string[]> times = new Dictionary<string,string[]>();
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

        public event ThreadUpdated Updated;
        public event ThreadUpdated Done;
        public event ThreadInitialized Initialized;
        public int Progress { get; private set; }
        public string Current { get; private set; }

        private static bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        private static bool RedirectionUrlValidationCallback(String redirectionUrl)
        {
            return true;
        }

    }


    public delegate void ThreadUpdated();
    public delegate void ThreadInitialized(int Max);
}
