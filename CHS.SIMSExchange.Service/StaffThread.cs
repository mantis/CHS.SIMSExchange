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
    public class StaffThread
    {
        public Progress Progress { get; set; }
        FileSystemWatcher fws;
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
            fws = new System.IO.FileSystemWatcher(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "report.xml");
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

        public List<Staff> staff { get; private set; }

        public void Process()
        {
            Progress = new SIMSExchange.Service.Progress() { Finished = false, Value = "" };
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Staff.log");
            File.WriteAllText(logpath, "");
            StreamWriter sr = new StreamWriter(File.Open(logpath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite));
            sr.AutoFlush = true;
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Wait 10s before start for SIMS to complete it's writing");
            Thread.Sleep(new TimeSpan(0, 0, 10));
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Starting Staff Processing");
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "report.xml"));
                XmlNodeList nodes = doc.SelectNodes("/SuperStarReport/Record");
                this.Progress.Total = 0;
                List<Staff> staff = new List<Staff>();

                foreach (XmlNode node in nodes)
                {
                    this.Progress.Total++;
                    if (node.SelectSingleNode("Class") != null || (node.SelectSingleNode("ShortName") != null))
                    {
                        Staff cs = null;
                        foreach (Staff s in staff) if (s.FirstName == (node.SelectSingleNode("Preferred_x0020_Forename") == null ? node.SelectSingleNode("ChosenName") : node.SelectSingleNode("Preferred_x0020_Forename")).InnerText && s.Surname == (node.SelectSingleNode("Legal_x0020_Surname") == null ? node.SelectSingleNode("LegalSurname") : node.SelectSingleNode("Legal_x0020_Surname")).InnerText) { cs = s; break; }
                        if (cs == null) { cs = new Staff(node); staff.Add(cs); }
                        cs.Lessons.Add(new Lesson(node));

                        sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Found " + cs.Title + " " + cs.FirstName + " " + cs.Surname + " " + cs.Lessons.Last().Day + " " + cs.Lessons.Last().Start + " " + cs.Lessons.Last().Class);
                    }
                }
                sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Found " + this.Progress.Total + " staff events");

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                if (string.IsNullOrEmpty(Properties.Settings.Default.Domain))
                    service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword);
                else
                    service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword, Properties.Settings.Default.Domain);
                if (string.IsNullOrEmpty(Properties.Settings.Default.ExchangeUri))
                    service.AutodiscoverUrl(Properties.Settings.Default.EXIMPUser, RedirectionUrlValidationCallback);
                else service.Url = new Uri(Properties.Settings.Default.ExchangeUri + "/ews/exchange.asmx");
                Terms terms = new Terms();
                List<Appointment> Appointments = new List<Appointment>();

                this.Progress.Current = 0;

                sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Loading Staff Mappings");
                XmlDocument doc2 = new XmlDocument();
                doc2.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "staffmapping.xml"));


                foreach (Staff s in staff)
                {
                    try
                    {
                        if (doc.SelectSingleNode("/staffmappings/staff[@first=\"" + s.FirstName + "\" and @last=\"" + s.Surname + "\"]") != null)
                        {
                            sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Removing previous appointments for " + s.Title + " " + s.FirstName + " " + s.Surname);
                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, doc.SelectSingleNode("/staffmappings/staff[@first=\"" + s.FirstName + "\" and @last=\"" + s.Surname + "\"]").Attributes["email"].Value);


                            SearchFilter.SearchFilterCollection searchFilter = new SearchFilter.SearchFilterCollection();
                            searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, terms[0].StartDate));
                            searchFilter.Add(new SearchFilter.ContainsSubstring(AppointmentSchema.Subject, "Lesson:"));
                            bool removecompleted = false;
                            while (!removecompleted)
                            {
                                ItemView view = new ItemView(1000);
                                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.AppointmentType);
                                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Calendar, searchFilter, view);
                                Console.WriteLine(findResults.TotalCount + " " + findResults.Items.Count);
                                var appsids = new List<ItemId>();
                                foreach (Item item in findResults.Items)
                                {
                                    Appointment appt = item as Appointment;
                                    if (appt.AppointmentType == AppointmentType.RecurringMaster)
                                        appsids.Add(appt.Id);
                                }
                                sr.WriteLine("Removing " + appsids.Count);
                                if (appsids.Count > 0)
                                    service.DeleteItems(appsids, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences, true);
                                var c = service.FindItems(WellKnownFolderName.Calendar, searchFilter, view).TotalCount;
                                removecompleted = c == 0;
                                if (!removecompleted) removecompleted = c == 1;
                                if (!removecompleted) { sr.WriteLine("Remove not completed, still " + c + " to remove"); System.Threading.Thread.Sleep(5000); }
                            }
                            sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + "Creating appointments for " + s.Title + " " + s.FirstName + " " + s.Surname);
                            var apps = new List<Appointment>();
                            foreach (Lesson l in s.Lessons)
                            {
                                foreach (Term t in terms)
                                {
                                    if (t.HalfTerm.HasValue)
                                    {
                                        DateTime hts = t.HalfTerm.Value.StartDate;
                                        if (hts.DayOfWeek == DayOfWeek.Monday) hts = hts.AddDays(-3);
                                        DateTime hte = t.HalfTerm.Value.EndDate;
                                        if (hte.DayOfWeek == DayOfWeek.Friday) hte = hte.AddDays(3);
                                        if (apps.Count(f => f.Body.Text.Contains(GenBody(l, t.StartDate, hts))) == 0) apps.Add(CreateApp(l, t.StartDate, hts, t.StartWeekNum == 2, service));
                                        if (hte < t.EndDate)
                                        {
                                            if (apps.Count(f => f.Body.Text.Contains(GenBody(l, hte, t.EndDate))) == 0) apps.Add(CreateApp(l, hte, t.EndDate, t.WeekNum(hte) == 2, service));
                                        }
                                    }
                                    else if (apps.Count(f => f.Body.Text.Contains(GenBody(l, t.StartDate, t.EndDate))) == 0)
                                        apps.Add(CreateApp(l, t.StartDate, t.EndDate, t.StartWeekNum == 2, service));
                                }
                            }
                            Console.WriteLine("Creating " + apps.Count + " appointments");
                            service.CreateItems(apps, null, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
                        }
                    }
                    catch (Exception ex)
                    {
                        sr.WriteLine(DateTime.Now.ToString() + " Error " + ex.Message);
                        sr.WriteLine(DateTime.Now.ToString() + " Error " + ex.Source);
                        sr.WriteLine(DateTime.Now.ToString() + " Error " + ex);
                    }
                    this.Progress.Current++;
                }
            }
            catch (Exception e)
            {
                sr.WriteLine(DateTime.Now.ToString() + " Error " + e.Message);
                sr.WriteLine(DateTime.Now.ToString() + " Error " + e.Source);
                sr.WriteLine(DateTime.Now.ToString() + " Error " + e);
            }
            finally
            {
                sr.Close();
                this.Progress.Finished = true;
            }
        }

        public string GenBody(Lesson l, DateTime start, DateTime end)
        {
            return "Lesson " + l.Class + " on day " + l.Day + " from " + l.Start + " to " + l.End + " from " + start.ToShortDateString() + " to " + end.ToShortDateString();
        }

        public Appointment CreateApp(Lesson l, DateTime start, DateTime end, bool week2, ExchangeService service)
        {
            Appointment appointment = new Appointment(service);
            appointment.Body = GenBody(l, start, end) + "<br />Auto Created: " + DateTime.Now.ToString();
            appointment.Subject = "Lesson: " + l.Class;
            int plusday = l.Day;
            plusday--;
            if (week2)
            {
                if (plusday < 5) plusday += 5;
                else if (plusday > 4) plusday -= 5;
            }
            if (plusday > 4) plusday += 2;
            if (start.DayOfWeek == DayOfWeek.Tuesday) plusday--;
            if (start.DayOfWeek == DayOfWeek.Wednesday) plusday -= 2;
            appointment.Start = new DateTime(start.AddDays(plusday).Year, start.AddDays(plusday).Month, start.AddDays(plusday).Day, int.Parse(l.Start.Split(new char[] { ':' })[0]), int.Parse(l.Start.Split(new char[] { ':' })[1]), 0);
            appointment.End = new DateTime(start.AddDays(plusday).Year, start.AddDays(plusday).Month, start.AddDays(plusday).Day, int.Parse(l.End.Split(new char[] { ':' })[0]), int.Parse(l.End.Split(new char[] { ':' })[1]), 0);
            appointment.Location = l.Room;
            appointment.IsReminderSet = false;
            appointment.Recurrence = new Recurrence.WeeklyPattern(appointment.Start.Date, (Properties.Settings.Default.TwoWeekTimetable ? 2 : 1), new DayOfTheWeek[] { CovertDOTW(appointment.Start) });
            appointment.Recurrence.StartDate = appointment.Start.Date;
            appointment.Recurrence.EndDate = end.Date;

            return appointment;
        }

        public DayOfTheWeek CovertDOTW(DateTime date)
        {
            if (date.DayOfWeek == DayOfWeek.Monday) return DayOfTheWeek.Monday;
            else if (date.DayOfWeek == DayOfWeek.Tuesday) return DayOfTheWeek.Tuesday;
            else if (date.DayOfWeek == DayOfWeek.Wednesday) return DayOfTheWeek.Wednesday;
            else if (date.DayOfWeek == DayOfWeek.Thursday) return DayOfTheWeek.Thursday;
            else return DayOfTheWeek.Friday;
        }
    }
}
