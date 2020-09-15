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
    public class RoomsThread
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
            fws = new System.IO.FileSystemWatcher(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "rooms.xml");
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
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Rooms.log");
            File.WriteAllText(logpath, "");
            StreamWriter sr = new StreamWriter(File.Open(logpath, FileMode.Create, FileAccess.ReadWrite, FileShare.Read));
            sr.AutoFlush = true;
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Wait 10s before start for SIMS to complete it's writing");
            Thread.Sleep(new TimeSpan(0, 0, 10));
            sr.WriteLine(Progress.Value = DateTime.Now.ToString() + " Starting Room Processing");
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "rooms.xml"));
                XmlNodeList nodes = doc.SelectNodes("/SuperStarReport/Record");
                List<Room> room = new List<Room>();
                this.Progress.Total = 0;
                foreach (XmlNode node in nodes)
                {
                    this.Progress.Total++;
                    if (node.SelectSingleNode("Class") != null || (node.SelectSingleNode("Name") != null))
                    {
                        Room cr = null;
                        foreach (Room s in room) if (s.Name == node.SelectSingleNode("Name").InnerText) { cr = s; break; }
                        if (cr == null) { cr = new Room(node.SelectSingleNode("Name").InnerText); room.Add(cr); }
                        cr.Add(new Lesson(node));

                        sr.WriteLine(this.Progress.Value = this.Progress.Value = DateTime.Now.ToString() + "Parsing " + cr.Name + " " + cr.Last().Day + " " + cr.Last().Start + " " + cr.Last().Class);
                    }
                }
                sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Found " + this.Progress.Total + " room events");
                
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.Credentials = new WebCredentials(Properties.Settings.Default.EXIMPUser, Properties.Settings.Default.EXIMPPassword);
                if (string.IsNullOrEmpty(Properties.Settings.Default.ExchangeUri))
                    service.AutodiscoverUrl(Properties.Settings.Default.EXIMPUser, RedirectionUrlValidationCallback);
                else service.Url = new Uri(Properties.Settings.Default.ExchangeUri + "/ews/exchange.asmx");
                Terms terms = new Terms();

                List<Appointment> Appointments = new List<Appointment>();

                XmlDocument doc2 = new XmlDocument();
                doc2.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "roommapping.xml"));
                sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Loading Room Mappings");
                this.Progress.Current = 0;
                foreach (Room r in room)
                {
                    try
                    {
                        if (doc2.SelectSingleNode("/RoomMaps/Room[@name=\"" + r.Name + "\"]") != null)
                        {
                            sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Processing " + r.Name + " room events");
                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, doc2.SelectSingleNode("/RoomMaps/Room[@name=\"" + r.Name + "\"]").Attributes["email"].Value);


                            SearchFilter.SearchFilterCollection searchFilter = new SearchFilter.SearchFilterCollection();
                            searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, terms[0].StartDate));
                            searchFilter.Add(new SearchFilter.ContainsSubstring(AppointmentSchema.Subject, "Lesson:"));
                            ItemView view = new ItemView(999);
                            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.AppointmentType);
                            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Calendar, searchFilter, view);
                            foreach (Item item in findResults.Items)
                            {
                                Appointment appt = item as Appointment;
                                if (appt.AppointmentType == AppointmentType.RecurringMaster)
                                {
                                    sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Removing previous appointments for " + r.Name + " " + appt.Subject);
                                    appt.Delete(DeleteMode.HardDelete);
                                }
                            }
                            foreach (Lesson l in r)
                            {
                                sr.WriteLine(this.Progress.Value = DateTime.Now.ToString() + " Creating appointments for " + r.Name + " - " + l.Class + " on Day " + l.Day);
                                foreach (Term t in terms)
                                {
                                    DateTime hts = t.HalfTerm.Value.StartDate;
                                    if (hts.DayOfWeek == DayOfWeek.Monday) hts = hts.AddDays(-3);
                                    DateTime hte = t.HalfTerm.Value.EndDate;
                                    if (hte.DayOfWeek == DayOfWeek.Friday) hte = hte.AddDays(3);
                                    CreateApp(l, t.StartDate, hts, t.StartWeekNum == 2, service).Save();
                                    CreateApp(l, hte, t.EndDate, t.WeekNum(hte) == 2, service).Save();
                                }
                            }
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

        public Appointment CreateApp(Lesson l, DateTime start, DateTime end, bool week2, ExchangeService service)
        {
            Appointment appointment = new Appointment(service);
            appointment.Subject = appointment.Body = l.Class;
            appointment.Body += "<br />Auto Created: " + DateTime.Now.ToString();
            appointment.Subject = "Lesson: " + appointment.Subject;
            int plusday = l.Day;
            plusday--;
            if (week2)
            {
                if (plusday < 5) plusday += 5;
                else if (plusday > 4) plusday -= 5;
            }
            if (plusday > 4) plusday += 2;
            if (start.DayOfWeek == DayOfWeek.Tuesday) plusday--;
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
