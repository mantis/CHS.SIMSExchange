using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Xml;

namespace CHS.SIMSExchange
{
    public class StaffThread
    {
        public void Start(object o)
        {
            string logpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "stafferror.log");
            if (File.Exists(logpath)) File.Delete(logpath);
            StreamWriter sr = new StreamWriter(logpath);
            try
            {
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
                staff = o as List<Staff>;
                if (Initialized != null) Initialized(staff.Count);

                List<Appointment> Appointments = new List<Appointment>();

                XmlDocument doc = new XmlDocument();
                doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "staffmapping.xml"));

                foreach (Staff s in staff)
                {
                    try
                    {
                        if (doc.SelectSingleNode("/staffmappings/staff[@first=\"" + s.FirstName + "\" and @last=\"" + s.Surname + "\"]") != null)
                        {
                            this.Current = "Removing previous appointments for " + s.Title + " " + s.FirstName + " " + s.Surname;
                            if (Updated != null) Updated();
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
                                this.Current = "Found Existing" + findResults.TotalCount + " out of " + findResults.Items.Count + " for " + s.Title + " " + s.FirstName + " " + s.Surname;
                                if (Updated != null) Updated();
                                var appsids = new List<ItemId>();
                                foreach (Item item in findResults.Items)
                                {
                                    Appointment appt = item as Appointment;
                                    if (appt.AppointmentType == AppointmentType.RecurringMaster)
                                        appsids.Add(appt.Id);
                                }
                                this.Current = "Removing previous appointments for " + s.Title + " " + s.FirstName + " " + s.Surname + " " + appsids.Count;
                                if (Updated != null) Updated();
                                try
                                {
                                    if (appsids.Count > 0)
                                        service.DeleteItems(appsids, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences, true);
                                }
                                catch { System.Threading.Thread.Sleep(1000); }
                                var c = service.FindItems(WellKnownFolderName.Calendar, searchFilter, view).TotalCount;
                                removecompleted = c == 0;
                                if (!removecompleted) removecompleted = c == 1;
                                if (!removecompleted)
                                {
                                    this.Current = "Remove not completed, still " + c + " to remove for " + s.Title + " " + s.FirstName + " " + s.Surname;
                                    if (Updated != null) Updated();
                                    System.Threading.Thread.Sleep(2000);
                                }
                            }
                            this.Current = "Creating appointments for " + s.Title + " " + s.FirstName + " " + s.Surname;
                            if (Updated != null) Updated();
                            var apps = new List<Appointment>();
                            foreach (Lesson l in s.Lessons)
                            {
                                foreach (Term t in terms)
                                {
                                    try
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
                                    catch (Exception ex1)
                                    {
                                        sr.WriteLine(ex1.Message);
                                        sr.WriteLine(ex1.Source);
                                        sr.WriteLine(ex1);
                                        sr.WriteLine(s.ToString());
                                        sr.Flush();
                                    }
                                }
                            }
                            this.Current = "Creating " + apps.Count + " appointments for " + s.Title + " " + s.FirstName + " " + s.Surname;
                            if (Updated != null) Updated();
                            service.CreateItems(apps, null, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
                        }
                    }
                    catch (Exception ex)
                    {
                        sr.WriteLine(ex.Message);
                        sr.WriteLine(ex.Source);
                        sr.WriteLine(ex);
                        sr.WriteLine(s.ToString());
                        sr.Flush();
                    }
                    this.Progress++;
                    if (Updated != null) Updated();
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

        public event ThreadUpdated Updated;
        public event ThreadUpdated Done;
        public event ThreadInitialized Initialized;
        public int Progress { get; private set; }
        public List<Staff> staff { get; private set; }
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
}
