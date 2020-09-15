using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace CHS.SIMSExchange
{
    public struct HalfTerm
    {
        public HalfTerm(DateTime startDate, DateTime endDate)
        {
            this.startDate = startDate;
            this.endDate = endDate;
        }

        DateTime startDate, endDate;

        public DateTime EndDate
        {
            get { return endDate; }
            set { endDate = value; }
        }

        public DateTime StartDate
        {
            get { return startDate; }
            set { startDate = value; }
        }
    }

    public struct Term
    {
        public Term(string name, DateTime startDate, DateTime endDate, int startWeekNum, HalfTerm? halfTerm)
        {
            this.name = name;
            this.startDate = startDate;
            this.endDate = endDate;
            this.startWeekNum = startWeekNum;
            this.halfTerm = halfTerm;
        }

        HalfTerm? halfTerm;

        public HalfTerm? HalfTerm
        {
            get { return this.halfTerm; }
            set { this.halfTerm = value; }
        }

        int startWeekNum;

        public int StartWeekNum
        {
            get { return startWeekNum; }
            set { startWeekNum = value; }
        }

        string name;

        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        DateTime startDate, endDate;

        public DateTime EndDate
        {
            get { return endDate; }
            set { endDate = value; }
        }

        public DateTime StartDate
        {
            get { return startDate; }
            set { startDate = value; }
        }

        public int WeekNum(DateTime date)
        {
            if ((date >= this.startDate) && (date <= this.endDate))
            {
                System.Globalization.Calendar cal = CultureInfo.InvariantCulture.Calendar;
                int swn = this.startWeekNum;
                int x = cal.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                if (this.halfTerm.HasValue)
                {
                    if ((date >= this.startDate) && (date < this.halfTerm.Value.StartDate))
                    {
                        int y = cal.GetWeekOfYear(this.startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

                        int a = ((((x - y) % 2) == 0) ? this.startWeekNum : this.startWeekNum + 1);
                        if (a == 3) a = 1;
                        return a;
                    }
                    else if ((date > this.halfTerm.Value.EndDate) && (date <= this.EndDate))
                    {
                        int y = cal.GetWeekOfYear(this.startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                        int a = ((((x - y) % 2) == 0) ? this.startWeekNum : this.startWeekNum + 1);
                        //if (this.startWeekNum == 2) a = (a == 2 ? 1 : 2);
                        if (a > 2) a = 1;
                        a = a == 2 ? 1 : 2;
                        return a;
                    }
                    else return 0;
                }
                else if ((date >= this.startDate) && (date <= this.endDate))
                {
                    int y = cal.GetWeekOfYear(this.startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                    int a = ((((x - y) % 2) == 0) ? this.startWeekNum : this.startWeekNum + 1);
                    if (a == 3) a = 1;
                    return a;
                }
                else return 0;
            }
            else return -1;
        }
    }

    public class Terms : List<Term>
    {

        public Terms()
        {
            ReadTerms();
        }

        public static string isTerm(DateTime day)
        {
            foreach (Term term in new Terms())
                if (term.StartDate.Date <= day.Date && term.EndDate >= day.Date)
                {
                    if (term.HalfTerm.HasValue && term.HalfTerm.Value.StartDate <= day.Date && term.HalfTerm.Value.EndDate >= day.Date)
                        return "HalfTerm " + term.Name.Replace(" ", "");
                    else return "Term " + term.Name.Replace(" ", "");
                }
            return "invalid";
        }

        public static Term getTerm(DateTime day)
        {
            foreach (Term term in new Terms())
                if (term.StartDate.Date <= day.Date && term.EndDate >= day.Date)
                    return term;
            return new Term();
        }

        private XmlDocument TermsDoc
        {
            get
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "terms.xml"));
                return doc;
            }
        }

        public void ReadTerms()
        {
            XmlDocument doc = TermsDoc;

            foreach (XmlNode node in doc.SelectNodes("/Terms/Term"))
            {
                XmlNode halfTerm = node.SelectSingleNode("HalfTerm");
                string[] s, s2;
                HalfTerm? ht = null;
                if (halfTerm != null)
                {
                    s = halfTerm.Attributes["startDate"].Value.Split(new char[] { '/' });
                    s2 = halfTerm.Attributes["endDate"].Value.Split(new char[] { '/' });
                    ht = new HalfTerm(new DateTime(int.Parse(s[2]), int.Parse(s[1]), int.Parse(s[0])), new DateTime(int.Parse(s2[2]), int.Parse(s2[1]), int.Parse(s2[0])));
                }
                s = node.Attributes["startDate"].Value.Split(new char[] { '/' });
                s2 = node.Attributes["endDate"].Value.Split(new char[] { '/' });
                Add(new Term(node.Attributes["name"].Value, new DateTime(int.Parse(s[2]), int.Parse(s[1]), int.Parse(s[0])), new DateTime(int.Parse(s2[2]), int.Parse(s2[1]), int.Parse(s2[0])), int.Parse(node.Attributes["startWeekNum"].Value), ht));
            }
        }
    }
}
