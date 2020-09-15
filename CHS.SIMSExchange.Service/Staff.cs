using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace CHS.SIMSExchange.Service
{
    public class Staff
    {
        public Staff() { Lessons = new List<Lesson>(); }
        private XmlNode node;
        public Staff(XmlNode node)
            : this()
        {
            this.node = node;
            this.FirstName = node.SelectSingleNode("Preferred_x0020_Forename") != null ? node.SelectSingleNode("Preferred_x0020_Forename").InnerText : node.SelectSingleNode("ChosenName") != null ? node.SelectSingleNode("ChosenName").InnerText : "";
            this.Surname = node.SelectSingleNode("Legal_x0020_Surname") != null ? node.SelectSingleNode("Legal_x0020_Surname").InnerText : node.SelectSingleNode("LegalSurname") != null ? node.SelectSingleNode("LegalSurname").InnerText : node.SelectSingleNode("Surname") != null ? node.SelectSingleNode("Surname").InnerText : "";
            this.Title = node.SelectSingleNode("Title").InnerText;
        }
        public override string ToString()
        {
            return node.OuterXml;
        }
        public string Title { get; set; }
        public string FirstName { get; set; }
        public string Surname { get; set; }
        public List<Lesson> Lessons { get; set; }
    }

    public class Lesson
    {
        public Lesson() { }
        public Lesson(XmlNode node)
        {
            int a;
            if (int.TryParse((node.SelectSingleNode("Day_x0020_name") == null ? node.SelectSingleNode("DayName") : node.SelectSingleNode("Day_x0020_name")).InnerText.Substring(0, 1), out a))
            {
                this.Day = _d.IndexOf((node.SelectSingleNode("Day_x0020_name") == null ? node.SelectSingleNode("DayName") : node.SelectSingleNode("Day_x0020_name")).InnerText.Remove(0, 1));
                this.Day++;
                this.Day *= a;
            }
            else this.Day = int.Parse((node.SelectSingleNode("Day_x0020_name") == null ? node.SelectSingleNode("DayName") : node.SelectSingleNode("Day_x0020_name")).InnerText.Remove(0, 3));
            if (this.Day == 0) this.Day = 10;
            if (node.SelectSingleNode("Room") != null) this.Room = node.SelectSingleNode("Room").InnerText;
            if (node.SelectSingleNode("Name") != null) this.Room = node.SelectSingleNode("Name").InnerText;
            this.Start = (node.SelectSingleNode("Start_x0020_Time") == null ? node.SelectSingleNode("StartTime") : node.SelectSingleNode("Start_x0020_Time")).InnerText;
            this.End = (node.SelectSingleNode("End_x0020_Time") == null ? node.SelectSingleNode("EndTime") : node.SelectSingleNode("End_x0020_Time")).InnerText;
            this.Class = (node.SelectSingleNode("Class") == null ? node.SelectSingleNode("ShortName") : node.SelectSingleNode("Class")).InnerText;
            int y;
            if (int.TryParse(this.Class.Substring(0, 1), out y)) { this.Year = y; if (this.Year == 1) this.Year = int.Parse(this.Class.Substring(0, 2)); }
            else this.Year = 0;
        }

        public int Day { get; set; }
        public string Room { get; set; }
        public string Start { get; set; }
        public string End { get; set; }
        public string Class { get; set; }
        public int Year { get; set; }
        public Staff Staff { get; set; }
        private List<string> _d = new List<string> { "Mon", "Tue", "Wed", "Thu", "Fri" };
    }
}
