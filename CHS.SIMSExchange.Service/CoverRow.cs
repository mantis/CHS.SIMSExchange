using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace CHS.SIMSExchange.Service
{
    public class CoverRow
    {
        public static CoverRow Parse(XmlNode node)
        {
            CoverRow r = new CoverRow();
            r.Period = node.ChildNodes[0].InnerText;
            r.Time = node.ChildNodes[1].InnerText;
            r.Absent = node.ChildNodes[2].InnerText;
            r.Class = node.ChildNodes[4].InnerText;
            r.Room = node.ChildNodes[5].InnerText;
            r.OldRoom = node.ChildNodes[6].InnerText;
            if (node.ChildNodes[7].InnerText.Contains('$') && !node.ChildNodes[7].InnerText.Contains(','))
                r.Cover = node.ChildNodes[7].InnerText.Split(new char[] { ' ' })[0].Remove(0, 1) + "," + node.ChildNodes[7].InnerText.Remove(0, node.ChildNodes[7].InnerText.Split(new char[] { ' ' })[0].Length);
            else r.Cover = node.ChildNodes[7].InnerText;
            return r;
        }
        public override string ToString()
        {
            return Period + " " + Time + " " + Absent + " " + Class + " " + Room + " " + OldRoom + " " + Cover;
        }
        public string Period { get; set; }
        public string Absent { get; set; }
        public string Class { get; set; }
        public string Room { get; set; }
        public string OldRoom { get; set; }
        public string Cover { get; set; }
        public string Time { get; set; }
    }
}
