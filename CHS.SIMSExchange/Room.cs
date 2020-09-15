using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CHS.SIMSExchange
{
    public class Room:List<Lesson>
    {
        public string Name { get; set; }
        public Room(string Name): base() { this.Name = Name; }
    }

    public class o1
    {
        public List<Room> Rooms;
        public List<Staff> Staff;
    }
}
