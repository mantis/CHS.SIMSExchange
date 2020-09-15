using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CHS.SIMSExchange.Service
{
    public class Progress
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public string Value { get; set; }
        public bool Finished { get; set; }
    }
}
