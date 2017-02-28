using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsio
{
    class RawEntry
    {
        public string origin;
        public string destination;
        public string price;
        public int size;

        public RawEntry(string o, string d, string p, int s) {
            origin = o;
            destination = d;
            price = p;
            size = s;
        }
    }
}
