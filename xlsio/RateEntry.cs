using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsio
{
    struct RateEntry
    {
        public string category;
        public string mode;
        public string startDate;
        public string endDate;
        public string contractNum;
        public string origin;
        public string destination;
        public string commodityCode;
        public string ContainerCode;
        public List<RateLine> rl;

        public RateEntry(string comm) {
            category = "";
            mode = "";
            startDate = "";
            endDate = "";
            contractNum = "";
            origin = "";
            destination = "";
            commodityCode = comm;
            ContainerCode = "";
            rl = new List<RateLine>();
        }
    }
}
