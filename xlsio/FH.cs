using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using xml = System.Xml;
using linq = System.Xml.Linq;

namespace xlsio
{
    class FileHandler
    {
        private Excel.Application app = new Excel.Application();
        private Excel.Workbook inBook;
        private Excel.Worksheet inSheet;
        private Excel.Range xlsIn;
        private List<RateEntry> reList;
        private string xmlPath;
        public bool importXLSFile(string name, string p)
        {
            //open the original excel file
            int n = 0;
            if (name == "")
            {
                System.Windows.MessageBox.Show("Excel path is empty.");
                return false;

            }
            else if (!int.TryParse(p, out n) && n > 0)
            {
                System.Windows.MessageBox.Show("Page number is invalid.");
                return false;
            }
            else {
                this.inBook = app.Workbooks.Open(@name);
                this.inSheet = this.inBook.Sheets[n];
                this.xlsIn = this.inSheet.UsedRange;
                return true;
            }
            
        }

        public string getCell(int x, int y) {
            if (this.xlsIn.Cells[x, y].Value != null)
            {
                return this.xlsIn.Cells[x, y].Value2.ToString();
            }
            else {
                return "";
            }
        }

        public string updateTime(string t) {
            string[] dt = t.Split(' ');
            string[] date = dt[0].Split('/');
            for (int i = 0; i < date.Length; i++) {
                if (date[i].Length == 1) {
                    date[i] = "0" + date[i];
                }
            }
            Console.WriteLine(dt[0]);
            string dateS = date[2] + "-" + date[0] + "-" + date[1];
            string f = dateS + "T" + dt[1];
            return f;
        }
        public int createRateEntryList() {
            //create the list of all rate entries
            reList = new List<RateEntry>();
            string contractNo;
            string currency;
            string startDate;
            string endDate;
            string commCode;
            if (getCell(1, 2) != "")
            {
                commCode = getCell(1, 2);
            } else
            {
                return 1;
            }
            if (getCell(2, 2) != "")
            {
                contractNo = getCell(2, 2);
            }
            else
            {
                return 2;
            }
            if (getCell(3, 2) != "")
            {
                currency = getCell(3, 2);
            }
            else
            {
                return 3;
            }
            if (getCell(4, 2) != "")
            {
                startDate = updateTime(this.xlsIn.Cells[4, 2].Value.ToString());
            }
            else
            {
                return 4;
            }
            if (getCell(5, 2) != "")
            {
                endDate = updateTime(this.xlsIn.Cells[5, 2].Value.ToString());
            }
            else
            {
                return 5;
            }
            //setting the initial numbers and indices for the loop
            const int ORIGrow = 7;
            const int CONTrow = 8;
            const int DESTcol = 1;
            string origTemp = "";
            int indDest = 9;
            int indCon = 3;

            while (getCell(CONTrow, indCon) != "") {
                
                while (getCell(indDest, DESTcol) != "") {
                    //instantiate a new copy of the rate entry
                    //setting all the same stuff
                    RateEntry tmp = new RateEntry(commCode);
                    tmp.contractNum = contractNo;
                    tmp.startDate = startDate;
                    tmp.endDate = endDate;
                    tmp.category = "FCL";
                    tmp.mode = "SEA";
                    RateLine r = new RateLine();
                    r.chargeCode = "FRT";
                    r.currency = currency;
                    r.note = getCell(indDest, 2);
                    //updating the current origin and container type
                    if (getCell(ORIGrow, indCon) != "")
                    {
                        origTemp = getCell(ORIGrow, indCon);
                    }
                    tmp.origin = origTemp;
                    tmp.ContainerCode = getCell(CONTrow, indCon);
                    //loop through the destinations and update each destination and price.
                    tmp.destination = getCell(indDest, DESTcol);
                    if (getCell(indDest, indCon) != "")
                    {
                        //checking if the price exists
                        //if it does, add it to the list
                        r.price = getCell(indDest, indCon);
                        tmp.rl.Add(r);
                        reList.Add(tmp);
                    }
                    indDest++;
                }
                //reset the index for destination back to the start at 9
                indDest = 9;
                indCon++;

            }
            //close the book
            inBook.Close();
            return 0;
        }

        public void exportXML() {
            //output to the xml file
            string path = this.xmlPath;
            linq.XNamespace ns = "http://www.edi.com.au/EnterpriseService/";

            linq.XDocument doc = linq.XDocument.Load(path);
            
            linq.XElement rateEntries = new linq.XElement(ns + "RateEntries");
            foreach (RateEntry r in reList) {
                linq.XElement entry = new linq.XElement(ns + "RateEntry");
                if (r.category != "")
                {
                    entry.Add(new linq.XElement(ns + "Category", r.category));
                }
                entry.Add(new linq.XElement(ns + "Mode", "SEA"));
                if (r.startDate != "")
                {
                    entry.Add(new linq.XElement(ns + "StartDate", r.startDate));
                }
                if (r.endDate != "")
                {
                    entry.Add(new linq.XElement(ns + "EndDate", r.endDate));
                }
                if (r.ContainerCode != "")
                {
                    linq.XElement ct = new linq.XElement(ns + "ContainerType");
                    ct.Add(new linq.XElement(ns + "ContainerCode", r.ContainerCode));
                    entry.Add(ct);
                }
                if (r.contractNum != "")
                {
                    entry.Add(new linq.XElement(ns + "ContractNumber", r.contractNum));
                }
                if (r.origin != "")
                {
                    entry.Add(new linq.XElement(ns + "Origin", r.origin));
                }
                if (r.destination != "")
                {
                    entry.Add(new linq.XElement(ns + "Destination", r.destination));
                }
                if (r.commodityCode != "")
                {
                    entry.Add(new linq.XElement(ns + "CommodityCode", r.commodityCode));
                }
                linq.XElement rls = new linq.XElement(ns + "RateLines");
                foreach (RateLine rl in r.rl) {
                    linq.XElement rateline = new linq.XElement(ns + "RateLine");
                    rateline.Add(new linq.XElement(ns + "Description", "Ocean Freight"));
                    rateline.Add(new linq.XElement(ns + "Currency", "USD"));
                    rateline.Add(new linq.XElement(ns + "ChargeCode", "FRT"));
                    linq.XElement rc = new linq.XElement(ns + "RateCalculator");
                    linq.XElement unt = new linq.XElement(ns + "UNTCalculator");
                    unt.Add(new linq.XElement(ns + "PerUnitPrice", rl.price));
                    linq.XElement notes = new linq.XElement(ns + "Notes");
                    linq.XElement note = new linq.XElement(ns + "Note");
                    note.Add(new linq.XElement(ns + "NoteType", "TradeLaneChargeInformation"));
                    note.Add(new linq.XElement(ns + "NoteData", rl.note));
                    rc.Add(unt);
                    notes.Add(note);
                    rateline.Add(rc);
                    rateline.Add(notes);
                    rls.Add(rateline);
                }
                entry.Add(rls);
                rateEntries.Add(entry);
            }
            var ele = doc.Descendants(ns + "Rate").FirstOrDefault();
            if (doc.Descendants(ns + "RateEntries").Any())
            {
                doc.Descendants(ns + "RateEntries").FirstOrDefault().Remove();
            }
            ele.Add(rateEntries);
            doc.Save(path);
        }

        public bool importXMLFile(string name)
        {
            //open the original xml file
            if (name != "")
            {
                xmlPath = name;
                return true;
            } else
            {
                System.Windows.MessageBox.Show("XML path is empty.");
                return false;
            }
        }
    }

}
