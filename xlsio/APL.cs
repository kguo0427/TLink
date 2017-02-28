using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsio
{
    class APL
    {
        private Excel.Application app = new Excel.Application();
        private Excel.Workbook inBook;
        private Excel.Worksheet sinSheet;
        private Excel.Worksheet soutSheet;
        private Excel.Worksheet sinDictSheet;
        private Excel.Worksheet soutDictSheet;
        private Excel.Range inSheet;
        private Excel.Range outSheet;
        private Excel.Range inDictSheet;
        private Excel.Range outDictSheet;
        private List<string> origList;
        private List<string> destList;
        private Dictionary<string, string> origDict;
        private Dictionary<string, string> destDict;
        private List<RawEntry> rawList;

        public bool importXLSFile(string name, string pin, string pout, string pdin, string pdout)
        {
            //open the original excel file
            int nin = 0;
            int nout = 0;
            int ndin = 0;
            int ndout = 0;
            int.TryParse(pin, out nin);
            int.TryParse(pout, out nout);
            int.TryParse(pdin, out ndin);
            int.TryParse(pdout, out ndout);
            if (name == "")
            {
                System.Windows.MessageBox.Show("Excel path is empty.");
                return false;

            }
            else if (!int.TryParse(pin, out nin) && nin > 0 && !int.TryParse(pout, out nout) && nout > 0 && !int.TryParse(pdin, out ndin) && ndin > 0 && !int.TryParse(pdout, out ndout) && ndout > 0)
            {
                System.Windows.MessageBox.Show("Page number is invalid.");
                return false;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(pout);
                this.inBook = app.Workbooks.Open(@name);

                this.sinDictSheet = this.inBook.Sheets[ndin];
                this.soutDictSheet = this.inBook.Sheets[ndout];
                this.sinSheet = this.inBook.Sheets[nin];
                this.soutSheet = this.inBook.Sheets[nout];
                this.inDictSheet = this.sinDictSheet.UsedRange;
                this.outDictSheet = this.soutDictSheet.UsedRange;
                this.inSheet = this.sinSheet.UsedRange;
                this.outSheet = this.soutSheet.UsedRange;
                return true;
            }

        }

        public string getCell(Excel.Range r, int x, int y)
        {
            if (r.Cells[x, y].Value != null)
            {
                return r.Cells[x, y].Value2.ToString();

            }
            else
            {
                return "";
            }
        }

        public void setCell(Excel.Range r, int x, int y, string s)
        {
            r.Cells[x, y].Value2 = s;
        }

        public int createList() {
            origList = new List<string>();
            destList = new List<string>();
            rawList = new List<RawEntry>();
            origDict = new Dictionary<string, string>();
            destDict = new Dictionary<string, string>();

            //create dictionary for the destination cities
            int ind = 2;
            while (getCell(outDictSheet, ind, 1) != "")
            {
                destDict.Add(getCell(outDictSheet, ind, 1).ToUpper(), getCell(outDictSheet, ind, 2));
                ind++;
            }

            //create dictionary for the origin cities
            ind = 2;
            while (getCell(inDictSheet, ind, 1) != "") {
                
                origDict.Add(getCell(inDictSheet, ind, 1).ToUpper(), getCell(inDictSheet, ind, 2));
                ind++;
            }

            

            //create the list for the lines and along with the list for origins and destinations
            ind = 3;
            string origCode = "";
            string destCode = "";
            string priceFourty = "";
            string priceFFive = "";
            while (getCell(inSheet, ind, 1) != "") {
                if (origDict.ContainsKey(getCell(inSheet, ind, 1).ToUpper()))
                {
                    origCode = origDict[getCell(inSheet, ind, 1).ToUpper()];
                }
                else {
                    System.Windows.MessageBox.Show("There is no UNLOCO for " + getCell(inSheet, ind, 1).ToUpper());
                }

                if (destDict.ContainsKey(getCell(inSheet, ind, 2).ToUpper()))
                {
                    destCode = destDict[getCell(inSheet, ind, 2).ToUpper()];
                }
                else
                {
                    System.Windows.MessageBox.Show("There is no UNLOCO for " + getCell(inSheet, ind, 2).ToUpper());
                }
                
                //make sure both origin and destination has the proper code
                if (origCode != "N/A" && destCode != "N/A") {
                    priceFourty = getCell(inSheet, ind, 8);
                    priceFFive = getCell(inSheet, ind, 9);
                    //update the one with 40
                    if (priceFourty != "" && priceFourty != "N/A" && priceFourty != "DELETE") {
                        //add to the origin list if hasn't already
                        if (!origList.Contains(origCode)) {
                            origList.Add(origCode);
                        }
                        //add to the destination list if hasn't already
                        if (!destList.Contains(destCode))
                        {
                            destList.Add(destCode);
                        }
                        //add the entry to the raw list
                        rawList.Add(new RawEntry(origCode, destCode, priceFourty, 40));
                    }
                    //update the one with 45
                    if (priceFFive != "" && priceFFive != "N/A" && priceFFive != "DELETE")
                    {
                        //add to the origin list if hasn't already
                        if (!origList.Contains(origCode))
                        {
                            origList.Add(origCode);
                        }
                        //add to the destination list if hasn't already
                        if (!destList.Contains(destCode))
                        {
                            destList.Add(destCode);
                        }
                        //add the entry to the raw list
                        rawList.Add(new RawEntry(origCode, destCode, priceFFive, 45));
                        
                    }

                }
                ind++;
            }

            return 0;
        }

        public int writeToXLS() {
            // writing down all the origin ports and the 2 container sizes
            int indOrig = 1;
            
            while (indOrig <= origList.Count) {
                setCell(outSheet, 5, 2 * indOrig, origList[indOrig - 1]);
                setCell(outSheet, 6, 2 * indOrig, "40HC");
                setCell(outSheet, 6, 2 * indOrig + 1, "45HC");
                indOrig++;
            }

            //writing down all the destination ports
            int indDest = 1;
            while (indDest <= destList.Count)
            {
                setCell(outSheet, indDest + 6, 1, destList[indDest - 1]);
                indDest++;
            }
            //filling the blanks in the table
            int indT = 0;
            int oi = 0;
            int di = 0;
            int flag = 0;
            RawEntry temp;
            while (indT < rawList.Count) {

                temp = rawList[indT];
                
                //find index for origin
                flag = 0;
                oi = 0;
                while (flag == 0 && oi < origList.Count) {
                    if (origList[oi].Equals(temp.origin)) {
                        flag = 1;
                    }
                    oi++;
                }
                //find index for destination
                flag = 0;
                di = 0;
                while (flag == 0 && di < destList.Count)
                {
                    if (destList[di].Equals(temp.destination))
                    {
                        flag = 1;
                    }
                    di++;
                }
                //set the price based on the container type
                if (temp.size == 40) {
                    setCell(outSheet, di + 6, (oi) * 2, (int.Parse(temp.price) + int.Parse(getCell(outSheet, 2, 2))).ToString());
                } else if (temp.size == 45)
                {
                    setCell(outSheet, di + 6, (oi) * 2 + 1, (int.Parse(temp.price) + int.Parse(getCell(outSheet, 2, 3))).ToString());
                }

                indT++;

            }
            inBook.Save();
            inBook.Close();
            return 0;
        }
    }
}
