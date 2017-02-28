using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsio
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        FileHandler fh = new FileHandler();
        APL apl = new APL();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void convert(object sender, RoutedEventArgs e)
        {
            // convert the xls file to the desired format
            if (fh.importXLSFile(inXLSName.Text, PageNo.Text) && fh.importXMLFile(inXMLName.Text))
            {
                switch (fh.createRateEntryList()) {
                    case 0:
                        System.Windows.MessageBox.Show("Complete.");
                        inXLSName.Text = "";
                        inXMLName.Text = "";
                        break;
                    case 1:
                        System.Windows.MessageBox.Show("Commodity code in Excel form is empty.");
                        break;
                    case 2:
                        System.Windows.MessageBox.Show("Contract number in Excel form is empty.");
                        break;
                    case 3:
                        System.Windows.MessageBox.Show("Currency in Excel form is empty.");
                        break;
                    case 4:
                        System.Windows.MessageBox.Show("Start Date in Excel form is empty.");
                        break;
                    case 5:
                        System.Windows.MessageBox.Show("End date in Excel form is empty.");
                        break;
                }
                Console.WriteLine("finished making the list.");
                fh.exportXML();
            }
            
            
           
        }

        private void browse_in(object sender, RoutedEventArgs e)
        {
            // show an open dialog box to browse for the input file.

            OpenFileDialog dlg = new OpenFileDialog();
            var a = dlg.ShowDialog();

            if (a == System.Windows.Forms.DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName;
                inXLSName.Text = fileName;
            }


        }

        private void browse_in_xml(object sender, RoutedEventArgs e)
        {
            // show an open dialog box to browse for the input file.

            OpenFileDialog dlg = new OpenFileDialog();
            var a = dlg.ShowDialog();

            if (a == System.Windows.Forms.DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName;
                inXMLName.Text = fileName;
            }


        }

        private void browse_in_APL(object sender, RoutedEventArgs e)
        {
            // show an open dialog box to browse for the input file.

            OpenFileDialog dlg = new OpenFileDialog();
            var a = dlg.ShowDialog();

            if (a == System.Windows.Forms.DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName;
                xlsPathAPL.Text = fileName;
            }


        }

        private void convertAPL(object sender, RoutedEventArgs e)
        {
            // convert the xls file to the desired format
            apl.importXLSFile(xlsPathAPL.Text, inSheetAPL.Text, outSheetAPL.Text, OrigTAPL.Text, DestTAPL.Text);
            apl.createList();
            apl.writeToXLS();

            System.Windows.MessageBox.Show("Finished filling the table.");



        }
    }
    
}
