using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // print data from text boxes to console for debugging purposes
            Console.WriteLine("MLS Input File Path: " + MLS_Input_File.Text);
            Console.WriteLine("AIM Input File Path: " + AIM_Input_File.Text);
            Console.WriteLine("Output File Path: " + Output_File.Text);

            // open all excel files for use
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbookMLS = null;
            Excel.Workbook xlWorkbookAIM = null;
            try
            {
                xlWorkbookMLS = xlApp.Workbooks.Open(MLS_Input_File.Text);
                xlWorkbookAIM = xlApp.Workbooks.Open(AIM_Input_File.Text);
            }
            catch (Exception ex) // catch possible "file could not open" exception
            {
                MessageBox.Show(ex.ToString());
            }

            if (xlWorkbookAIM != null && xlWorkbookMLS != null) // check that excel files opened properly
            {
                Excel._Worksheet xlWorksheetMLS = xlWorkbookMLS.Sheets[1];
                Excel._Worksheet xlWorksheetAIM = xlWorkbookAIM.Sheets[1];
                Excel.Range xlRangeMLS = xlWorksheetMLS.UsedRange;
                Excel.Range xlRangeAIM = xlWorksheetAIM.UsedRange;

                // get the range of rows and columns for each excel file opened
                int rowCountMLS = xlRangeMLS.Rows.Count;
                int colCountMLS = xlRangeMLS.Columns.Count;
                int rowCountAIM = xlRangeAIM.Rows.Count;
                int colCountAIM = xlRangeAIM.Columns.Count;

                // relevant columns indeces
                int MLSOwnerCol = 0, MLSAddressCol = 0, MLSCloseDateCol = 0, MLSGFCol = 0,
                    AIMFileNoCol = 0, AIMCloseDateCol = 0, AIMAddressCol = 0, AIMSellerCol = 0;

                // determine the columns in MLS file that have relevant information
                for (int i = 1; i <= colCountMLS; i++)
                {
                    if (xlRangeMLS.Cells[1, i].Value2 != null) // check that the cell is not empty
                    {
                        if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Owner"))
                            MLSOwnerCol = i;
                        else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Address"))
                            MLSAddressCol = i;
                        else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Close Date"))
                            MLSCloseDateCol = i;
                        else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("GF"))
                            MLSGFCol = i;
                    }
                }

                // determine the columns in AIM file that have relevant information
                for (int i = 1; i <= colCountAIM; i++)
                {
                    if (xlRangeAIM.Cells[1, i].Value2 != null) // check that the cell is not empty
                    {
                        if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("File Number"))
                            AIMFileNoCol = i;
                        else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Closing Date"))
                            AIMCloseDateCol = i;
                        else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Property Address"))
                            AIMAddressCol = i;
                        else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Seller"))
                            AIMSellerCol = i;
                    }
                }

                // print relevant columns indeces for debugging purposes
                Console.WriteLine("MLSOwnerCol: " + MLSOwnerCol.ToString() + ", "
                    + "MLSAddressCol: " + MLSAddressCol.ToString() + ", " + "MLSCloseDateCol: "
                    + MLSCloseDateCol.ToString() + ", " + "MLSGFCol: " + MLSGFCol.ToString() + "\n"
                    + "AimFileNoCol: " + AIMFileNoCol.ToString() + ", " + "AIMCloseDateCol: "
                    + AIMCloseDateCol.ToString() + ", " + "AIMAddressCol: " + AIMAddressCol.ToString()
                    + ", " + "AIMSellerCol: " + AIMSellerCol.ToString());



                // cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // release com objects so the excel processes are
                // fully killed from running in the background
                Marshal.ReleaseComObject(xlRangeMLS);
                Marshal.ReleaseComObject(xlRangeAIM);
                Marshal.ReleaseComObject(xlWorksheetMLS);
                Marshal.ReleaseComObject(xlWorksheetAIM);

                // close and release workbooks
                xlWorkbookMLS.Close();
                xlWorkbookAIM.Close();
                Marshal.ReleaseComObject(xlWorkbookMLS);
                Marshal.ReleaseComObject(xlWorkbookAIM);

                // quit and release excel app
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
