﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    class Determiner
    {
        public void mainDeterminer (string MLSFileName, string AIMFileName)
        {
            // open all excel files for use
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbookMLS = null;
            Excel.Workbook xlWorkbookAIM = null;
            try
            {
                xlWorkbookMLS = xlApp.Workbooks.Open(MLSFileName);
                xlWorkbookAIM = xlApp.Workbooks.Open(AIMFileName);
            }
            catch (Exception ex) // catch possible "file could not open" exception
            {
                throw ex;
            }

            if (xlWorkbookAIM != null && xlWorkbookMLS != null) // check that excel files opened properly
            {
                // open worksheets and range in excel files for use
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

                // loop through the files and do the main work
                for (int i = 2; i <= rowCountMLS; i++)
                {
                    if (xlRangeMLS.Cells[i, MLSCloseDateCol].Value != null) // check that the next MLS close date cell is not empty
                    {
                        // parse MLS close date into a DateTime struct
                        string MLSRawCloseDate = xlRangeMLS.Cells[i, MLSCloseDateCol].Value.ToString();
                        DateTime MLSCloseDate = DateTime.Parse(MLSRawCloseDate);

                        for (int j = 2; j <= rowCountAIM; j++)
                        {
                            if (xlRangeAIM.Cells[j, AIMCloseDateCol].Value != null) // check that the next AIM close date cell is not empty
                            {
                                // parse AIM close date into a DateTime struct
                                string AIMRawCloseDate = xlRangeAIM.Cells[j, AIMCloseDateCol].Value.ToString();
                                DateTime AIMCloseDate = DateTime.Parse(AIMRawCloseDate);

                                if ((MLSCloseDate - AIMCloseDate).TotalDays <= 15) // check that the two close dates are within 15 days of each other
                                {
                                    Console.WriteLine("Found match in days between row " + i
                                        + " in MLS file and row " + j + " in AIM file");
                                }
                            }
                        }
                    }
                }

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