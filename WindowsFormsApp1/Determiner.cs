using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    public class Determiner
    {
        /// <summary>
        /// run through each file in MLSFileName and check against AIMFileName files
        /// to determine whether the files in MLSFileName closed with hatco.
        /// 
        /// addressThreshold, addressThresholdWeak,
        /// ownerThreshold, and ownerThresholdWeak are percentages (between 0 and 1)
        /// </summary>
        /// <param name="MLSFileName"></param>
        /// <param name="AIMFileName"></param>
        /// <param name="addressThreshold"></param>
        /// <param name="ownerThreshold"></param>
        public void mainDeterminer (string MLSFileName, string AIMFileName, double addressThreshold,
            double addressThresholdWeak, double ownerThreshold, double ownerThresholdWeak)
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
                for (int currentMLSFile = 2; currentMLSFile <= rowCountMLS; currentMLSFile++)
                {
                    // initialize the variables that will determine if file in row closed
                    //
                    // For addressMatch and ownerMatch: 0 = no match,
                    // 1 = likely a match, 2 = definitely a match
                    bool dateClosedMatch = false;
                    int addressMatch = 0;
                    int ownerMatch = 0;
                    int ClosedGFNumRow = 2;

                    /// determine if date closed is a match
                    if (xlRangeMLS.Cells[currentMLSFile, MLSCloseDateCol].Value != null) // check that the next MLS close date cell is not empty
                    {
                        // parse MLS close date into a DateTime struct
                        string MLSRawCloseDate = xlRangeMLS.Cells[currentMLSFile, MLSCloseDateCol].Value.ToString();
                        DateTime MLSCloseDate = DateTime.Parse(MLSRawCloseDate);

                        for (int currentAIMFile = 2; currentAIMFile <= rowCountAIM; currentAIMFile++)
                        {
                            if (xlRangeAIM.Cells[currentAIMFile, AIMCloseDateCol].Value != null) // check that the next AIM close date cell is not empty
                            {
                                // parse AIM close date into a DateTime struct
                                string AIMRawCloseDate = xlRangeAIM.Cells[currentAIMFile, AIMCloseDateCol].Value.ToString();
                                DateTime AIMCloseDate = DateTime.Parse(AIMRawCloseDate);

                                if ((MLSCloseDate - AIMCloseDate).TotalDays <= 15) // check that the two close dates are within 15 days of each other
                                {
                                    dateClosedMatch = true;
                                    Console.WriteLine("Found match in days between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                }
                            }
                        }
                    }


                    /// determine if property addresses are a match only if date closed is already a match
                    if (dateClosedMatch && xlRangeMLS.Cells[currentMLSFile, MLSAddressCol].Value != null)
                    {
                        for (int currentAIMFile = 2; currentAIMFile <= rowCountAIM; currentAIMFile++)
                        {
                            if (xlRangeAIM.Cells[currentAIMFile, AIMAddressCol].Value != null)
                            {
                                string addressMLS = xlRangeMLS.Cells[currentMLSFile, MLSAddressCol].Value.ToString();
                                string addressAIM = xlRangeAIM.Cells[currentAIMFile, AIMAddressCol].Value.ToString();
                                int addressDistance = StringDistance.GetStringDistance(addressMLS, addressAIM); // get distance between the two strings

                                // check addressDistance against the percentage threshold of the longer test string to see if it is a match
                                if (addressDistance <= Math.Ceiling(addressThreshold * Math.Max(addressMLS.Length, addressAIM.Length)))
                                {
                                    addressMatch = 2;
                                    Console.WriteLine("Found match in address between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                }
                                else if (addressDistance <= Math.Ceiling(addressThresholdWeak * Math.Max(addressMLS.Length, addressAIM.Length)))
                                {
                                    addressMatch = 1;
                                    Console.WriteLine("Found likely match in address between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                }
                            }
                        }
                    }

                    /// determine if owner/seller name are a match only if date closed and addrees are already a match
                    if (dateClosedMatch && addressMatch > 0 && xlRangeMLS.Cells[currentMLSFile, MLSOwnerCol].Value != null)
                    {
                        for (int currentAIMFile = 2; currentAIMFile <= rowCountAIM; currentAIMFile++)
                        {
                            if (xlRangeAIM.Cells[currentAIMFile, AIMSellerCol].Value != null)
                            {
                                string owner = xlRangeMLS.Cells[currentMLSFile, MLSOwnerCol].Value.ToString();
                                string seller = xlRangeAIM.Cells[currentAIMFile, AIMSellerCol].Value.ToString();
                                int ownerDistance = StringDistance.GetStringDistance(owner, seller); // get distance between the two strings

                                // check ownerDistance against the percentage threshold of the longer test string to see if it is a match
                                if (ownerDistance <= Math.Ceiling(ownerThreshold * Math.Max(owner.Length, seller.Length)))
                                {
                                    ownerMatch = 2;
                                    ClosedGFNumRow = currentAIMFile;
                                    Console.WriteLine("Found match in owner/seller between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                    break; // if a match is found, there's no need to search any further
                                }
                                else if (ownerDistance <= Math.Ceiling(ownerThresholdWeak * Math.Max(owner.Length, seller.Length)))
                                {
                                    ownerMatch = 1;
                                    Console.WriteLine("Found likely match in owner/seller between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                }
                            }
                        }
                    }

                    /// determine whether the file was closed with hatco or not and print to xl file
                    if (dateClosedMatch && addressMatch == 2 && ownerMatch == 2)
                    {
                        string closedGF = xlRangeAIM.Cells[ClosedGFNumRow, AIMFileNoCol].Value.ToString();
                        xlRangeMLS.Cells[currentMLSFile, MLSGFCol].Value = closedGF;
                        Console.WriteLine("File on row " + currentMLSFile + " of MLS xl file closed with GF #"
                            + closedGF);
                    }
                    else if (dateClosedMatch && addressMatch > 0 && ownerMatch > 0)
                    {
                        xlRangeMLS.Cells[currentMLSFile, MLSGFCol].Value = "likely closed";
                        Console.WriteLine("File on row " + currentMLSFile + " likely closed");
                    }
                    else
                    {
                        xlRangeMLS.Cells[currentMLSFile, MLSGFCol].Value = "did not close";
                        Console.WriteLine("File on row " + currentMLSFile + " did not close");
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

                // save, close, and release workbooks
                xlWorkbookMLS.Save();
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
