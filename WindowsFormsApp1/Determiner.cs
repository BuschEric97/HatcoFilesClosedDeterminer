using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
        /// 
        /// progress is used to update the windowsform's progress bar
        /// </summary>
        /// <param name="MLSFileName"></param>
        /// <param name="AIMFileName"></param>
        /// <param name="addressThreshold"></param>
        /// <param name="ownerThreshold"></param>
        public void mainDeterminer (string MLSFileName, string AIMFileName, double addressThreshold,
            double addressThresholdWeak, double ownerThreshold, double ownerThresholdWeak, IProgress<int> progress,
            mainForm form)
        {
            Application.UseWaitCursor = true; // set the cursor to waiting symbol

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

                Dictionary<string, int> rangeCount = new Dictionary<string, int>();
                Dictionary<string, int> relevantCols = new Dictionary<string, int>();
                Dictionary<string, double> thresholds = new Dictionary<string, double>();
                thresholds.Add("addressThreshold", addressThreshold);
                thresholds.Add("addressThresholdWeak", addressThresholdWeak);
                thresholds.Add("ownerThreshold", ownerThreshold);
                thresholds.Add("ownerThresholdWeak", ownerThresholdWeak);

                try
                {
                    // do the main processing on the excel files and catch any exceptions that are thrown
                    // get the range of rows and columns for AIM excel file
                    rangeCount.Add("rowCountMLS", xlRangeMLS.Rows.Count);
                    rangeCount.Add("colCountMLS", xlRangeMLS.Columns.Count);
                    rangeCount.Add("rowCountAIM", xlRangeAIM.Rows.Count);
                    rangeCount.Add("colCountAIM", xlRangeAIM.Columns.Count);

                    // relevant columns indeces
                    relevantCols.Add("MLSOwnerCol", 0);
                    relevantCols.Add("MLSAddressCol", 0);
                    relevantCols.Add("MLSCloseDateCol", 0);
                    relevantCols.Add("MLSGFCol", 0);
                    relevantCols.Add("AIMFileNoCol", 0);
                    relevantCols.Add("AIMCloseDateCol", 0);
                    relevantCols.Add("AIMAddressCol", 0);
                    relevantCols.Add("AIMSellerCol", 0);
                    relevantCols.Add("AIMEscrowCol", 0);

                    // determine the columns in MLS file that have relevant information
                    for (int i = 1; i <= rangeCount["colCountMLS"]; i++)
                    {
                        if (xlRangeMLS.Cells[1, i].Value2 != null) // check that the cell is not empty
                        {
                            if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Owner"))
                                relevantCols["MLSOwnerCol"] = i;
                            else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Address"))
                                relevantCols["MLSAddressCol"] = i;
                            else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("Close Date"))
                                relevantCols["MLSCloseDateCol"] = i;
                            else if (xlRangeMLS.Cells[1, i].Value2.ToString().Contains("GF"))
                                relevantCols["MLSGFCol"] = i;
                        }
                    }

                    // determine the columns in AIM file that have relevant information
                    for (int i = 1; i <= rangeCount["colCountAIM"]; i++)
                    {
                        if (xlRangeAIM.Cells[1, i].Value2 != null) // check that the cell is not empty
                        {
                            if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("File Number"))
                                relevantCols["AIMFileNoCol"] = i;
                            else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Date"))
                                relevantCols["AIMCloseDateCol"] = i;
                            else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Property Address"))
                                relevantCols["AIMAddressCol"] = i;
                            else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Seller"))
                                relevantCols["AIMSellerCol"] = i;
                            else if (xlRangeAIM.Cells[1, i].Value2.ToString().Contains("Escrow"))
                                relevantCols["AIMEscrowCol"] = i;
                        }
                    }

                    // add new columns to MLS file
                    relevantCols.Add("MLSLikelyCloseCol", rangeCount["colCountMLS"] + 1);
                    relevantCols.Add("MLSEscrowOfficerCol", rangeCount["colCountMLS"] + 2);
                    xlRangeMLS.Cells[1, relevantCols["MLSLikelyCloseCol"]].Value = "Likely Closed";
                    xlRangeMLS.Cells[1, relevantCols["MLSEscrowOfficerCol"]].Value = "Escrow Officer";

                    // set the progress bar to the first little tick
                    if (progress != null)
                        progress.Report(100 / rangeCount["rowCountMLS"]);

                    DeterminerWork det = new DeterminerWork();
                    det.determinerDoWork(xlWorksheetMLS, xlWorksheetAIM, xlRangeMLS, xlRangeAIM,
                        rangeCount, relevantCols, thresholds, progress, form);
                }
                catch (Exception ex) // if an exception is caught, close the excel files so they aren't held hostage
                {
                    Console.WriteLine("Problem with determiner processing. Closing excel files.");

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
                    xlWorkbookMLS.Close();
                    Console.WriteLine("closed MLS workbook");
                    xlWorkbookAIM.Close();
                    Console.WriteLine("closed AIM workbook");
                    Marshal.ReleaseComObject(xlWorkbookMLS);
                    Marshal.ReleaseComObject(xlWorkbookAIM);

                    // quit and release excel app
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    throw ex;
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
                Console.WriteLine("saved MLS workbook");
                xlWorkbookMLS.Close();
                Console.WriteLine("closed MLS workbook");
                xlWorkbookAIM.Close();
                Console.WriteLine("closed AIM workbook");
                Marshal.ReleaseComObject(xlWorkbookMLS);
                Marshal.ReleaseComObject(xlWorkbookAIM);

                // quit and release excel app
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            Application.UseWaitCursor = false; // set cursor back to default
        }
    }
}
