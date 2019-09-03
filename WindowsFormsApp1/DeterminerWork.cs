using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace WindowsFormsApp1
{
    class DeterminerWork
    {
        private static object xlLock = new object();

        public void determinerThreadDoWork(object data)
        {
            DeterminerThread det = null; // get processing data

            lock (xlLock)
            {
                det = (DeterminerThread)data; // get processing data

                // get the range of the MLS file that this thread will work on
                int threadId = Thread.CurrentThread.ManagedThreadId % 4;
                int rowCountDivisor = (int)Math.Ceiling(det.progressNum / 4.0);
                if (threadId == 0)
                    det.rangeCount["rowCountMLSMin"] = 2;
                else
                    det.rangeCount["rowCountMLSMin"] = (rowCountDivisor * threadId) + 1;
                if (threadId == 3)
                    det.rangeCount["rowCountMLS"] = det.progressNum;
                else
                    det.rangeCount["rowCountMLS"] = (rowCountDivisor * (threadId + 1)) + 1;

                Console.WriteLine("Current thread: " + threadId.ToString());
                Console.WriteLine("rowCountMLS: " + det.rangeCount["rowCountMLS"].ToString());
                Console.WriteLine("rowCountMLSMin: " + det.rangeCount["rowCountMLSMin"].ToString());

                // process the given range of the MLS file
                determinerDoWork(det.xlRangeMLS, det.xlRangeAIM, det.rangeCount, det.relevantCols,
                    det.thresholds, det.progress, det.progressNum);
            }
        }

        /// <summary>
        /// Do the main work of mainDeterminer from Determiner class
        /// </summary>
        public void determinerDoWork(Excel.Range xlRangeMLS, Excel.Range xlRangeAIM, Dictionary<string, int> rangeCount,
            Dictionary<string, int> relevantCols, Dictionary<string, double> thresholds, IProgress<int> progress, int progressNum)
        {
            // loop through the files and do the main work
            for (int currentMLSFile = rangeCount["rowCountMLSMin"];
                currentMLSFile <= rangeCount["rowCountMLS"]; currentMLSFile++)
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
                if (xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSCloseDateCol"]].Value2 != null) // check that the next MLS close date cell is not empty
                {
                    // parse MLS close date into a DateTime struct
                    string MLSRawCloseDate = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSCloseDateCol"]].Value.ToString();
                    DateTime MLSCloseDate = DateTime.Parse(MLSRawCloseDate);

                    for (int currentAIMFile = 2; currentAIMFile <= rangeCount["rowCountAIM"]; currentAIMFile++)
                    {
                        if (xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMCloseDateCol"]].Value2 != null) // check that the next AIM close date cell is not empty
                        {
                            // parse AIM close date into a DateTime struct
                            string AIMRawCloseDate = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMCloseDateCol"]].Value.ToString();
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
                if (dateClosedMatch && xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSAddressCol"]].Value2 != null)
                {
                    for (int currentAIMFile = 2; currentAIMFile <= rangeCount["rowCountAIM"]; currentAIMFile++)
                    {
                        if (xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMAddressCol"]].Value2 != null)
                        {
                            string addressMLS = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSAddressCol"]].Value2.ToString();
                            string addressAIM = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMAddressCol"]].Value2.ToString();
                            int addressDistance = StringDistance.GetStringDistance(addressMLS, addressAIM); // get distance between the two strings

                            // check addressDistance against the percentage threshold of the longer test string to see if it is a match
                            if (addressDistance <= Math.Ceiling(thresholds["addressThreshold"] * Math.Max(addressMLS.Length, addressAIM.Length)))
                            {
                                addressMatch = 2;
                                Console.WriteLine("Found match in address between row " + currentMLSFile
                                    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                            else if (addressDistance <= Math.Ceiling(thresholds["addressThresholdWeak"] * Math.Max(addressMLS.Length, addressAIM.Length)))
                            {
                                addressMatch = 1;
                                Console.WriteLine("Found likely match in address between row " + currentMLSFile
                                    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                        }
                    }
                }

                /// determine if owner/seller name are a match only if date closed and addrees are already a match
                if (dateClosedMatch && addressMatch > 0 && xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSOwnerCol"]].Value2 != null)
                {
                    for (int currentAIMFile = 2; currentAIMFile <= rangeCount["rowCountAIM"]; currentAIMFile++)
                    {
                        if (xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMSellerCol"]].Value2 != null)
                        {
                            string owner = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSOwnerCol"]].Value2.ToString();
                            string seller = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMSellerCol"]].Value2.ToString();
                            int ownerDistance = StringDistance.GetStringDistance(owner, seller); // get distance between the two strings

                            // check ownerDistance against the percentage threshold of the longer test string to see if it is a match
                            if (ownerDistance <= Math.Ceiling(thresholds["ownerThreshold"] * Math.Max(owner.Length, seller.Length)))
                            {
                                ownerMatch = 2;
                                ClosedGFNumRow = currentAIMFile;
                                Console.WriteLine("Found match in owner/seller between row " + currentMLSFile
                                    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                break; // if a match is found, there's no need to search any further
                            }
                            else if (ownerDistance <= Math.Ceiling(thresholds["ownerThresholdWeak"] * Math.Max(owner.Length, seller.Length)))
                            {
                                ownerMatch = 1;
                                Console.WriteLine("Found likely match in owner/seller between row " + currentMLSFile
                                    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                        }
                    }
                }

                lock (xlLock)
                {
                    /// determine whether the file was closed with hatco or not and print to xl file
                    if (dateClosedMatch && addressMatch == 2 && ownerMatch == 2)
                    {
                        string closedGF = xlRangeAIM.Cells[ClosedGFNumRow, relevantCols["AIMFileNoCol"]].Value.ToString();
                        xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSGFCol"]].Value = closedGF;
                        Console.WriteLine("File on row " + currentMLSFile + " of MLS xl file closed with GF #"
                            + closedGF);
                    }
                    else if (dateClosedMatch && addressMatch > 0 && ownerMatch > 0)
                    {
                        xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSGFCol"]].Value = "likely closed";
                        Console.WriteLine("File on row " + currentMLSFile + " likely closed");
                    }
                    else
                    {
                        xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSGFCol"]].Value = "did not close";
                        Console.WriteLine("File on row " + currentMLSFile + " did not close");
                    }
                }

                // update progress bar after each row of MLS file
                if (progress != null)
                    progress.Report(100 / progressNum);
            }
        }
    }

    class DeterminerThread
    {
        public Excel.Range xlRangeMLS;
        public Excel.Range xlRangeAIM;
        public Dictionary<string, int> rangeCount;
        public Dictionary<string, int> relevantCols;
        public Dictionary<string, double> thresholds;
        public IProgress<int> progress;
        public int progressNum;

        public DeterminerThread(Excel.Range xlRangeMLS1, Excel.Range xlRangeAIM1, Dictionary<string, int> rangeCount1,
            Dictionary<string, int> relevantCols1, Dictionary<string, double> thresholds1, IProgress<int> progress1, int progressNum1)
        {
            xlRangeMLS = xlRangeMLS1;
            xlRangeAIM = xlRangeAIM1;
            rangeCount = rangeCount1;
            relevantCols = relevantCols1;
            thresholds = thresholds1;
            progress = progress1;
            progressNum = progressNum1;
        }
    }
}
