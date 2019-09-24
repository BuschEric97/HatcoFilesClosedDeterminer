using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using USAddress;

namespace WindowsFormsApp1
{
    class DeterminerWork
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="xlWorksheetMLS"></param>
        /// <param name="xlWorksheetAIM"></param>
        /// <param name="xlRangeMLS"></param>
        /// <param name="xlRangeAIM"></param>
        /// <param name="rangeCount"></param>
        /// <param name="relevantCols"></param>
        /// <param name="thresholds"></param>
        /// <param name="progress"></param>
        public void determinerDoWork(Excel._Worksheet xlWorksheetMLS, Excel._Worksheet xlWorksheetAIM,
            Excel.Range xlRangeMLS, Excel.Range xlRangeAIM, Dictionary<string, int> rangeCount,
            Dictionary<string, int> relevantCols, Dictionary<string, double> thresholds, IProgress<int> progress)
        {
            // loop through the files and do the main work
            for (int currentMLSFile = 2; currentMLSFile <= rangeCount["rowCountMLS"]; currentMLSFile++)
            {
                // initialize the variables that will determine if file in row closed
                //
                // For addressMatch and ownerMatch: 0 = no match,
                // 1 = likely a match, 2 = definitely a match
                bool dateClosedMatch = false;
                int addressMatch = 0;
                int ownerMatch = 0;
                int closedGFNumRow = 2;
                List<int> consideredRowsAIM = new List<int>();

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
                                consideredRowsAIM.Add(currentAIMFile); // add the current AIM file row to the list of considered AIM rows
                                dateClosedMatch = true;
                                //Console.WriteLine("Found match in days between row " + currentMLSFile
                                //    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                        }
                    }
                }

                /// determine if owner/seller name are a match only if date closed and addrees are already a match
                if (dateClosedMatch && xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSOwnerCol"]].Value2 != null)
                {
                    for (int currentAIMFile = 2; currentAIMFile <= rangeCount["rowCountAIM"]; currentAIMFile++)
                    {
                        if (xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMSellerCol"]].Value2 != null &&
                            consideredRowsAIM.Contains(currentAIMFile))
                        {
                            string owner = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSOwnerCol"]].Value2.ToString();
                            string seller = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMSellerCol"]].Value2.ToString();
                            int ownerDistance = StringDistance.GetStringDistance(owner, seller); // get distance between the two strings

                            // compute updated thresholds
                            int ownerThresholdUpdated = (int)Math.Ceiling(thresholds["ownerThreshold"] * Math.Max(owner.Length, seller.Length));
                            int ownerThresholdWeakUpdated = (int)Math.Ceiling(thresholds["ownerThresholdWeak"] * Math.Max(owner.Length, seller.Length));

                            // check ownerDistance against the percentage threshold of the longer test string to see if it is a match
                            if (ownerDistance <= ownerThresholdUpdated)
                            {
                                ownerMatch = 2;
                                //Console.WriteLine("Found match in owner/seller between row " + currentMLSFile
                                //    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                            else if (ownerDistance <= ownerThresholdWeakUpdated && ownerDistance > ownerThresholdUpdated)
                            {
                                ownerMatch = 1;
                                //Console.WriteLine("Found likely match in owner/seller between row " + currentMLSFile
                                //    + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                            }
                            else
                                consideredRowsAIM.Remove(currentAIMFile);
                        }
                    }
                }

                /// determine if property addresses are a match only if date closed is already a match
                if (dateClosedMatch && ownerMatch > 0 && xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSAddressCol"]].Value2 != null)
                {
                    for (int currentAIMFile = 2; currentAIMFile <= rangeCount["rowCountAIM"]; currentAIMFile++)
                    {
                        if (xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMAddressCol"]].Value2 != null &&
                            consideredRowsAIM.Contains(currentAIMFile))
                        {
                            // get the address strings from the xl files and parse them by the space character
                            string addressMLS = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSAddressCol"]].Value2.ToString();
                            string addressAIM = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMAddressCol"]].Value2.ToString();
                            string[] parsedAddressMLS = addressMLS.Split(' ');
                            string[] parsedAddressAIM = addressAIM.Split(' ');

                            if (parsedAddressMLS[0] == parsedAddressAIM[0]) // check that the address numbers match
                            {
                                int addressDistance = StringDistance.GetStringDistance(addressMLS, addressAIM); // get distance between the two strings

                                // compute updated thresholds
                                int addressThresholdUpdated = (int)Math.Ceiling(thresholds["addressThreshold"] * Math.Max(addressMLS.Length, addressAIM.Length));
                                int addressThresholdWeakUpdated = (int)Math.Ceiling(thresholds["addressThresholdWeak"] * Math.Max(addressMLS.Length, addressAIM.Length));

                                // check addressDistance against the percentage threshold of the longer test string to see if it is a match
                                if (addressDistance <= addressThresholdUpdated)
                                {
                                    addressMatch = 2;
                                    closedGFNumRow = currentAIMFile;
                                    Console.WriteLine("Found match in address between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                    break; // if a match is found, there's no need to search any further
                                }
                                else if (addressDistance <= addressThresholdWeakUpdated && addressDistance > addressThresholdUpdated)
                                {
                                    addressMatch = 1;
                                    Console.WriteLine("Found likely match in address between row " + currentMLSFile
                                        + " in MLS xl file and row " + currentAIMFile + " in AIM xl file");
                                }
                                else
                                    consideredRowsAIM.Remove(currentAIMFile);
                            }
                        }
                    }
                }

                /// go through the remaining files in consideredRowsAIM and determine which one is the best match
                for (int currentAIMFile = 2; currentAIMFile < consideredRowsAIM.Count; currentAIMFile++)
                {
                    string addressMLS = xlRangeMLS.Cells[currentMLSFile, relevantCols["MLSAddressCol"]].Value2.ToString();
                    string addressAIMNew = xlRangeAIM.Cells[currentAIMFile, relevantCols["AIMAddressCol"]].Value2.ToString();
                    string addressAIMPrev = xlRangeAIM.Cells[closedGFNumRow, relevantCols["AIMAddressCol"]].Value2.ToString();
                    int addressDistanceNew = StringDistance.GetStringDistance(addressMLS, addressAIMNew);
                    int addressDistancePrev = StringDistance.GetStringDistance(addressMLS, addressAIMPrev);

                    if (addressDistanceNew < addressDistancePrev)
                        closedGFNumRow = currentAIMFile;
                }

                /// determine whether the file was closed with hatco or not and print to xl file
                if (dateClosedMatch && addressMatch == 2 && ownerMatch == 2)
                {
                    string closedGF = xlRangeAIM.Cells[closedGFNumRow, relevantCols["AIMFileNoCol"]].Value.ToString();
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

                // update progress bar after each row of MLS file
                if (progress != null)
                    progress.Report(100 / rangeCount["rowCountMLS"]);
            }
        }
    }
}
