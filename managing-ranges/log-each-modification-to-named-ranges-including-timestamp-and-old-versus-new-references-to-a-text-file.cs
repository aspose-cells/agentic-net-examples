using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace NamedRangeChangeLogger
{
    class Program
    {
        static void Main()
        {
            // Paths for the workbook and the log file
            string workbookPath = "NamedRangeDemo.xlsx";
            string logPath = "NamedRangeChangesLog.txt";

            // -----------------------------------------------------------------
            // Step 1: Create a workbook, enable shared mode to track revisions
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            wb.Settings.Shared = true; // Enable revision tracking

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "Sheet1";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Apple");
            sheet.Cells["A2"].PutValue("Banana");
            sheet.Cells["A3"].PutValue("Cherry");

            // Create a named range "Fruits"
            int nameIndex = wb.Worksheets.Names.Add("Fruits");
            Name fruitName = wb.Worksheets.Names[nameIndex];
            fruitName.RefersTo = "=Sheet1!$A$1:$A$3";

            // Save the initial version (revision 0)
            wb.Save(workbookPath);

            // -----------------------------------------------------------------
            // Step 2: Modify the named range to generate a revision
            // -----------------------------------------------------------------
            // Change the reference to include an additional cell
            fruitName.RefersTo = "=Sheet1!$A$1:$A$4";

            // Save again – this creates a revision entry for the defined name change
            wb.Save(workbookPath);

            // -----------------------------------------------------------------
            // Step 3: Log all named‑range modifications to a text file
            // -----------------------------------------------------------------
            LogNamedRangeChanges(workbookPath, logPath);

            Console.WriteLine("Logging completed. Check the file: " + logPath);
        }

        /// <summary>
        /// Reads revision logs from the specified workbook and writes
        /// each named‑range change (timestamp, name, old formula, new formula)
        /// to the given log file.
        /// </summary>
        static void LogNamedRangeChanges(string workbookFile, string logFile)
        {
            // Load the workbook (must be the same file that was saved in shared mode)
            Workbook wb = new Workbook(workbookFile);

            // Open the log file for appending
            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                // Iterate through all revision logs in the workbook
                foreach (RevisionLog log in wb.Worksheets.RevisionLogs)
                {
                    // Each log may contain multiple revisions
                    foreach (Revision rev in log.Revisions)
                    {
                        // We are interested only in defined‑name revisions
                        if (rev.Type == RevisionType.DefinedName && rev is RevisionDefinedName nameRev)
                        {
                            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            string nameText = nameRev.Text;               // The name of the defined range
                            string oldFormula = nameRev.OldFormula ?? "N/A";
                            string newFormula = nameRev.NewFormula ?? "N/A";

                            // Write a formatted line to the log file
                            writer.WriteLine($"{timeStamp} | Name: {nameText} | Old: {oldFormula} | New: {newFormula}");
                        }
                    }
                }
            }
        }
    }
}