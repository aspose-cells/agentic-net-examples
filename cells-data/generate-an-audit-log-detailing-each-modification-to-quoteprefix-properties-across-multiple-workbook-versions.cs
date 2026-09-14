// Title: Create an audit log of QuotePrefix style changes across Excel workbook revisions using Aspose.Cells for .NET
// AI Prompts: Generate two Excel workbooks with different Settings.QuotePrefixToStyle values, modify a quoted cell in each, and output a revision log that highlights cells where the QuotePrefix flag switched. | Write C# code that loads the latest workbook version, iterates through Worksheets.RevisionLogs, and prints the cell address, old and new QuotePrefix values, and revision metadata for each style change. | Enhance the program to write the QuotePrefix change details into a CSV report by leveraging Aspose.Cells' saving capabilities.
// Common Searches: asp.net how to audit QuotePrefix style changes with Aspose.Cells revision logs | c# retrieve cells where QuotePrefix flag changed between workbook versions | track QuotePrefixToStyle differences using Aspose.Cells shared workbook revisions | export QuotePrefix property audit to CSV with Aspose.Cells
// Tags: Aspose.Cells revision log audit QuotePrefix | C# detect QuotePrefix style change | shared workbook revision history Aspose.Cells | QuotePrefixToStyle property usage | export QuotePrefix audit to CSV

using System;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace QuotePrefixAudit
{
    // The example creates two workbook versions with opposite Settings.QuotePrefixToStyle settings, modifies a quoted cell in each version, saves revisions, then loads the latest file, walks through its revision logs, and prints detailed information for every cell revision where the QuotePrefix flag differs between the old and new styles.
    class Program
    {
        static void Main()
        {
            // Paths for the workbook versions
            string version1Path = "QuotePrefix_V1.xlsx";
            string version2Path = "QuotePrefix_V2.xlsx";

            // -------------------------------------------------
            // Create first version of the workbook
            // -------------------------------------------------
            Workbook wb1 = new Workbook();
            // Enable shared workbook to track revisions
            wb1.Settings.Shared = true;
            // Preserve revision history for 30 days
            wb1.Worksheets.RevisionLogs.DaysPreservingHistory = 30;
            // When true, strings starting with a single quote will have QuotePrefix style applied
            wb1.Settings.QuotePrefixToStyle = true;

            // Insert a string that starts with a single quote
            Cell cellA1_v1 = wb1.Worksheets[0].Cells["A1"];
            cellA1_v1.PutValue("'FirstVersion");
            // Save first version
            wb1.Save(version1Path);

            // -------------------------------------------------
            // Create second version of the workbook (modify QuotePrefix behavior)
            // -------------------------------------------------
            Workbook wb2 = new Workbook(version1Path);
            // Change the setting so QuotePrefix is NOT applied as a style
            wb2.Settings.QuotePrefixToStyle = false;

            // Modify the same cell with another quoted string
            Cell cellA1_v2 = wb2.Worksheets[0].Cells["A1"];
            cellA1_v2.PutValue("'SecondVersion");
            // Save second version (this creates a revision entry)
            wb2.Save(version2Path);

            // -------------------------------------------------
            // Load the latest version and generate audit log
            // -------------------------------------------------
            Workbook auditWb = new Workbook(version2Path);

            // Verify that the workbook contains revisions
            if (!auditWb.HasRevisions)
            {
                Console.WriteLine("No revisions found in the workbook.");
                return;
            }

            // Iterate through all revision logs
            RevisionLogCollection revisionLogs = auditWb.Worksheets.RevisionLogs;
            foreach (RevisionLog log in revisionLogs)
            {
                // Metadata provides information such as when the revision was saved
                RevisionHeader metadata = log.MetadataTable;
                Console.WriteLine($"Revision Log Saved Time: {metadata.SavedTime}");
                Console.WriteLine($"User: {metadata.UserName}");
                Console.WriteLine($"Number of revisions in this log: {log.Revisions.Count}");

                // Process each revision
                foreach (Revision rev in log.Revisions)
                {
                    // We are interested only in cell changes
                    if (rev is RevisionCellChange cellChange)
                    {
                        // Compare QuotePrefix flag between old and new styles
                        bool oldQuotePrefix = cellChange.OldStyle?.QuotePrefix ?? false;
                        bool newQuotePrefix = cellChange.NewStyle?.QuotePrefix ?? false;

                        if (oldQuotePrefix != newQuotePrefix)
                        {
                            Console.WriteLine("-------------------------------------------------");
                            Console.WriteLine($"Cell: {cellChange.CellName}");
                            Console.WriteLine($"Row: {cellChange.Row}, Column: {cellChange.Column}");
                            Console.WriteLine($"Old QuotePrefix: {oldQuotePrefix}");
                            Console.WriteLine($"New QuotePrefix: {newQuotePrefix}");
                            Console.WriteLine($"Old Value: {cellChange.OldValue}");
                            Console.WriteLine($"New Value: {cellChange.NewValue}");
                            Console.WriteLine($"Revision ID: {cellChange.Id}");
                            Console.WriteLine($"Revision Type: {cellChange.Type}");
                        }
                    }
                }
            }

            // Optional: Save a copy of the audited workbook (not required for the log)
            // auditWb.Save("AuditResult.xlsx");
        }
    }
}
