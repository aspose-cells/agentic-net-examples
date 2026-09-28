// Title: Verify a changed Excel cell formula by reading Aspose.Cells revision logs in C#
// AI Prompts: Write C# code that creates an Excel workbook, sets a formula in a cell, saves it, modifies the formula, saves again, then loads the file and uses Aspose.Cells RevisionLogCollection to extract the old and new formula strings for that cell. | Show how to iterate through RevisionLog entries in Aspose.Cells to locate a RevisionCellChange for a specific cell and compare its OldFormula and NewFormula values against expected expressions. | Provide a C# example that validates that a formula update (e.g., from =SUM(B1:B3) to =AVERAGE(B1:B3)) was recorded correctly in the workbook's revision log.
// Common Searches: aspocells c# read revision log old formula | compare original and updated formula in Excel using Aspose.Cells | how to detect formula changes in a workbook with Aspose.Cells revision tracking | C# verify cell A1 formula modification via revision logs | Aspose.Cells retrieve formula change history programmatically
// Tags: Aspose.Cells revision log API | C# compare old and new cell formula | track Excel formula changes Aspose.Cells | retrieve RevisionCellChange formula text | validate formula modification programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace AsposeCellsFormulaComparison
{
    // The example creates a workbook, writes a SUM formula in A1, saves it, changes the formula to AVERAGE, saves again, then reloads the file, iterates the RevisionLogCollection to find the A1 cell change, extracts OldFormula and NewFormula, and confirms the formula change was recorded.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path for the workbook that will store revisions
                string filePath = "RevisionFormulaComparison.xlsx";

                // -------------------------------------------------
                // 1. Create a new workbook and set an initial formula
                // -------------------------------------------------
                Workbook workbook = new Workbook();

                // Enable revision tracking if supported (commented out for compatibility)
                // workbook.Settings.RevisionLogEnabled = true;
                // workbook.Settings.TrackChangesEnabled = true;

                Worksheet sheet = workbook.Worksheets[0];

                // Initial formula in cell A1
                sheet.Cells["A1"].Formula = "=SUM(B1:B3)";

                // Populate referenced cells so the formula is valid
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Save the workbook – this creates the first revision state
                workbook.Save(filePath);

                // -------------------------------------------------
                // 2. Modify the formula to generate a revision entry
                // -------------------------------------------------
                sheet.Cells["A1"].Formula = "=AVERAGE(B1:B3)";

                // Save again – the change is recorded in the revision log
                workbook.Save(filePath);

                // -------------------------------------------------
                // 3. Load the workbook and inspect revision information
                // -------------------------------------------------
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                Workbook revWorkbook = new Workbook(filePath);
                bool changeVerified = false;

                // Get the collection of revision logs (may be null if no revisions)
                RevisionLogCollection revLogs = revWorkbook.Worksheets.RevisionLogs;

                if (revLogs != null)
                {
                    // Iterate through all revision logs
                    foreach (RevisionLog log in revLogs)
                    {
                        foreach (Revision rev in log.Revisions)
                        {
                            // Look for cell changes
                            if (rev is RevisionCellChange cellChange && cellChange.CellName == "A1")
                            {
                                // Retrieve old and new formula texts
                                string oldFormula = cellChange.OldFormula;
                                string newFormula = cellChange.NewFormula;

                                // Output the formulas for visibility
                                Console.WriteLine($"Cell: {cellChange.CellName}");
                                Console.WriteLine($"Old Formula: {oldFormula}");
                                Console.WriteLine($"New Formula: {newFormula}");

                                // Verify that the intended modification occurred
                                if (!string.IsNullOrEmpty(oldFormula) &&
                                    !string.IsNullOrEmpty(newFormula) &&
                                    oldFormula != newFormula &&
                                    newFormula == "=AVERAGE(B1:B3)")
                                {
                                    changeVerified = true;
                                    Console.WriteLine("Formula change verified successfully.");
                                }
                                else
                                {
                                    Console.WriteLine("Formula change verification failed.");
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No revision logs were found in the workbook.");
                }

                // Final result
                if (!changeVerified)
                {
                    Console.WriteLine("No matching formula change was detected in the revision logs.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
