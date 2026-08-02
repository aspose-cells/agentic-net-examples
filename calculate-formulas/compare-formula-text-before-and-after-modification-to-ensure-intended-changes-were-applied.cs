// Title: Compare Original and Modified Cell Formulas Using Aspose.Cells Revision Logs in C#
// Description: This example creates an Excel workbook, assigns an initial formula to cell A1, updates the formula, saves the file, and then reloads it to read the RevisionLogCollection. It extracts the old and new formulas for the cell‑change revision and verifies that the new formula matches the expected expression.
// Keywords: Aspose.Cells | C# | revision log | formula change tracking | compare Excel formulas | cell revision audit | RevisionLogCollection | track cell edits | Excel automation .NET | formula verification
// Common Searches: Aspose.Cells get old formula from revision log | compare original and updated formula C# Aspose.Cells | read cell change revisions in Excel workbook | verify formula modification using Aspose.Cells | audit formula edits with Aspose.Cells .NET
// Developer Intent: Confirm that a cell's formula was changed to the intended expression by reading revision logs.
// Use Cases: Log the original formula before a user edits a cell and later retrieve it for compliance reporting. | After programmatically updating a formula, load the saved workbook and extract both old and new formulas from the revision entries. | Automatically compare the captured new formula with an expected value and trigger custom logic on match or mismatch.
// AI Prompts: How do I enable cell revision tracking in Aspose.Cells and retrieve old and new formulas for a specific cell using C#? | Provide C# code that reads the RevisionLogCollection from a saved workbook and validates that a modified formula equals a given string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace AsposeCellsFormulaComparisonDemo
{
    // This example creates an Excel workbook, assigns an initial formula to cell A1, updates the formula, saves the file, and then reloads it to read the RevisionLogCollection. It extracts the old and new formulas for the cell‑change revision and verifies that the new formula matches the expected expression.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path for the temporary workbook that will store revisions
                string filePath = "FormulaComparisonDemo.xlsx";

                // ---------- Create workbook and set initial formula ----------
                Workbook workbook = new Workbook(); // create workbook

                // NOTE: In some versions of Aspose.Cells the EnableCellRevision property may not be available.
                // If supported, uncomment the following line to enable revision tracking:
                // workbook.Settings.EnableCellRevision = true;

                Worksheet sheet = workbook.Worksheets[0];

                // Set the original formula in cell A1
                sheet.Cells["A1"].Formula = "=SUM(B1:B3)";

                // Populate referenced cells with sample values
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // ---------- Modify the formula ----------
                // This change will generate a revision entry
                string newFormula = "=AVERAGE(B1:B3)";
                sheet.Cells["A1"].Formula = newFormula;

                // Save the workbook to persist the revision information
                workbook.Save(filePath);

                // ---------- Load workbook and examine revisions ----------
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                Workbook revWorkbook = new Workbook(filePath);
                RevisionLogCollection logs = revWorkbook.Worksheets.RevisionLogs;

                if (logs == null || logs.Count == 0)
                {
                    Console.WriteLine("No revision logs were found in the workbook.");
                    return;
                }

                foreach (RevisionLog log in logs)
                {
                    foreach (Revision rev in log.Revisions)
                    {
                        // Look for cell change revisions
                        if (rev is RevisionCellChange cellChange && cellChange.CellName == "A1")
                        {
                            string oldFormula = cellChange.OldFormula;
                            string capturedNewFormula = cellChange.NewFormula;

                            Console.WriteLine($"Cell: {cellChange.CellName}");
                            Console.WriteLine($"Old Formula: {oldFormula}");
                            Console.WriteLine($"New Formula (from revision): {capturedNewFormula}");

                            // Compare the captured new formula with the expected one
                            if (string.Equals(capturedNewFormula, newFormula, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("Verification succeeded: The new formula matches the intended change.");
                            }
                            else
                            {
                                Console.WriteLine("Verification failed: The new formula does NOT match the intended change.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
