// Title: Track and Compare Excel Formula Changes with Aspose.Cells Revision Logs in C#
// Description: This example creates a workbook, writes a SUM formula to cell A1, updates it to AVERAGE, saves the file, reloads it, and reads the revision logs. It locates the RevisionCellChange for A1, extracts the OldFormula and NewFormula values, and confirms that the change matches the expected expressions.
// Keywords: Aspose.Cells | C# | .NET | revision log | formula change | RevisionCellChange | OldFormula | NewFormula | audit Excel formulas | track cell edits
// Common Searches: Aspose.Cells read old and new formulas from revision log | How to detect formula changes in an Excel workbook using C# | Audit Excel cell formula modifications with Aspose.Cells | Retrieve RevisionCellChange details in .NET
// Developer Intent: Read a workbook’s revision history, locate a specific cell’s change entry, and compare its previous and current formulas.
// Use Cases: Identify unauthorized formula edits by scanning revision logs for critical cells. | Generate compliance reports that list every formula modification in a spreadsheet. | Validate automated formula updates before publishing the workbook.
// AI Prompts: Write C# code that opens an existing Excel file and prints the OldFormula and NewFormula for a given cell using Aspose.Cells revision logs. | Create a method that accepts a cell address, expected old formula, and expected new formula, then returns true if a matching RevisionCellChange exists. | Show how to iterate over all RevisionCellChange objects in a workbook and build a dictionary mapping cell addresses to their formula change history.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace AsposeCellsExamples
{
    // This example creates a workbook, writes a SUM formula to cell A1, updates it to AVERAGE, saves the file, reloads it, and reads the revision logs. It locates the RevisionCellChange for A1, extracts the OldFormula and NewFormula values, and confirms that the change matches the expected expressions.
    public class FormulaComparisonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set initial formula in cell A1 and populate referenced cells
                worksheet.Cells["A1"].Formula = "=SUM(B1:B3)";
                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["B2"].PutValue(20);
                worksheet.Cells["B3"].PutValue(30);

                // Change the formula in A1 to generate a revision entry
                worksheet.Cells["A1"].Formula = "=AVERAGE(B1:B3)";

                // Save the workbook to persist revision information
                string filePath = "FormulaRevisionDemo.xlsx";
                workbook.Save(filePath);

                // Ensure the file exists before attempting to reopen
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Reopen the workbook to read revision logs
                Workbook revisionWorkbook = new Workbook(filePath);

                // Iterate through revision logs to find the cell change for A1
                foreach (RevisionLog log in revisionWorkbook.Worksheets.RevisionLogs)
                {
                    foreach (Revision revision in log.Revisions)
                    {
                        if (revision is RevisionCellChange cellChange && cellChange.CellName == "A1")
                        {
                            // Retrieve old and new formulas from the revision entry
                            string oldFormula = cellChange.OldFormula;
                            string newFormula = cellChange.NewFormula;

                            Console.WriteLine($"Cell {cellChange.CellName} formula changed.");
                            Console.WriteLine($"Old Formula: {oldFormula}");
                            Console.WriteLine($"New Formula: {newFormula}");

                            // Verify the formulas match expected values
                            if (oldFormula == "=SUM(B1:B3)" && newFormula == "=AVERAGE(B1:B3)")
                            {
                                Console.WriteLine("Formula change verified successfully.");
                            }
                            else
                            {
                                Console.WriteLine("Formula change does not match expected values.");
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

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            FormulaComparisonDemo.Run();
        }
    }
}
