// Title: Detect and list INDIRECT formulas in an Excel workbook with Aspose.Cells for .NET – performance impact guide
// Description: C# sample that loads a workbook, iterates every worksheet and used cell, identifies formulas containing the INDIRECT function (case‑insensitive), outputs their addresses, and explains why INDIRECT’s volatility can slow calculation in large files.
// Keywords: Aspose.Cells INDIRECT detection | C# scan Excel formulas | volatile Excel functions .NET | performance impact INDIRECT | list formulas using INDIRECT | Excel workbook analysis Aspose | calculate speed Excel volatile functions
// Common Searches: how to find INDIRECT formulas with Aspose.Cells | list volatile formulas in an Excel file using C# | detect INDIRECT function in .NET workbook | Excel performance issues caused by INDIRECT | scan workbook for volatile functions Aspose
// Developer Intent: Locate every formula that uses the INDIRECT function in a workbook to assess its effect on calculation speed.
// Use Cases: Generate a report of all INDIRECT formulas for review and possible replacement with direct references. | Add the scan to a CI/CD pipeline to flag excessive volatile formulas before releasing a spreadsheet. | Combine INDIRECT usage data with workbook size metrics to prioritize optimization in large models.
// AI Prompts: Create a method that returns a dictionary where each worksheet name maps to a list of cell addresses containing INDIRECT formulas using Aspose.Cells. | Extend the example to count INDIRECT occurrences per worksheet and display a summary after scanning. | Rewrite the program so it writes the detected INDIRECT formulas to a CSV file instead of printing to the console.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsIndirectScanner
{
    // C# sample that loads a workbook, iterates every worksheet and used cell, identifies formulas containing the INDIRECT function (case‑insensitive), outputs their addresses, and explains why INDIRECT’s volatility can slow calculation in large files.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold formulas that use INDIRECT
            List<string> indirectFormulas = new List<string>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (cell.IsFormula)
                    {
                        // Normalize formula string for case‑insensitive search
                        string formula = cell.Formula?.Trim() ?? string.Empty;

                        // Identify usage of the INDIRECT function
                        if (formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Store the full address and formula for reporting
                            indirectFormulas.Add($"{sheet.Name}!{cell.Name}: {formula}");
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Formulas that use the INDIRECT function:");
            if (indirectFormulas.Count == 0)
            {
                Console.WriteLine("  None found.");
            }
            else
            {
                foreach (string entry in indirectFormulas)
                {
                    Console.WriteLine("  " + entry);
                }
            }

            // Note on performance impact
            Console.WriteLine();
            Console.WriteLine("Performance Note:");
            Console.WriteLine("  The INDIRECT function is volatile—it recalculates whenever any cell changes,");
            Console.WriteLine("  even if the referenced cells are not directly affected. Excessive use of");
            Console.WriteLine("  INDIRECT can significantly degrade calculation speed, especially in large");
            Console.WriteLine("  workbooks or when calculation chains are enabled.");

            // Optionally, save the workbook after analysis (no changes made here)
            // string outputPath = "output.xlsx";
            // workbook.Save(outputPath);
        }
    }
}
