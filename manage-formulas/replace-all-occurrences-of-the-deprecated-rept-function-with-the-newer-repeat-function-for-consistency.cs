// Title: Replace REPT with REPEAT in Excel formulas using Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet and cell, detects formulas that contain the deprecated REPT function (case‑insensitive), swaps it for the newer REPEAT function, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | Excel formula replace | REPT to REPEAT | bulk formula update | deprecated Excel function | programmatic Excel cleanup | formula search and replace | Excel automation
// Common Searches: replace REPT with REPEAT Aspose.Cells C# | bulk update Excel formulas .NET | iterate worksheets and modify formulas Aspose.Cells | programmatically change deprecated Excel functions | search and replace formulas in Excel workbook C#
// Developer Intent: Swap every REPT occurrence for REPEAT in all formula cells of a workbook.
// Use Cases: Modernize legacy spreadsheets that still use REPT. | Ensure consistency after upgrading to newer Excel versions. | Automate cleanup of deprecated functions across multiple workbooks.
// AI Prompts: Write C# code with Aspose.Cells that finds and replaces REPT with REPEAT in all formula cells, handling case‑insensitivity. | Provide a snippet that logs the address of each cell where REPT was changed to REPEAT. | Explain how to extend the logic to replace several deprecated functions in one pass.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans every worksheet and cell, detects formulas that contain the deprecated REPT function (case‑insensitive), swaps it for the newer REPEAT function, and saves the updated file.
    public class ReplaceReptWithRepeat
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula)
                        {
                            // Detect deprecated REPT function (case‑insensitive)
                            if (cell.Formula.IndexOf("REPT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Replace REPT with REPEAT, preserving case‑insensitivity
                                string updatedFormula = cell.Formula.Replace("REPT", "REPEAT", StringComparison.OrdinalIgnoreCase);
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
