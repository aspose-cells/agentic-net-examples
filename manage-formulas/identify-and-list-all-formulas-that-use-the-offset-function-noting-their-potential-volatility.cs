// Title: Find and list OFFSET formulas in an Excel workbook with Aspose.Cells for .NET
// Description: C# sample that opens a workbook, scans all worksheets and used cells, detects formulas containing the OFFSET function (case‑insensitive), records the sheet name, cell address and formula text, prints the findings, and saves an unchanged copy. Highlights the volatility of OFFSET for performance analysis.
// Keywords: Aspose.Cells | C# | .NET | Excel OFFSET function | volatile formula detection | list Excel formulas | programmatic workbook analysis | find OFFSET in .xlsx | formula audit | Excel performance tuning
// Common Searches: Aspose.Cells list OFFSET formulas | C# detect volatile Excel formulas | enumerate OFFSET functions in .xlsx using .NET | how to find cells with OFFSET in Excel programmatically | Excel OFFSET volatility detection code
// Developer Intent: Locate every cell that uses the OFFSET function to evaluate its effect on calculation speed and stability.
// Use Cases: Create a performance‑focused report of all OFFSET formulas in large workbooks. | Audit workbook for volatile functions before migration or optimization. | Generate documentation of OFFSET usage for refactoring to non‑volatile alternatives. | Provide a checklist for compliance teams to verify formula stability.
// AI Prompts: Write C# code with Aspose.Cells that extracts all formulas containing OFFSET and outputs sheet, address, and formula. | Explain how to assess the volatility of Excel functions like OFFSET using Aspose.Cells and suggest mitigation techniques. | Provide a unit test for ListOffsetFormulas that confirms correct identification of OFFSET formulas. | Suggest ways to replace OFFSET with INDEX or other non‑volatile functions in a workbook processed by Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that opens a workbook, scans all worksheets and used cells, detects formulas containing the OFFSET function (case‑insensitive), records the sheet name, cell address and formula text, prints the findings, and saves an unchanged copy. Highlights the volatility of OFFSET for performance analysis.
    public class ListOffsetFormulas
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // List to hold cells with OFFSET formulas
                List<string> offsetFormulas = new List<string>();

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
                            // Determine if the formula uses the OFFSET function (case‑insensitive)
                            if (cell.Formula.IndexOf("OFFSET", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the cell address and its formula
                                offsetFormulas.Add($"{sheet.Name}!{cell.Name}: {cell.Formula}");
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Formulas that use the OFFSET function (volatile):");
                foreach (string entry in offsetFormulas)
                {
                    Console.WriteLine(entry);
                }

                // Save a copy of the workbook (no changes made here)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListOffsetFormulas.Run();
        }
    }
}
