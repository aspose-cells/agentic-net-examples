using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsVolatileFunctionFinder
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure formulas are calculated (optional, but required for some APIs)
            workbook.CalculateFormula();

            // List to hold addresses of cells that use volatile functions
            List<string> volatileCells = new List<string>();

            // Define a set of known volatile function names (case‑insensitive)
            string[] volatileFunctions = new string[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only formula cells
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula?.ToUpperInvariant() ?? string.Empty;

                        // Check if the formula contains any volatile function
                        foreach (string volFunc in volatileFunctions)
                        {
                            // Simple containment check; more sophisticated parsing can be added if needed
                            if (formula.Contains(volFunc))
                            {
                                // Record the full address including sheet name
                                volatileCells.Add($"{sheet.Name}!{cell.Name}");
                                break; // No need to check other functions for this cell
                            }
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells containing volatile functions:");
            foreach (string address in volatileCells)
            {
                Console.WriteLine(address);
            }

            // Save the workbook (unchanged, but required by lifecycle rule)
            workbook.Save("output.xlsx");
        }
    }
}