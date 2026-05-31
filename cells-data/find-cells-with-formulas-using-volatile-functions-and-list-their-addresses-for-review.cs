using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace VolatileFunctionFinder
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain to ensure formulas are evaluated correctly
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // List of known volatile functions (case‑insensitive search)
            string[] volatileFunctions = new[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET(", "INDIRECT(", "INFO(", "CELL(", "NOW", "TODAY"
            };

            // Collect addresses of cells that contain volatile functions
            List<string> volatileCellAddresses = new List<string>();

            // Iterate through all worksheets and their used cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Use the maximum used row/column to limit the scan
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            // Simple case‑insensitive check for any volatile function name
                            foreach (string vf in volatileFunctions)
                            {
                                if (formula.IndexOf(vf, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    volatileCellAddresses.Add($"{sheet.Name}!{cell.Name}");
                                    break; // No need to check other functions for this cell
                                }
                            }
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells containing volatile functions:");
            foreach (string address in volatileCellAddresses)
            {
                Console.WriteLine(address);
            }

            // Save the workbook (optional – here we just save a copy)
            workbook.Save("output.xlsx");
        }
    }
}