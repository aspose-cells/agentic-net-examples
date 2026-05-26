using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsVolatileFormulaScanner
{
    class Program
    {
        static void Main()
        {
            // ---------- Create / Load ----------
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // List of known volatile Excel functions (case‑insensitive)
            var volatileFunctions = new List<string>
            {
                "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET",
                "INDIRECT", "INFO", "CELL", "AREAS", "ERROR.TYPE",
                "GETPIVOTDATA", "NOW", "TODAY", "RAND", "RANDBETWEEN"
            };

            // ---------- Scan for volatile formulas ----------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data or formulas
                foreach (Cell cell in cells)
                {
                    // Skip cells without a formula
                    if (string.IsNullOrEmpty(cell.Formula))
                        continue;

                    string formula = cell.Formula;

                    // Check if the formula contains any volatile function name
                    foreach (string vf in volatileFunctions)
                    {
                        // Look for the function name followed by '(' to avoid false positives
                        if (formula.IndexOf(vf + "(", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Log the location and the formula
                            Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}");
                            // Once a volatile function is found, no need to check the rest for this cell
                            break;
                        }
                    }
                }
            }

            // ---------- Save ----------
            // Save the workbook (optional – unchanged in this example)
            workbook.Save("output.xlsx");
        }
    }
}