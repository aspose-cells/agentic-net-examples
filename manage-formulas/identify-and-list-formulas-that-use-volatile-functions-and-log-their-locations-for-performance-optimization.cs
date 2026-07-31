// Title: Aspose.Cells C# – Scan Workbook for Volatile Formulas and Generate a Report Sheet
// Description: Loads an Excel file with Aspose.Cells, walks every used cell in each worksheet, detects formulas that contain volatile functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL), records the sheet name, cell address and formula, writes the findings to a new worksheet called "VolatileFormulasReport" and saves the workbook. Ideal for performance tuning and audit of large workbooks.
// Keywords: Aspose.Cells volatile formula detection | C# scan Excel for volatile functions | list volatile Excel formulas Aspose | performance optimization Excel formulas | .NET identify volatile functions | generate volatile formula report | Aspose.Cells workbook audit | Excel volatile functions NOW OFFSET RAND
// Common Searches: How to find volatile Excel functions with Aspose.Cells C# | Aspose.Cells scan workbook for NOW or OFFSET formulas | Create a report of cells that use volatile functions in .NET | Identify performance‑impacting formulas using Aspose.Cells | C# code to list volatile formulas in an Excel file
// Developer Intent: Locate every cell that uses a volatile Excel function and log its sheet, address, and formula in a dedicated report worksheet for performance analysis.
// Use Cases: Audit large spreadsheets to pinpoint volatile formulas that slow recalculation. | Document volatile functions before migrating workbooks to a new platform. | Automate a quality‑gate that flags volatile formulas for manual refactoring. | Generate a compliance report showing where volatile functions are used.
// AI Prompts: Generate C# Aspose.Cells code that extracts volatile formulas and saves the result as CSV. | Show how to replace the simple string search with a regular‑expression matcher for exact volatile function names. | Explain how to exclude certain worksheets from the volatile‑function scan while still creating the report.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsVolatileFormulaScanner
{
    // Loads an Excel file with Aspose.Cells, walks every used cell in each worksheet, detects formulas that contain volatile functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL), records the sheet name, cell address and formula, writes the findings to a new worksheet called "VolatileFormulasReport" and saves the workbook. Ideal for performance tuning and audit of large workbooks.
    class Program
    {
        // List of known volatile Excel functions (case‑insensitive)
        static readonly string[] VolatileFunctions = new[]
        {
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT",
            "INFO", "CELL", "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", 
            "OFFSET()", "INDIRECT()", "INFO()", "CELL()"
        };

        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Prepare a list to hold information about volatile formulas
            var volatileCells = new List<(string SheetName, string CellName, string Formula)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            // Simple check: does the formula contain any volatile function name?
                            foreach (string vf in VolatileFunctions)
                            {
                                // Use InvariantCultureIgnoreCase for case‑insensitive search
                                if (formula.IndexOf(vf, StringComparison.InvariantCultureIgnoreCase) >= 0)
                                {
                                    volatileCells.Add((sheet.Name, cell.Name, formula));
                                    break; // No need to check other volatile functions for this cell
                                }
                            }
                        }
                    }
                }
            }

            // Output results to console
            Console.WriteLine("Volatile formulas found:");
            foreach (var entry in volatileCells)
            {
                Console.WriteLine($"Sheet: {entry.SheetName}, Cell: {entry.CellName}, Formula: {entry.Formula}");
            }

            // Optionally, write the findings to a new worksheet for documentation
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "VolatileFormulasReport";
            Cells reportCells = reportSheet.Cells;

            // Header row
            reportCells[0, 0].PutValue("Sheet");
            reportCells[0, 1].PutValue("Cell");
            reportCells[0, 2].PutValue("Formula");

            // Populate rows
            for (int i = 0; i < volatileCells.Count; i++)
            {
                var v = volatileCells[i];
                reportCells[i + 1, 0].PutValue(v.SheetName);
                reportCells[i + 1, 1].PutValue(v.CellName);
                reportCells[i + 1, 2].PutValue(v.Formula);
            }

            // Save the workbook with the report (replace with your desired output path)
            workbook.Save("output_with_volatile_report.xlsx");
        }
    }
}
