// Title: C# – Generate a Volatile Function Report After Workbook Recalculation with Aspose.Cells
// Description: Load or create an Excel workbook, enable the calculation chain, recalculate all formulas, scan every worksheet for volatile functions (NOW, TODAY, RAND, OFFSET, etc.), record the sheet name, cell address and formula, add a “VolatileReport” worksheet, optionally print the list to the console, and save the updated file.
// Keywords: Aspose.Cells volatile functions | C# detect volatile formulas | Excel volatile function audit | recalculate formulas Aspose.Cells | list volatile cells .NET | Excel performance tuning | calculation chain Aspose | scan worksheets for volatile | generate report worksheet | Aspose.Cells API example
// Common Searches: how to list cells with volatile functions using Aspose.Cells C# | Aspose.Cells create volatile formula report after recalculation | detect NOW() and RAND() formulas in Excel with Aspose.Cells | C# scan workbook for volatile Excel functions | Aspose.Cells generate audit sheet for volatile formulas
// Developer Intent: Identify every cell that uses a volatile Excel function after recalculating the workbook and produce a summary worksheet.
// Use Cases: Audit large workbooks to pinpoint volatile formulas that may degrade performance. | Provide developers with a quick console dump of volatile cells for debugging. | Create a reusable “VolatileReport” sheet for downstream analysis or documentation.
// AI Prompts: Write C# code with Aspose.Cells that recalculates a workbook, finds all volatile functions, and adds a summary worksheet. | Show a helper method that checks a formula string for any volatile function from a predefined list. | Explain how to enable the calculation chain, trigger full formula recalculation, and iterate through cells to detect volatile formulas using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace VolatileFunctionsReport
{
    // Load or create an Excel workbook, enable the calculation chain, recalculate all formulas, scan every worksheet for volatile functions (NOW, TODAY, RAND, OFFSET, etc.), record the sheet name, cell address and formula, add a “VolatileReport” worksheet, optionally print the list to the console, and save the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_volatile_report.xlsx";

                // -------------------------------------------------
                // 1. Load an existing workbook (or create a new one if missing)
                // -------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // -------------------------------------------------
                // 2. Ensure calculation chain is enabled and recalculate all formulas
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 3. Scan all worksheets and cells for volatile functions
                // -------------------------------------------------
                List<Cell> volatileCells = new List<Cell>();
                string[] volatileFunctions = new[]
                {
                    "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "AREAS()", "CHOOSE()", "HYPERLINK()"
                };

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.IsFormula && ContainsVolatile(cell.Formula, volatileFunctions))
                            {
                                volatileCells.Add(cell);
                            }
                        }
                    }
                }

                // -------------------------------------------------
                // 4. Create a report sheet and write the results
                // -------------------------------------------------
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "VolatileReport";
                Cells reportCells = reportSheet.Cells;

                // Header
                reportCells[0, 0].PutValue("Sheet");
                reportCells[0, 1].PutValue("Cell");
                reportCells[0, 2].PutValue("Formula");

                // Populate report
                for (int i = 0; i < volatileCells.Count; i++)
                {
                    Cell vCell = volatileCells[i];
                    int rowIdx = i + 1; // start after header

                    reportCells[rowIdx, 0].PutValue(vCell.Worksheet.Name);
                    reportCells[rowIdx, 1].PutValue(vCell.Name);
                    reportCells[rowIdx, 2].PutValue(vCell.Formula);
                }

                // -------------------------------------------------
                // 5. Output the list to console (optional)
                // -------------------------------------------------
                Console.WriteLine("Cells containing volatile functions:");
                foreach (Cell vCell in volatileCells)
                {
                    Console.WriteLine($"{vCell.Worksheet.Name}!{vCell.Name}  ->  {vCell.Formula}");
                }

                // -------------------------------------------------
                // 6. Save the workbook with the new report sheet
                // -------------------------------------------------
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper to detect volatile functions in a formula
        private static bool ContainsVolatile(string formula, string[] volatileFunctions)
        {
            if (string.IsNullOrEmpty(formula))
                return false;

            foreach (string func in volatileFunctions)
            {
                if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
    }
}
