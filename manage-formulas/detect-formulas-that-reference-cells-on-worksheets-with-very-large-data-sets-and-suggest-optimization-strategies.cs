// Title: Identify and Optimize Formulas in Massive Excel Worksheets with Aspose.Cells for .NET
// Description: C# sample that loads a workbook, flags sheets exceeding a configurable cell‑count (default 1 M), lists formula cells and provides performance tips such as limiting whole‑range references, removing volatile functions, applying shared formulas, enabling calculation cache, and using Excel tables.
// Keywords: Aspose.Cells | C# | formula optimization | large worksheet | whole column reference | volatile function detection | shared formulas | calculation cache | Excel tables | performance tuning
// Common Searches: How to detect formulas that reference whole columns in a big Excel file using Aspose.Cells C# | Best practices for removing volatile functions from worksheets with over 1 million cells | Enable calculation cache and set stack size for large workbooks in Aspose.Cells | Convert repetitive formulas to shared formulas with Aspose.Cells .NET | Optimize performance of massive Excel sheets using Aspose.Cells
// Developer Intent: Find formula cells in worksheets that exceed a size threshold and receive actionable suggestions to improve calculation speed and memory usage.
// Use Cases: Scan a workbook and list every formula cell in sheets larger than 1 M cells, flagging whole‑range references (e.g., A:A) for reduction. | Detect volatile functions such as NOW(), RAND(), or RANDBETWEEN in large sheets and recommend alternatives. | Suggest converting short, repetitive formulas to shared formulas via Cell.SetSharedFormula to lower memory consumption. | Provide workbook‑level recommendations like enabling calculation cache, adjusting the calculation stack, and converting data ranges to Excel tables.
// AI Prompts: Write C# code using Aspose.Cells that iterates through each worksheet, identifies formula cells in sheets with more than 1,000,000 cells, and logs suggestions for whole‑range references, volatile functions, and shared formulas. | Generate a snippet that enables the calculation cache and sets a custom calculation stack size before calling CalculateFormula on a large workbook with Aspose.Cells. | Create a function that automatically converts repetitive short formulas in a massive worksheet to shared formulas using Cell.SetSharedFormula.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaOptimization
{
    // C# sample that loads a workbook, flags sheets exceeding a configurable cell‑count (default 1 M), lists formula cells and provides performance tips such as limiting whole‑range references, removing volatile functions, applying shared formulas, enabling calculation cache, and using Excel tables.
    class Program
    {
        // Threshold to consider a worksheet as “very large”.
        // Adjust based on your environment (e.g., 1,000,000 cells).
        const long LargeSheetCellCountThreshold = 1_000_000;

        static void Main()
        {
            // Load an existing workbook (replace with your file path).
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range size.
                int maxRow = cells.MaxDataRow;   // zero‑based index of last row with data
                int maxCol = cells.MaxDataColumn; // zero‑based index of last column with data
                long totalCells = ((long)maxRow + 1) * ((long)maxCol + 1);

                // If the sheet is large, analyze its formulas.
                if (totalCells >= LargeSheetCellCountThreshold)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" is large ({totalCells:N0} cells).");

                    // Collect cells that contain formulas.
                    List<Cell> formulaCells = new List<Cell>();
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                formulaCells.Add(cell);
                            }
                        }
                    }

                    Console.WriteLine($"  Found {formulaCells.Count} formula cells.");

                    // Suggest optimization strategies for each formula.
                    foreach (Cell formulaCell in formulaCells)
                    {
                        // Example heuristics:
                        // 1. If the formula references a whole column/row range, suggest using structured tables or dynamic arrays.
                        // 2. If many similar formulas exist, suggest using shared formulas.
                        // 3. If the formula is volatile (e.g., NOW(), RAND()), suggest limiting its use.
                        // 4. If the formula result is used only for display, consider pre‑calculating values and storing them.

                        string formula = formulaCell.Formula;

                        // Simple check for whole‑column/row references (e.g., A:A or 1:1).
                        if (formula.Contains(":") && (formula.Contains("A:A") || formula.Contains("1:1")))
                        {
                            Console.WriteLine($"    Cell {formulaCell.Name} uses whole‑range reference. " +
                                              $"Consider limiting the range or using a Table.");
                        }

                        // Detect volatile functions.
                        if (formula.IndexOf("NOW()", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            formula.IndexOf("RAND()", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            formula.IndexOf("RANDBETWEEN", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"    Cell {formulaCell.Name} contains volatile function. " +
                                              $"Avoid if possible for large sheets.");
                        }

                        // Suggest shared formulas when many adjacent cells have similar formulas.
                        // Here we just demonstrate the idea; a real implementation would analyze patterns.
                        // For demonstration, we output a suggestion if the formula length is short.
                        if (formula.Length < 30)
                        {
                            Console.WriteLine($"    Cell {formulaCell.Name} may benefit from a shared formula. " +
                                              $"Use Cell.SetSharedFormula to reduce memory and calculation time.");
                        }
                    }

                    // General workbook‑level suggestions for large sheets.
                    Console.WriteLine("  General optimization suggestions for this worksheet:");
                    Console.WriteLine("    • Enable calculation cache: workbook.StartAccessCache(AccessCacheOptions.All);");
                    Console.WriteLine("    • Reduce recursive stack size if deep dependency chains exist:");
                    Console.WriteLine("        CalculationOptions opts = new CalculationOptions { CalcStackSize = 100 };");
                    Console.WriteLine("        workbook.CalculateFormula(opts);");
                    Console.WriteLine("    • If using dynamic array formulas, refresh them after data changes:");
                    Console.WriteLine("        workbook.RefreshDynamicArrayFormulas(true);");
                    Console.WriteLine("    • Convert repetitive formulas to shared formulas via Cell.SetSharedFormula.");
                    Console.WriteLine("    • Consider converting large data ranges to Excel Tables and use structured references.");
                }
            }

            // Save the workbook (optional – here we just save unchanged file).
            workbook.Save("output.xlsx");
        }
    }
}
