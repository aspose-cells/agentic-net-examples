using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaAnalysis
{
    class Program
    {
        // Thresholds to consider a worksheet as "very large"
        const int LargeRowThreshold = 100_000;
        const int LargeColumnThreshold = 1_000;

        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Identify worksheets that exceed the size thresholds
            var largeSheets = new HashSet<string>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow/Column give the last row/column that contains data
                int maxRow = sheet.Cells.MaxDataRow + 1;      // zero‑based index -> count
                int maxCol = sheet.Cells.MaxDataColumn + 1;

                if (maxRow > LargeRowThreshold || maxCol > LargeColumnThreshold)
                {
                    largeSheets.Add(sheet.Name);
                }
            }

            // Prepare a list to hold detected problematic formulas
            var problematicFormulas = new List<string>();

            // Scan all cells with formulas and see if they reference a large sheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Enumerate only cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;

                    // Simple detection: does the formula contain a sheet name that is large?
                    foreach (string largeSheetName in largeSheets)
                    {
                        // Look for patterns like 'Sheet1'! or Sheet1!
                        if (formula.IndexOf($"{largeSheetName}!", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            formula.IndexOf($"'{largeSheetName}'!", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            problematicFormulas.Add(
                                $"Worksheet '{sheet.Name}'!{cell.Name}: {formula}");
                            break;
                        }
                    }
                }
            }

            // Output analysis results
            Console.WriteLine("=== Large Worksheet Detection ===");
            if (largeSheets.Count == 0)
            {
                Console.WriteLine("No worksheets exceed the defined size thresholds.");
            }
            else
            {
                Console.WriteLine("Worksheets considered large:");
                foreach (string name in largeSheets)
                {
                    Console.WriteLine($"- {name}");
                }
            }

            Console.WriteLine("\n=== Formulas Referencing Large Worksheets ===");
            if (problematicFormulas.Count == 0)
            {
                Console.WriteLine("No formulas reference large worksheets.");
            }
            else
            {
                foreach (string info in problematicFormulas)
                {
                    Console.WriteLine(info);
                }
            }

            // Suggest optimization strategies
            Console.WriteLine("\n=== Suggested Optimization Strategies ===");
            Console.WriteLine("1. Use named ranges or structured table references instead of full sheet ranges.");
            Console.WriteLine("2. Limit the calculation range with CalculationOptions (e.g., set CalcStackSize).");
            Console.WriteLine("3. Enable AccessCacheOptions for read‑only large data sets:");
            Console.WriteLine("   workbook.StartAccessCache(AccessCacheOptions.All);");
            Console.WriteLine("   // perform read‑only operations");
            Console.WriteLine("   workbook.CloseAccessCache(AccessCacheOptions.All);");
            Console.WriteLine("4. If using dynamic array formulas, call RefreshDynamicArrayFormulas(false) to avoid full recalc.");
            Console.WriteLine("5. Consider converting volatile functions (e.g., INDIRECT, OFFSET) to static references.");
            Console.WriteLine("6. Split very large worksheets into multiple smaller ones if possible.");

            // Example: apply AccessCacheOptions to the workbook (read‑only scenario)
            workbook.StartAccessCache(AccessCacheOptions.All);
            // ... perform any read‑only analysis here ...
            workbook.CloseAccessCache(AccessCacheOptions.All);

            // Save the workbook (optional – here we just save a copy)
            workbook.Save("AnalyzedWorkbook.xlsx");
        }
    }
}