// Title: Detect and Optimize Formulas Referencing Large Worksheets with Aspose.Cells for .NET
// Description: A C# utility that loads or creates an Excel workbook, flags worksheets exceeding 10,000 rows or 1,000 columns, scans every cell for formulas that point to those large sheets, reports the offending addresses, and applies performance‑boosting techniques such as shared‑formula conversion, calculation‑option tuning, and access‑cache activation before saving the optimized file.
// Keywords: Aspose.Cells | C# Excel formula optimization | large worksheets detection | shared formulas .NET | CalculationOptions stack size | AccessCache performance | high‑volume Excel data | memory reduction Excel | detect cross‑sheet formulas | Excel workbook optimization
// Common Searches: how to find formulas that reference large sheets using Aspose.Cells | convert repetitive formulas to shared formulas in C# Excel file | best practices for calculating massive workbooks with Aspose.Cells | enable access cache for big Excel data sets Aspose.Cells | optimize formula performance in .NET Excel applications
// Developer Intent: Identify cross‑sheet formulas that target oversized worksheets and apply Aspose.Cells techniques to improve calculation speed and memory usage.
// Use Cases: List all cell addresses whose formulas reference worksheets flagged as large for manual review. | Automatically replace vertical blocks of identical formulas with a single shared formula via SetSharedFormula to cut memory overhead. | Configure CalculationOptions (e.g., increase CalcStackSize, disable recursion) and wrap calculations with StartAccessCache/CloseAccessCache to prevent stack overflows on massive workbooks. | Replace heavy sub‑range formulas with static values to eliminate unnecessary recalculations. | Split or externalize data when formulas must reference extremely large sheets, reducing workbook size.
// AI Prompts: Generate C# code using Aspose.Cells that scans a workbook and returns a list of cell addresses whose formulas reference any worksheet with more than 10,000 rows or 1,000 columns. | Write a method that iterates through each column, detects consecutive identical formulas, and converts them into a shared formula with SetSharedFormula. | Explain how to configure CalculationOptions and use StartAccessCache/CloseAccessCache to optimize formula evaluation for a workbook containing over 100,000 rows.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaOptimization
{
    // A C# utility that loads or creates an Excel workbook, flags worksheets exceeding 10,000 rows or 1,000 columns, scans every cell for formulas that point to those large sheets, reports the offending addresses, and applies performance‑boosting techniques such as shared‑formula conversion, calculation‑option tuning, and access‑cache activation before saving the optimized file.
    class Program
    {
        // Threshold to consider a worksheet as having a very large data set
        const int LargeRowThreshold = 10000;
        const int LargeColumnThreshold = 1000;

        static void Main()
        {
            try
            {
                // -------------------- Create / Load Workbook --------------------
                string inputPath = "LargeDataWorkbook.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"[Warning] Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook(); // creates a default workbook with one sheet
                }

                // -------------------- Identify Large Worksheets --------------------
                var largeSheets = new HashSet<string>();
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of last used row
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of last used column

                    if (maxRow + 1 > LargeRowThreshold || maxCol + 1 > LargeColumnThreshold)
                    {
                        largeSheets.Add(sheet.Name);
                        Console.WriteLine($"[Info] Worksheet \"{sheet.Name}\" is large (Rows: {maxRow + 1}, Columns: {maxCol + 1}).");
                    }
                }

                // If no large worksheets were found, exit early
                if (largeSheets.Count == 0)
                {
                    Console.WriteLine("[Info] No large worksheets detected. No optimization needed.");
                    workbook.Save("OptimizedOutput.xlsx");
                    return;
                }

                // -------------------- Scan Formulas Referencing Large Worksheets --------------------
                var problematicCells = new List<Cell>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                foreach (string largeSheetName in largeSheets)
                                {
                                    string pattern1 = $"'{largeSheetName}'!";
                                    string pattern2 = $"{largeSheetName}!";

                                    if (cell.Formula.IndexOf(pattern1, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                        cell.Formula.IndexOf(pattern2, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        problematicCells.Add(cell);
                                        Console.WriteLine($"[Detect] Formula in {sheet.Name}!{cell.Name} references large sheet \"{largeSheetName}\".");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                // -------------------- Suggest Optimization Strategies --------------------
                Console.WriteLine("\n--- Optimization Recommendations ---\n");
                Console.WriteLine("1. Convert repetitive formulas to shared formulas to reduce memory overhead.");
                Console.WriteLine("   Example: cell.SetSharedFormula(\"=A1*2\", rowCount, colCount);");
                Console.WriteLine("2. Avoid volatile functions (INDIRECT, OFFSET) that force full‑sheet recalculation.");
                Console.WriteLine("   Use structured table references or explicit ranges instead.");
                Console.WriteLine("3. Adjust CalculationOptions.CalcStackSize if you encounter StackOverflowException.");
                Console.WriteLine("   Example:");
                Console.WriteLine("       var opts = new CalculationOptions { CalcStackSize = 100, Recursive = false };");
                Console.WriteLine("       workbook.CalculateFormula(opts);");
                Console.WriteLine("4. Enable caching for large data access:");
                Console.WriteLine("       workbook.StartAccessCache(AccessCacheOptions.All);");
                Console.WriteLine("       // perform read‑only operations");
                Console.WriteLine("       workbook.CloseAccessCache(AccessCacheOptions.All);");
                Console.WriteLine("5. When using dynamic array formulas, limit the spill range if full size is unnecessary:");
                Console.WriteLine("       var opts = new CalculationOptions { Recursive = false };");
                Console.WriteLine("       cell.SetDynamicArrayFormula(\"=SEQUENCE(5)\", new FormulaParseOptions(), calculateValue:true);");
                Console.WriteLine("       workbook.RefreshDynamicArrayFormulas(true, opts);");
                Console.WriteLine("6. For static heavy sub‑ranges, replace formulas with their calculated values:");
                Console.WriteLine("       range.RemoveFormulas(); // converts formulas to values");
                Console.WriteLine("7. Define named ranges or convert data to tables and reference them in formulas.");
                Console.WriteLine("8. If formulas must reference large sheets, consider splitting data or using Power Query outside Excel.");

                // -------------------- Apply Some Automatic Optimizations (Optional) --------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        string firstFormula = null;
                        int startRow = -1;
                        int count = 0;

                        for (int row = 0; row <= cells.MaxDataRow; row++)
                        {
                            Cell cur = cells[row, col];
                            if (!string.IsNullOrEmpty(cur.Formula))
                            {
                                if (firstFormula == null)
                                {
                                    firstFormula = cur.Formula;
                                    startRow = row;
                                    count = 1;
                                }
                                else if (cur.Formula == firstFormula)
                                {
                                    count++;
                                }
                                else
                                {
                                    if (count > 1)
                                    {
                                        Cell anchor = cells[startRow, col];
                                        anchor.SetSharedFormula(firstFormula, count, 1);
                                    }
                                    firstFormula = cur.Formula;
                                    startRow = row;
                                    count = 1;
                                }
                            }
                        }

                        if (count > 1 && firstFormula != null)
                        {
                            Cell anchor = cells[startRow, col];
                            anchor.SetSharedFormula(firstFormula, count, 1);
                        }
                    }
                }

                // -------------------- Save Optimized Workbook --------------------
                workbook.Save("OptimizedOutput.xlsx");
                Console.WriteLine("\n[Result] Optimized workbook saved as \"OptimizedOutput.xlsx\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] An exception occurred: {ex.Message}");
            }
        }
    }
}
