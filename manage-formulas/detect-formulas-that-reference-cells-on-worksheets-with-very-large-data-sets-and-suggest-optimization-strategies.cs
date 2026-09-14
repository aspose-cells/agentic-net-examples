// Title: Identify formulas that reference very large worksheets and suggest performance optimizations with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that scans a workbook, marks worksheets with 100,000+ rows, and lists every formula that points to those sheets. | Extend the detection routine to output a set of concrete performance‑improvement recommendations for each formula that references a large worksheet. | Replace the simple string check with Aspose.Cells FormulaParser to reliably extract sheet names from complex formulas.
// Common Searches: aspnet find Excel formulas that reference worksheets with more than 100k rows using Aspose.Cells | how to improve performance of formulas that point to huge data tables in Aspose.Cells C# | detect and list formulas referencing massive sheets in a workbook with Aspose.Cells .NET | best practices for optimizing Excel calculations on large worksheets via Aspose.Cells
// Tags: detect formulas referencing massive worksheets Aspose.Cells | improve calculation speed for formulas on big sheets Aspose.Cells | named range usage to limit large sheet references Aspose.Cells | replace volatile functions with static references Aspose.Cells | manual calculation mode for faster workbook recalculation Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaOptimization
{
    // The example loads an Excel workbook, flags worksheets containing 100,000 or more rows, scans all formula cells, reports any formulas that reference those large sheets, and prints actionable optimization tips such as using named ranges, avoiding volatile functions, leveraging structured tables, creating summary helper sheets, and switching to manual calculation mode.
    class Program
    {
        // Threshold for considering a worksheet as having a very large data set.
        const int LargeRowThreshold = 100000;

        static void Main(string[] args)
        {
            // Load the workbook (replace with actual path or stream as needed).
            Workbook workbook = new Workbook("input.xlsx");

            // Identify worksheets that contain a large number of rows.
            var largeSheets = new HashSet<string>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the maximum used row index (0‑based). Add 1 to get count.
                int usedRows = sheet.Cells.MaxDataRow + 1;
                if (usedRows >= LargeRowThreshold)
                {
                    largeSheets.Add(sheet.Name);
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" has {usedRows} rows (>= {LargeRowThreshold}). Marked as large.");
                }
            }

            // Scan all formulas and detect references to large worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate only through cells that contain formulas.
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula; // e.g., =SUM(Sheet2!A1:A1000)
                    foreach (string largeSheetName in largeSheets)
                    {
                        // Simple detection: check if the formula text contains the sheet name followed by '!'
                        // This works for most cases; for more robust parsing, use FormulaParser if needed.
                        if (formula.IndexOf($"{largeSheetName}!", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"Formula in {sheet.Name}!{cell.Name} references large sheet \"{largeSheetName}\".");
                            SuggestOptimizations(sheet.Name, cell.Name, largeSheetName);
                            break; // Avoid duplicate messages for the same formula.
                        }
                    }
                }
            }

            // Optionally, save the workbook if any modifications were made.
            // workbook.Save("output.xlsx");
        }

        static void SuggestOptimizations(string formulaSheet, string formulaCell, string referencedLargeSheet)
        {
            Console.WriteLine($"--- Optimization suggestions for {formulaSheet}!{formulaCell} ---");
            Console.WriteLine("1. Use named ranges on the large sheet to limit the referenced area.");
            Console.WriteLine("2. Replace volatile functions (e.g., INDIRECT, OFFSET) with static references where possible.");
            Console.WriteLine("3. Consider using structured tables and referencing table columns instead of whole ranges.");
            Console.WriteLine("4. If only summary data is needed, create a helper sheet that aggregates the large data set and reference that sheet.");
            Console.WriteLine("5. Enable manual calculation mode and recalculate only when necessary to improve performance.");
            Console.WriteLine("------------------------------------------------------------");
        }
    }
}
