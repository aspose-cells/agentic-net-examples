// Title: Unmerge a merged range while preserving original formulas – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to store each cell's formula, merge a 3×3 block, unmerge it, and reapply the saved formulas before calculating and saving the workbook using Aspose.Cells.
// Keywords: Aspose.Cells unmerge preserve formulas | C# merge then unmerge cells | restore formulas after unmerge .NET | Aspose.Cells range handling | keep formulas when splitting merged cells
// Common Searches: Aspose.Cells unmerge range without losing formulas | how to keep cell formulas after unmerging in .NET | restore original formulas after merging cells Aspose | C# unmerge merged cells and retain formulas
// Developer Intent: The developer needs to split a merged block and ensure every resulting cell keeps the formula it had before the merge.
// Use Cases: Design a report with merged headers for layout, then programmatically split them while preserving calculation logic. | Process templates where merged cells are used for styling but later need individual formulas for data updates. | Automate workbook cleanup that separates merged areas without breaking dependent formulas.
// AI Prompts: Generate C# code with Aspose.Cells that merges a range, caches each cell's formula, unmerges the range, and restores the formulas before calling CalculateFormula. | Show how to capture formulas in a dictionary, unmerge a block, and reassign the formulas to the cells in Aspose.Cells for .NET. | Provide a step‑by‑step example of preserving formulas when unmerging a merged range, including workbook saving.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to store each cell's formula, merge a 3×3 block, unmerge it, and reapply the saved formulas before calculating and saving the workbook using Aspose.Cells.
    public class UnmergePreserveFormulas
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the range to work with (A1:C3)
                int firstRow = 0;      // A1 row index
                int firstColumn = 0;   // A1 column index
                int totalRows = 3;     // 3 rows
                int totalColumns = 3;  // 3 columns

                // Store original formulas of each cell in the range
                Dictionary<string, string> originalFormulas = new Dictionary<string, string>();
                for (int r = firstRow; r < firstRow + totalRows; r++)
                {
                    for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                    {
                        Cell cell = cells[r, c];
                        // Example formula: multiply row number by column number
                        cell.Formula = "=ROW()*COLUMN()";
                        // Save the formula text
                        originalFormulas[cell.Name] = cell.Formula;
                    }
                }

                // Merge the defined range into a single cell
                cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

                // Unmerge the range
                AsposeRange mergedRange = worksheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
                mergedRange.UnMerge();

                // Restore each cell's original formula after unmerging
                foreach (var kvp in originalFormulas)
                {
                    Cell cell = cells[kvp.Key];
                    cell.Formula = kvp.Value;
                }

                // Calculate formulas to populate values
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save("UnmergedPreserveFormulas.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnmergePreserveFormulas.Run();
        }
    }
}
