// Title: Validate visual consistency after removing unused styles with AspNet Aspose.Cells (C#)
// Description: C# sample that loads an original workbook, clones it, records the style pool size, calls RemoveUnusedStyles on the clone, then iterates every worksheet, row and column to compare each cell's display style using Style.Equals. It reports mismatches and confirms that formatting remains identical, finally saving the cleaned file for manual review.
// Keywords: Aspose.Cells RemoveUnusedStyles | C# compare workbook styles | Excel style pool size | GetDisplayStyle | Style.Equals | visual consistency check | Excel file size optimization | cell formatting verification
// Common Searches: how to ensure removing unused styles does not change Excel appearance Aspose.Cells | C# compare two Excel workbooks cell style equality | verify visual consistency after RemoveUnusedStyles | Aspose.Cells count of styles before after cleanup | detect formatting differences in cloned workbook
// Developer Intent: Confirm that calling RemoveUnusedStyles on a workbook does not alter any cell's visual formatting.
// Use Cases: Automated regression test to catch unintended style changes after workbook optimization. | Generate a size‑reduced Excel file for distribution while guaranteeing identical look and feel. | Audit large workbooks to ensure style cleanup does not affect end‑user presentation.
// AI Prompts: Write C# code that loads an Excel file, creates a copy, removes unused styles with Aspose.Cells, and lists cells where the display style differs. | Explain the effect of RemoveUnusedStyles on the style pool and how to use GetDisplayStyle and Style.Equals for visual verification. | Provide a step‑by‑step tutorial for comparing cell formatting between two workbooks after style cleanup using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsStyleComparison
{
    // C# sample that loads an original workbook, clones it, records the style pool size, calls RemoveUnusedStyles on the clone, then iterates every worksheet, row and column to compare each cell's display style using Style.Equals. It reports mismatches and confirms that formatting remains identical, finally saving the cleaned file for manual review.
    class Program
    {
        static void Main()
        {
            // Path to the original workbook (with all styles)
            string originalPath = "original.xlsx";

            // Load the original workbook
            Workbook originalWb = new Workbook(originalPath);

            // Create a copy of the original workbook to work on removing unused styles
            Workbook cleanedWb = new Workbook();
            cleanedWb.Copy(originalWb); // use the provided Copy method

            // Record style count before removal
            int styleCountBefore = originalWb.CountOfStylesInPool;
            Console.WriteLine($"Style count before removal: {styleCountBefore}");

            // Remove unused styles from the copy
            cleanedWb.RemoveUnusedStyles();

            // Record style count after removal
            int styleCountAfter = cleanedWb.CountOfStylesInPool;
            Console.WriteLine($"Style count after removal: {styleCountAfter}");

            // Compare visual consistency cell by cell
            bool allMatch = true;
            int sheetCount = originalWb.Worksheets.Count;

            for (int s = 0; s < sheetCount; s++)
            {
                Worksheet sheetOriginal = originalWb.Worksheets[s];
                Worksheet sheetCleaned = cleanedWb.Worksheets[s];

                // Determine the used range (max row/column) for iteration
                int maxRow = Math.Max(sheetOriginal.Cells.MaxDataRow, sheetCleaned.Cells.MaxDataRow);
                int maxCol = Math.Max(sheetOriginal.Cells.MaxDataColumn, sheetCleaned.Cells.MaxDataColumn);

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get display styles for the current cell in both workbooks
                        Style styleOriginal = sheetOriginal.Cells[row, col].GetDisplayStyle();
                        Style styleCleaned = sheetCleaned.Cells[row, col].GetDisplayStyle();

                        // Use Style.Equals to compare the two styles
                        if (!styleOriginal.Equals(styleCleaned))
                        {
                            allMatch = false;
                            Console.WriteLine($"Mismatch found at Sheet[{s}] Cell[{row},{col}]");
                        }
                    }
                }
            }

            if (allMatch)
                Console.WriteLine("Visual consistency verified: all cell styles match after removing unused styles.");
            else
                Console.WriteLine("Visual inconsistency detected: some cell styles differ after removing unused styles.");

            // Optionally, save the cleaned workbook for manual inspection
            cleanedWb.Save("cleaned_without_unused_styles.xlsx");
        }
    }
}
