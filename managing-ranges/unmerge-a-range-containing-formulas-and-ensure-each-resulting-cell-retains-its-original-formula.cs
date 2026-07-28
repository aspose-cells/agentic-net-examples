// Title: Aspose.Cells for .NET – Unmerge a merged range while preserving each cell’s original formula (C#)
// Description: Shows how to capture the formula of every cell in a merged block, unmerge the range with Range.UnMerge, and restore the original formulas to the individual cells before saving the workbook. The sample creates a 3×3 range (B2:D4), assigns a formula, merges, stores formulas in a 2‑D array, unmerges, and reapplies the formulas.
// Keywords: Aspose.Cells unmerge range C# | preserve formulas after unmerge | Range.UnMerge Aspose.Cells | restore cell formulas .NET | merged cells to individual cells Aspose | C# spreadsheet formula preservation | Aspose.Cells example unmerge preserve formulas
// Common Searches: how to unmerge merged cells without losing formulas Aspose.Cells | preserve individual formulas when unmerging a range C# | Aspose.Cells unmerge range keep formulas | save formulas before unmerge Aspose.Cells .NET | restore formulas after unmerge spreadsheet
// Developer Intent: Unmerge a previously merged range and ensure each resulting cell retains the formula it had before the merge.
// Use Cases: Create a report layout with merged cells, then split them for further calculations while keeping all formulas intact. | Clean up imported workbooks that contain merged cells with formulas before exporting to another system. | Automate worksheet preparation where merged regions serve as temporary placeholders and must be unmerged without losing formula data.
// AI Prompts: Generate C# code using Aspose.Cells that unmerges a merged range and automatically restores each cell’s original formula without external storage. | Explain step‑by‑step how to capture formulas from a merged block, unmerge it with Range.UnMerge, and reapply the formulas to the individual cells in Aspose.Cells for .NET. | Provide an alternative method (e.g., Clone, Copy, or Style) to preserve formulas when unmerging a range, including sample C# code.

using System;
using Aspose.Cells;

// Shows how to capture the formula of every cell in a merged block, unmerge the range with Range.UnMerge, and restore the original formulas to the individual cells before saving the workbook. The sample creates a 3×3 range (B2:D4), assigns a formula, merges, stores formulas in a 2‑D array, unmerges, and reapplies the formulas.
class UnmergePreserveFormulas
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the range to merge (B2:D4) – zero‑based indices
            int firstRow = 1;   // B2 row index
            int firstCol = 1;   // B2 column index
            int totalRows = 3;  // B2 to D4 spans 3 rows
            int totalCols = 3;  // B2 to D4 spans 3 columns

            // Populate each cell in the range with its own formula
            // Example formula: =ROW()+COLUMN()
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[firstRow + r, firstCol + c].Formula = "=ROW()+COLUMN()";
                }
            }

            // Merge the defined range
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(firstRow, firstCol, totalRows, totalCols);
            range.Merge();

            // Store the formulas of each cell before unmerging
            string[,] savedFormulas = new string[totalRows, totalCols];
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    savedFormulas[r, c] = cells[firstRow + r, firstCol + c].Formula;
                }
            }

            // Unmerge the range using the Range.UnMerge method
            range.UnMerge();

            // Restore the original formulas to each cell after unmerge
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    if (!string.IsNullOrEmpty(savedFormulas[r, c]))
                    {
                        cells[firstRow + r, firstCol + c].Formula = savedFormulas[r, c];
                    }
                }
            }

            // Save the workbook
            string outputPath = "UnmergedPreserveFormulas.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
