using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsStyleComparison
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook (contains all styles)
                string sourcePath = "SourceWorkbook.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the original workbook (contains all styles)
                Workbook wbOriginal = new Workbook(sourcePath);

                // Load a second instance of the same workbook that we will clean up
                Workbook wbCleaned = new Workbook(sourcePath);

                // Remove all unused styles from the second workbook
                wbCleaned.RemoveUnusedStyles();

                // Display style pool counts for both workbooks
                Console.WriteLine($"Original workbook style pool count: {wbOriginal.CountOfStylesInPool}");
                Console.WriteLine($"Cleaned workbook style pool count:   {wbCleaned.CountOfStylesInPool}");

                // Compare visual consistency cell by cell
                bool allCellsMatch = true;

                // Ensure both workbooks have the same number of worksheets
                if (wbOriginal.Worksheets.Count != wbCleaned.Worksheets.Count)
                {
                    Console.WriteLine("Worksheet count mismatch between the two workbooks.");
                    allCellsMatch = false;
                }
                else
                {
                    for (int sheetIndex = 0; sheetIndex < wbOriginal.Worksheets.Count; sheetIndex++)
                    {
                        Worksheet sheetOriginal = wbOriginal.Worksheets[sheetIndex];
                        Worksheet sheetCleaned = wbCleaned.Worksheets[sheetIndex];

                        // Determine the used ranges for both sheets
                        AsposeRange originalRange = sheetOriginal.Cells.MaxDisplayRange;
                        AsposeRange cleanedRange = sheetCleaned.Cells.MaxDisplayRange;

                        // If a sheet is empty, skip it
                        if (originalRange == null && cleanedRange == null)
                            continue;

                        // Use empty range defaults when one side is null
                        if (originalRange == null) originalRange = cleanedRange;
                        if (cleanedRange == null) cleanedRange = originalRange;

                        // Compute the union of the two ranges
                        int startRow = Math.Min(originalRange.FirstRow, cleanedRange.FirstRow);
                        int startCol = Math.Min(originalRange.FirstColumn, cleanedRange.FirstColumn);
                        int endRow = Math.Max(originalRange.FirstRow + originalRange.RowCount - 1,
                                             cleanedRange.FirstRow + cleanedRange.RowCount - 1);
                        int endCol = Math.Max(originalRange.FirstColumn + originalRange.ColumnCount - 1,
                                             cleanedRange.FirstColumn + cleanedRange.ColumnCount - 1);

                        // Iterate through the union area
                        for (int row = startRow; row <= endRow; row++)
                        {
                            for (int col = startCol; col <= endCol; col++)
                            {
                                // Get display styles (styles after considering conditional formatting, merged cells, etc.)
                                Style styleOriginal = sheetOriginal.Cells[row, col].GetDisplayStyle();
                                Style styleCleaned = sheetCleaned.Cells[row, col].GetDisplayStyle();

                                // Compare the two styles
                                if (!styleOriginal.Equals(styleCleaned))
                                {
                                    Console.WriteLine($"Style mismatch at Sheet[{sheetIndex}] Cell[{row},{col}]");
                                    allCellsMatch = false;
                                }
                            }
                        }
                    }
                }

                // Report the final result
                if (allCellsMatch)
                    Console.WriteLine("Visual consistency verified: all cell styles match after removing unused styles.");
                else
                    Console.WriteLine("Visual inconsistency detected: some cell styles differ after removing unused styles.");

                // Optionally, save the cleaned workbook for manual inspection
                string cleanedPath = "CleanedWorkbook.xlsx";
                wbCleaned.Save(cleanedPath);
                Console.WriteLine($"Cleaned workbook saved to: {cleanedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}