// Title: Count distinct cells in a specified range that contains merged cells using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that calculates the number of unique cells in a given range, treating each merged block as a single cell. | Implement logic to adjust the cell count for merged areas intersecting the range and output a verification message comparing the result with the expected count.
// Common Searches: Aspose.Cells .NET how to get cell count of a range with merged cells | C# count unique cells in worksheet range when cells are merged | calculate distinct cell number in merged area using Aspose.Cells | verify merged cell count in Aspose.Cells range A1:D3 | determine cell count after merging cells with Aspose.Cells API
// Tags: Aspose.Cells merged cells counting logic | C# distinct cell count in range | range cell counting with merged areas | verify merged area cell count Aspose.Cells | cell area intersection handling Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, merges cells A1:C2, defines the range A1:D3, computes the distinct cell count by treating the merged block as a single cell, and verifies that the calculated count matches the expected value.
class MergedCellCountExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells A1:C2 (covers 2 rows x 3 columns)
            sheet.Cells.Merge(0, 0, 2, 3); // Row index, column index, total rows, total columns

            // Define a range that includes the merged area, e.g., A1:D3
            AsposeRange range = sheet.Cells.CreateRange("A1:D3");

            // Initial cell count (total rows * total columns in the range)
            int cellCount = range.RowCount * range.ColumnCount;

            // Adjust count for merged cells (each merged block counts as one cell)
            int rangeFirstRow = range.FirstRow;
            int rangeFirstColumn = range.FirstColumn;
            int rangeLastRow = rangeFirstRow + range.RowCount - 1;
            int rangeLastColumn = rangeFirstColumn + range.ColumnCount - 1;

            foreach (CellArea merged in sheet.Cells.MergedCells)
            {
                // Determine if the merged area intersects the defined range
                bool intersect = !(merged.EndRow < rangeFirstRow ||
                                   merged.StartRow > rangeLastRow ||
                                   merged.EndColumn < rangeFirstColumn ||
                                   merged.StartColumn > rangeLastColumn);

                if (intersect)
                {
                    int mergedRows = merged.EndRow - merged.StartRow + 1;
                    int mergedCols = merged.EndColumn - merged.StartColumn + 1;
                    int mergedCellCount = mergedRows * mergedCols;

                    // Subtract the extra cells (mergedCellCount - 1) because they count as a single cell
                    cellCount -= (mergedCellCount - 1);
                }
            }

            // Expected count:
            // Without merge: 3 rows * 4 columns = 12 cells
            // Merged area (A1:C2) occupies 6 cells but counts as 1, so subtract 5
            int expectedCount = 12 - 5; // = 7

            // Output the results
            Console.WriteLine($"Cell count in range A1:D3 (including merged cells): {cellCount}");
            Console.WriteLine($"Expected cell count: {expectedCount}");

            // Verify that the count reflects the merged area
            if (cellCount == expectedCount)
            {
                Console.WriteLine("Verification passed: Cell count reflects the merged area.");
            }
            else
            {
                Console.WriteLine("Verification failed: Cell count does not reflect the merged area.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
