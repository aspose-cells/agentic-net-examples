// Title: C# Example: Count Cells in Merged Range A1:C3 with Aspose.Cells for .NET
// Description: Creates a workbook, merges cells A1:C3, retrieves the merged range via GetMergedRange, multiplies RowCount by ColumnCount to obtain the total of 9 cells, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | merged range | cell count | GetMergedRange | RowCount | ColumnCount | A1:C3 | Excel automation | GitHub sample
// Common Searches: Aspose.Cells count cells in merged range | How to get merged range size C# Aspose.Cells | Calculate total cells of A1:C3 after merge | GetMergedRange RowCount ColumnCount example | Aspose.Cells merge cells and count
// Developer Intent: Learn how to determine the number of individual cells covered by a merged range using Aspose.Cells for .NET.
// Use Cases: Validate merged area dimensions before applying formatting | Adjust chart data ranges based on merged region size | Generate spreadsheet reports that include merged range statistics | Enforce size constraints for merged cells in automated Excel creation
// AI Prompts: Write C# code using Aspose.Cells to merge A1:C3, retrieve the merged range, and output the total cell count. | Show how to handle exceptions while calculating merged range cell count in Aspose.Cells for .NET. | Explain the use of RowCount and ColumnCount properties of a merged range to compute its total cells.

using System;
using Aspose.Cells;

// Creates a workbook, merges cells A1:C3, retrieves the merged range via GetMergedRange, multiplies RowCount by ColumnCount to obtain the total of 9 cells, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge the range A1:C3 (rows 0‑2, columns 0‑2)
            // firstRow = 0, firstColumn = 0, totalRows = 3, totalColumns = 3
            cells.Merge(0, 0, 3, 3);

            // Retrieve the merged range via the upper‑left cell (A1)
            Aspose.Cells.Range mergedRange = cells["A1"].GetMergedRange();

            // Calculate total number of cells in the merged range
            int totalCells = mergedRange.RowCount * mergedRange.ColumnCount;

            Console.WriteLine($"Total cells in merged range A1:C3: {totalCells}");

            // Save the workbook (lifecycle rule)
            workbook.Save("MergedRangeCellCount.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
