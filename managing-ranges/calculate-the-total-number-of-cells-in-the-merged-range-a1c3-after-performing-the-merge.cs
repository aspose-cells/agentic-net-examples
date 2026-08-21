// Title: Compute cell count of merged range A1:C3 using Aspose.Cells for .NET (C#)
// Description: A C# example that creates a workbook, merges the cells A1:C3 (3 rows × 3 columns), retrieves the merged range via the top‑left cell, and calculates the total number of cells by multiplying RowCount and ColumnCount, yielding 9 cells.
// Keywords: Aspose.Cells | merged range | cell count | C# | .NET | GetMergedRange | RowCount | ColumnCount | A1:C3 | calculate merged area size | Excel merge cells programmatically
// Common Searches: Aspose.Cells get merged range size | C# count cells in merged range A1:C3 | How many cells are in a merged area using Aspose.Cells | GetMergedRange RowCount ColumnCount example | Calculate merged area cell count .NET
// Developer Intent: Determine how many individual cells are covered by the merged range A1:C3.
// Use Cases: Verify that a merged region spans the expected number of cells before applying formatting or formulas. | Log merged area dimensions to aid debugging of complex spreadsheet layouts. | Drive conditional logic based on the size of a merged block, such as splitting data across worksheets.
// AI Prompts: Show how to retrieve RowCount and ColumnCount of a merged range and compute its total cells with Aspose.Cells for .NET. | Provide an alternative method to obtain the cell count of a merged region without manual multiplication. | Explain how to safely handle GetMergedRange when the target cell is not part of any merged area.

using System;
using Aspose.Cells;

// A C# example that creates a workbook, merges the cells A1:C3 (3 rows × 3 columns), retrieves the merged range via the top‑left cell, and calculates the total number of cells by multiplying RowCount and ColumnCount, yielding 9 cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge the range A1:C3 (rows 0-2, columns 0-2)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(0, 0, 3, 3);

            // Retrieve the merged range via the top‑left cell (A1)
            Aspose.Cells.Range mergedRange = cells["A1"].GetMergedRange();

            // Calculate the total number of cells in the merged range
            int totalCells = mergedRange.RowCount * mergedRange.ColumnCount;

            Console.WriteLine($"Total cells in merged range A1:C3: {totalCells}");

            // Save the workbook (optional)
            string outputPath = "MergedRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
