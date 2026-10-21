// Title: Calculate the total number of cells in a merged A1:C3 range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# count cells in merged range A1:C3 | how to determine size of merged area with Aspose.Cells Merge method | C# Aspose.Cells get RowCount and ColumnCount of a merged range | example of calculating total cells after merging cells in Aspose.Cells workbook | retrieve total cell count from merged range in Aspose.Cells .NET API
// Tags: Aspose.Cells merge range cell count | C# Aspose.Cells RowCount ColumnCount calculation | merged range size Aspose.Cells .NET | Aspose.Cells CreateRange usage | total cells in merged area C#

using Aspose.Cells;
using System;

// The example creates a new workbook, merges cells A1:C3 on the first worksheet, builds a Range object for the merged area, computes the total number of cells by multiplying RowCount and ColumnCount, and outputs the result to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge (A1:C3)
            int startRow = 0;      // Row 1 (zero‑based index)
            int startColumn = 0;   // Column A
            int totalRows = 3;     // Rows 1 to 3
            int totalColumns = 3;  // Columns A to C

            // Perform the merge operation
            sheet.Cells.Merge(startRow, startColumn, totalRows, totalColumns);

            // Create a Range object representing the merged area
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Calculate the total number of cells in the merged range
            int totalCells = mergedRange.RowCount * mergedRange.ColumnCount;

            // Output the result
            Console.WriteLine("Total cells in merged range A1:C3: " + totalCells);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
