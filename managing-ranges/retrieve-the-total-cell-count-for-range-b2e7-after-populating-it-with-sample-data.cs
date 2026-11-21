// Title: Calculate total cells in the B2:E7 range after filling it with data using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get total cell count for a specific range | C# count cells in Excel range B2:E7 using Aspose.Cells | How to determine number of cells in a populated range with Aspose.Cells for .NET | Retrieve row and column count of range B2:E7 in Aspose.Cells C# example
// Tags: Aspose.Cells calculate range cell count | C# Aspose.Cells CreateRange example | Excel range RowCount ColumnCount with Aspose.Cells | populate Excel range B2:E7 using Aspose.Cells | total cells in specific Excel range Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, fills cells B2 through E7 with incremental numbers, defines the area with CreateRange, calculates the total number of cells by multiplying RowCount and ColumnCount, and writes the result to the console.
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

            // Populate range B2:E7 with incremental numbers
            int startRow = 1; // Row index for B2 (zero‑based)
            int startCol = 1; // Column index for B2 (zero‑based)
            int value = 1;

            for (int row = startRow; row <= 6; row++)          // Rows 2 to 7
            {
                for (int col = startCol; col <= 4; col++)      // Columns B to E
                {
                    sheet.Cells[row, col].PutValue(value++);
                }
            }

            // Define the range B2:E7 using Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:E7");

            // Retrieve total cell count in the range
            int totalCellCount = range.RowCount * range.ColumnCount;

            // Output the result
            Console.WriteLine($"Total cell count in range B2:E7: {totalCellCount}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
