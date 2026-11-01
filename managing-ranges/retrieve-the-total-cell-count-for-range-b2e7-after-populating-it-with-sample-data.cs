// Title: Calculate total number of cells in range B2:E7 after populating it with data using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get total cells in a specific range after filling data | C# count cells in Excel range B2:E7 with Aspose.Cells API | retrieve row and column count of a range using Aspose.Cells .NET | determine number of cells in a populated range with Aspose.Cells C# | how to use Aspose.Cells Range.RowCount and ColumnCount to compute cell count
// Tags: Aspose.Cells range cell count calculation | C# Aspose.Cells RowCount ColumnCount usage | populate Excel range B2:E7 with Aspose.Cells | Aspose.Cells total cells in a defined range | Excel range dimensions Aspose.Cells .NET

using Aspose.Cells;
using System;

// The example creates a new workbook, fills cells B2 through E7 with sequential numbers, defines that area as an Aspose.Cells.Range, calculates the total cell count by multiplying RowCount and ColumnCount (or using range.Cells.Count), and outputs the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate range B2:E7 with sample data (sequential numbers)
            int value = 1;
            for (int row = 1; row <= 6; row++) // rows 2‑7 (0‑based index)
            {
                for (int col = 1; col <= 4; col++) // columns B‑E (0‑based index)
                {
                    sheet.Cells[row, col].PutValue(value++);
                }
            }

            // Define the range B2:E7 using Aspose.Cells.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("B2", "E7");

            // Retrieve total cell count in the range
            int totalCells = range.RowCount * range.ColumnCount; // or range.Cells.Count

            // Output the result
            Console.WriteLine($"Total cell count in range B2:E7: {totalCells}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
