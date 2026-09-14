// Title: How to compute the total number of cells in a populated B2:E7 range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that fills the worksheet range B2:E7 with sample data and returns the count of cells in that range using Aspose.Cells. | Show how to create a Range object for B2:E7 with Aspose.Cells, then obtain its RowCount and ColumnCount to calculate the total cells. | Provide a C# snippet that writes values into B2:E7, accesses the range via CreateRange, and prints the total cell count.
// Common Searches: Aspose.Cells C# count cells in range B2:E7 after writing data | How to get total cells of a specific Excel range using Aspose.Cells .NET | C# example for retrieving RowCount and ColumnCount of a range with Aspose.Cells
// Tags: Aspose.Cells calculate range cell count | CreateRange method Aspose.Cells C# | populate Excel range with sample data Aspose.Cells | retrieve RowCount ColumnCount Aspose.Cells | C# total cells in Excel range

using Aspose.Cells;
using System;

// The example creates a new workbook, populates cells B2:E7 with sample strings, creates a Range object for that area using CreateRange, calculates the total number of cells by multiplying RowCount and ColumnCount, and prints the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B2:E7 (zero‑based indices)
            int startRow = 1;      // B2 row index
            int startColumn = 1;   // B2 column index
            int totalRows = 6;     // rows 2 through 7 inclusive
            int totalColumns = 4;  // columns B through E inclusive

            // Populate the range with sample data
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    sheet.Cells[startRow + i, startColumn + j].PutValue($"R{i + 2}C{j + 2}");
                }
            }

            // Create a Range object for B2:E7 using fully qualified name to avoid ambiguity
            Aspose.Cells.Range range = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Retrieve total cell count in the range
            int cellCount = range.RowCount * range.ColumnCount;

            // Output the result
            Console.WriteLine($"Total cell count in range B2:E7: {cellCount}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
