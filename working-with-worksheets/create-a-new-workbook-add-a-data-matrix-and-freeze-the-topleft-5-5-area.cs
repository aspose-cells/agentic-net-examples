// Title: How to create a new workbook, fill a 10x10 matrix, and freeze the top‑left 5×5 area using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to populate a 10x10 range with calculated values | Freeze the first five rows and columns in an Excel file with Aspose.Cells .NET | Using FreezePanes to lock a 5x5 area in a worksheet created with Aspose.Cells | Save a workbook after filling cells programmatically with Aspose.Cells in C# | How to set up a data matrix and freeze panes in Aspose.Cells for .NET
// Tags: Aspose.Cells create workbook C# | populate worksheet matrix Aspose.Cells | FreezePanes top-left area Aspose.Cells | save workbook as xlsx Aspose.Cells | initialize sample data cells Aspose.Cells

using Aspose.Cells;
using System;

// The sample program creates a new Workbook, fills cells A1:J10 with the product of their 1‑based row and column indices, applies FreezePanes to lock rows 0‑4 and columns 0‑4 (the top‑left 5×5 area), and saves the result as output.xlsx, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a 10x10 matrix with sample values
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    // Fill each cell with (row+1)*(col+1) for demonstration
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Freeze the top‑left 5 × 5 area (rows 0‑4 and columns 0‑4)
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(5, 5, 5, 5);

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
