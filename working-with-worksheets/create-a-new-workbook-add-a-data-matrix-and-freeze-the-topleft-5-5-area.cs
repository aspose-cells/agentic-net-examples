// Title: Create a new workbook, fill a 10×10 data matrix, and freeze the top‑left 5×5 pane using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, write a 10×10 sample matrix into the first worksheet, and apply the pane‑freezing API to lock the first five rows and columns. | Demonstrate how to invoke the pane‑freezing method with row, column, totalRows, and totalColumns arguments to freeze a 5×5 area in an Excel file produced by Aspose.Cells.
// Common Searches: Aspose.Cells C# example to freeze the first five rows and columns in an Excel sheet | How to populate a 10x10 data range and set frozen panes with Aspose.Cells for .NET | C# code for creating a workbook, adding a sample data matrix, and freezing the top‑left block using Aspose.Cells | FreezePanes method usage Aspose.Cells .NET 5x5 area | Generate Excel file with sample matrix and frozen pane in C# Aspose.Cells
// Tags: Aspose.Cells pane freezing C# | 10x10 data matrix Aspose.Cells | top-left frozen pane Excel .NET | sample workbook generation Aspose.Cells | Excel frozen pane 5x5 Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new workbook, fills cells A1:J10 with values like R1C1, freezes the top‑left 5×5 area (rows 1‑5, columns A‑E) using the pane‑freezing API, and saves the file as FrozenTopLeft5x5.xlsx.
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

            // Add a 10x10 data matrix with sample values
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top‑left 5 × 5 area (rows 0‑4 and columns 0‑4)
            // row = first unfrozen row index, column = first unfrozen column index,
            // totalRows = number of rows to freeze, totalColumns = number of columns to freeze
            sheet.FreezePanes(5, 5, 5, 5);

            // Save the workbook
            workbook.Save("FrozenTopLeft5x5.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
