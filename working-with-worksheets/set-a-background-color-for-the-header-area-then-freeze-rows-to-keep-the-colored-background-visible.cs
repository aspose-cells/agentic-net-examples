// Title: Apply a solid light‑blue background to the first worksheet row and freeze the header using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a style with a light‑blue solid fill, applies it to cells A1:D1, and then freezes the first row using Aspose.Cells. | Show how to use Workbook.CreateStyle, set ForegroundColor and Pattern, assign the style to a header range, and call Worksheet.FreezePanes to keep the header visible while scrolling. | Provide a complete example that saves the styled and frozen worksheet as Output.xlsx.
// Common Searches: Aspose.Cells C# set background color for header row and freeze panes | how to freeze the top row after applying a style with Aspose.Cells .NET | C# example for coloring cells A1 to D1 and freezing first row using Aspose.Cells
// Tags: apply solid fill to header row Aspose.Cells | freeze top row worksheet Aspose.Cells | create custom style workbook C# | set cell background color Excel .NET | freeze panes Excel using Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Creates a new workbook, defines a light‑blue solid style, applies it to cells A1‑D1 as a header, freezes the first row, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create workbook

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header area (first row, columns A to D)
            int headerRow = 0;          // zero‑based index for the first row
            int firstColumn = 0;        // column A
            int lastColumn = 3;         // column D (adjust as needed)

            // Create a style with a solid background color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;   // set desired background color
            headerStyle.Pattern = BackgroundType.Solid;      // apply solid fill

            // Apply the style to each cell in the header row
            for (int col = firstColumn; col <= lastColumn; col++)
            {
                Cell cell = sheet.Cells[headerRow, col];
                cell.PutValue($"Header{col + 1}"); // optional header text
                cell.SetStyle(headerStyle);
            }

            // Freeze the header row so it remains visible while scrolling
            // Freeze first row (row index 1) and no columns, freezing 1 row total
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
