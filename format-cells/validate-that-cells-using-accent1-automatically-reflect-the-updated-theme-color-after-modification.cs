// Title: Apply a solid fill style to a cell and verify its foreground color using Aspose.Cells for .NET
// Common Searches: how to apply a solid fill color to a specific cell with Aspose.Cells in C# | retrieve and check the foreground color of a styled cell using Aspose.Cells .NET | Aspose.Cells example for setting and reading back a custom cell background color | C# code to validate that a cell style color was applied correctly with Aspose.Cells | read cell style properties after applying solid fill in Aspose.Cells workbook
// Tags: apply solid fill style Aspose.Cells C# | read cell foreground color Aspose.Cells | validate custom cell background color .NET | Aspose.Cells cell style verification | programmatic cell color extraction Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, defines a style with a Lime foreground color and solid fill, applies the style to cell A1, then retrieves the cell's style to read back the effective foreground color and prints it.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Define a style with a solid fill and a specific foreground color
            Style style = wb.CreateStyle();
            style.ForegroundColor = Color.Lime;          // Set desired color directly
            style.Pattern = BackgroundType.Solid;        // Apply solid fill to make the color visible

            // Apply the style to cell A1
            Cell cell = ws.Cells["A1"];
            cell.SetStyle(style);

            // Retrieve the cell's style to verify the applied color
            Style appliedStyle = cell.GetStyle();
            Color effectiveColor = appliedStyle.ForegroundColor;

            // Output validation results
            Console.WriteLine("Cell A1 foreground color: " + effectiveColor);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
