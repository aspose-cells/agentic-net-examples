// Title: Apply a diagonal crosshatch fill pattern to placeholder cells in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set diagonal crosshatch background for cells containing {{placeholder}} | apply custom fill pattern to placeholder text cells in Excel using Aspose.Cells .NET | style cells with BackgroundType.DiagonalCrosshatch based on string pattern in C# | highlight template variables in worksheet with crosshatch fill using Aspose.Cells | C# iterate cells and apply style when value includes double curly braces Aspose.Cells
// Tags: crosshatch background Aspose.Cells | placeholder cell styling .NET | BackgroundType.DiagonalCrosshatch usage | conditional cell style based on string content | Excel workbook custom fill pattern C#

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook, inserts placeholder strings like {{Name}} and {{Date}}, defines a style with BackgroundType.DiagonalCrosshatch and specific foreground/background colors, loops through all cells, applies the style to any cell whose text contains the {{…}} pattern, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example placeholder cells
            sheet.Cells["A1"].PutValue("{{Name}}");
            sheet.Cells["B2"].PutValue("{{Date}}");
            sheet.Cells["C3"].PutValue("Regular text");

            // Define a style with diagonal crosshatch fill pattern
            Style crosshatchStyle = workbook.CreateStyle();
            // Use the correct enum value for diagonal crosshatch pattern
            crosshatchStyle.Pattern = BackgroundType.DiagonalCrosshatch;
            crosshatchStyle.ForegroundColor = Color.LightGray;   // Color of the crosshatch lines
            crosshatchStyle.BackgroundColor = Color.White;      // Background color

            // Apply the style to cells that contain placeholder text (e.g., {{...}})
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Type == CellValueType.IsString &&
                    cell.StringValue.Contains("{{") &&
                    cell.StringValue.Contains("}}"))
                {
                    cell.SetStyle(crosshatchStyle);
                }
            }

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
