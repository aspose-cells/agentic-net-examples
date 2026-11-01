// Title: Create and apply a custom font style (Arial, size 12) to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Arial 12 font for cells A1 to C3 | apply a style with specific font name and size to a range in Aspose.Cells .NET | using StyleFlag to change only font attributes in an Aspose.Cells workbook | how to create and reuse a custom style for multiple cells with Aspose.Cells C#
// Tags: create custom font style Aspose.Cells C# | apply style to cell range Aspose.Cells | StyleFlag font name size Aspose.Cells | range.ApplyStyle method example | set Arial font size 12 workbook Aspose.Cells

using Aspose.Cells;
using System;

// The example demonstrates how to create a Style object, set its font family to Arial and size to 12, configure a StyleFlag to apply only those font attributes, and then apply the style to the range A1:C3 on the first worksheet of a new workbook using Aspose.Cells for .NET. The workbook is saved as StyledWorkbook.xlsx, with exception handling included.
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

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set the desired font family and size
            style.Font.Name = "Arial";
            style.Font.Size = 12;

            // Define which style attributes to apply (font name and size)
            StyleFlag flag = new StyleFlag
            {
                FontName = true,
                FontSize = true
            };

            // Apply the style to a range of cells (e.g., A1 to C3)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C3");
            range.ApplyStyle(style, flag);

            // Save the workbook to a file
            workbook.Save("StyledWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
