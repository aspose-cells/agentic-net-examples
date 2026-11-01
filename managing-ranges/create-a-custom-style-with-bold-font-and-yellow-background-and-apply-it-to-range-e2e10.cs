// Title: How to create a bold font with yellow background style and apply it to cells E2:E10 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply bold text and yellow fill to a range of cells | Create and use a custom style for cells E2 to E10 in Aspose.Cells .NET | Set font bold and background color for a specific range with Aspose.Cells API | Using StyleFlag to apply font and shading to a cell range in Aspose.Cells C# | How to style a column range with solid yellow background in Aspose.Cells
// Tags: custom style creation Aspose.Cells .NET | apply style to specific range Aspose.Cells | font boldness with background color Aspose.Cells | StyleFlag configuration for cell shading Aspose.Cells | solid yellow background using Aspose.Cells style

using System;
using Aspose.Cells;
using System.Drawing;

// The example demonstrates creating a custom style with a bold font and yellow solid background, configuring a StyleFlag for font boldness and cell shading, applying the style to the range E2:E10, and saving the workbook as StyledWorkbook.xlsx.
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

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsBold = true;                     // Bold font
            customStyle.ForegroundColor = Color.Yellow;        // Yellow background
            customStyle.Pattern = BackgroundType.Solid;        // Solid fill

            // Apply the style to the range E2:E10
            Aspose.Cells.Range range = sheet.Cells.CreateRange("E2", "E10");
            range.ApplyStyle(customStyle, new StyleFlag
            {
                FontBold = true,
                CellShading = true
            });

            // Save the workbook
            workbook.Save("StyledWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
