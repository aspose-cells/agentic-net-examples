// Title: Apply a solid blue background to every cell in an ODS worksheet using Aspose.Cells for .NET and save the file
// Common Searches: how to set a solid blue background for all cells in an ODS file using Aspose.Cells C# | Aspose.Cells C# apply background color to entire worksheet and export as ODS | C# code to save an ODS workbook with a page background using Aspose.Cells | apply a style to a large cell range in Aspose.Cells before ODS export | set foreground color to blue and solid pattern in Aspose.Cells workbook
// Tags: solid blue background ODS Aspose.Cells | bulk cell styling with StyleFlag All Aspose.Cells | export workbook to ODS with formatting Aspose.Cells | apply style to whole worksheet range C# | ForegroundColor blue pattern solid Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, defines a style with a solid blue foreground color, applies this style to a large range that covers the entire worksheet, and then saves the workbook as an ODS file preserving the blue background.
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

            // Create a style with solid blue background
            Style blueStyle = workbook.CreateStyle();
            blueStyle.ForegroundColor = Color.Blue;
            blueStyle.Pattern = BackgroundType.Solid;

            // Apply the style to a large range covering all cells (e.g., 1000 rows x 1000 columns)
            Aspose.Cells.Range allCells = sheet.Cells.CreateRange(0, 0, 1000, 1000);
            StyleFlag flag = new StyleFlag() { All = true };
            allCells.ApplyStyle(blueStyle, flag);

            // Save the workbook as ODS with the background applied
            workbook.Save("OdsWithBlueBackground.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
