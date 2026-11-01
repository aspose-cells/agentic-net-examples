// Title: Add a rectangle shape with custom dimensions to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape to an Excel worksheet with Aspose.Cells in C# | Aspose.Cells C# set shape height and width on a worksheet | create a glass‑pane style rectangle shape in Excel using Aspose.Cells | C# example for adding shapes to a workbook with Aspose.Cells | Aspose.Cells add shape at specific row and column in C#
// Tags: Aspose.Cells add rectangle shape | C# set shape dimensions Aspose.Cells | Excel worksheet shape creation Aspose.Cells | Aspose.Cells shape size parameters | C# add shape to workbook Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a rectangle shape sized 300 × 150 pixels at row 1, column 1 of the first worksheet, and saves the file as GlassReflection.xlsx.
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

            // Add a rectangle shape that will act as the glass pane
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape glassShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                300, // height (pixels)
                150  // width (pixels)
            );

            // Save the workbook
            workbook.Save("GlassReflection.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
