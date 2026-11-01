// Title: Add rectangle shapes and control their Z-order layering in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate a C# program that adds three rectangle shapes to a worksheet and assigns ZOrderPosition values 0, 1, and 2 with Aspose.Cells. | Write C# code to create an Excel file, insert rectangle shapes at different cell positions, and adjust their Z-order to illustrate stacking order using Aspose.Cells.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition example | how to change shape stacking order in Excel using Aspose.Cells .NET | C# code to add multiple shapes with different Z-order in an Excel workbook | visualize shape layering in Excel with Aspose.Cells Z-order property
// Tags: add rectangle shapes Aspose.Cells C# | set ZOrderPosition property Aspose.Cells | shape layering Excel Aspose.Cells | Z-order stacking Excel shapes .NET | create workbook with layered shapes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds three rectangle shapes to the first worksheet, sets their ZOrderPosition to 0, 1, and 2 to define bottom, middle, and top layers, and saves the file as ZOrderDemo.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (bottom layer) and set its Z-order position
            Shape bottomShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2, 2,                     // upper left row & column
                0, 0,                     // top & left offset (in pixels)
                100, 200);                // height & width (in pixels)
            bottomShape.Name = "BottomShape";
            bottomShape.ZOrderPosition = 0; // lowest Z-order

            // Add a rectangle shape (middle layer) and set its Z-order position
            Shape middleShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                3, 3,
                0, 0,
                100, 200);
            middleShape.Name = "MiddleShape";
            middleShape.ZOrderPosition = 1; // above BottomShape

            // Add a rectangle shape (top layer) and set its Z-order position
            Shape topShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                4, 4,
                0, 0,
                100, 200);
            topShape.Name = "TopShape";
            topShape.ZOrderPosition = 2; // above MiddleShape

            // Save the workbook to observe the layering effect
            workbook.Save("ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
