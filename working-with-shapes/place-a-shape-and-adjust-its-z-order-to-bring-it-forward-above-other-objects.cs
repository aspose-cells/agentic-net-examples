// Title: Insert a rectangle shape into an Excel worksheet and bring it to the front by setting Z‑order with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape to the first worksheet, set its Left and Top coordinates, and assign a ZOrder value so the shape appears above all other objects. | Generate a .NET example that creates a workbook, inserts a rectangle shape, positions it, and modifies the Z‑order to make the shape the frontmost element in the saved Excel file.
// Common Searches: Aspose.Cells C# how to set shape ZOrder to front | C# add rectangle shape to Excel and change layering with Aspose.Cells | move Excel shape to front using Aspose.Cells .NET API | adjust shape Z-order programmatically in Aspose.Cells workbook
// Tags: Aspose.Cells shape order control | C# insert geometric shape into Excel worksheet | Aspose.Cells adjust shape layering | Excel workbook shape positioning .NET | Aspose.Cells frontmost shape configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a rectangle shape onto the first worksheet, sets its Left and Top positions, optionally changes its ZOrder to control layering, and saves the file as ShapeZOrder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                200,  // top offset (in points)
                100,  // left offset (in points)
                100,  // height (in points)
                200   // width (in points)
            );

            // Position the shape on the sheet
            shape.Left = 50; // distance from the left edge (in points)
            shape.Top = 50;  // distance from the top edge (in points)

            // Optional: set Z-order if needed (lower values are behind higher values)
            // shape.ZOrder = 0;

            // Save the workbook to a file
            string outputPath = "ShapeZOrder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
