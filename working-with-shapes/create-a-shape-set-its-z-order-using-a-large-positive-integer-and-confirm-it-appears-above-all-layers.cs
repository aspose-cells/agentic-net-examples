// Title: Add a rectangle shape to an Excel worksheet and set a high ZOrderPosition to bring it to the front using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to bring shape to front | how to make a shape appear on top layer in Excel using Aspose.Cells | C# example for assigning a large Z-order value to a worksheet shape with Aspose.Cells | verify that a shape is the topmost object in an Aspose.Cells workbook
// Tags: Aspose.Cells set shape ZOrderPosition | C# add rectangle shape to Excel worksheet | Excel shape layering with Aspose.Cells | topmost shape verification Aspose.Cells | worksheet shape Z-order manipulation

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, assigns a ZOrderPosition of 1,000,000 to ensure the shape is rendered above all other objects, checks that it has the maximum Z-order among the sheet's shapes, and saves the file as ShapeZOrderDemo.xlsx.
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

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                100  // width (pixels)
            );

            // Set a large Z-order value to ensure the shape appears above all other layers
            shape.ZOrderPosition = 1_000_000;

            // Verify that this shape has the highest Z-order among all shapes on the sheet
            int maxZOrder = sheet.Shapes.Cast<Shape>().Max(s => s.ZOrderPosition);
            if (shape.ZOrderPosition == maxZOrder)
            {
                Console.WriteLine("The shape is on the top layer.");
            }
            else
            {
                Console.WriteLine("The shape is not on the top layer.");
            }

            // Save the workbook to a file
            workbook.Save("ShapeZOrderDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
