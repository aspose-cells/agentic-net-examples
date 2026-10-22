// Title: How to set a worksheet shape's Z-order to the default layer (0) using Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet, assign ZOrderPosition = 0, and check the property value with Aspose.Cells in C#. | Move an existing shape to the back layer by setting its ZOrderPosition to zero and programmatically verify the change. | Save the workbook after adjusting a shape's Z-order to the default layer and output a confirmation message.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to 0 default layer | How to move a shape to the back of an Excel worksheet using Aspose.Cells .NET | Verify shape Z-order after setting it in an Aspose.Cells workbook | C# example for shape layer ordering with Aspose.Cells | Set shape Z-order to bottom in generated Excel file using Aspose.Cells
// Tags: Aspose.Cells set shape ZOrderPosition | C# worksheet shape layer positioning | Aspose.Cells default shape Z-order | Excel shape back layer Aspose.Cells | validate shape Z-order property C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to 0 (the default layer), verifies the property, and saves the file as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                1,    // upper left column
                0,    // top offset (in points)
                0,    // left offset (in points)
                100,  // height (in points)
                50    // width (in points)
            );

            // Set the shape's Z-order to zero (default layer position)
            shape.ZOrderPosition = 0;

            // Verify that the Z-order is set to zero
            if (shape.ZOrderPosition == 0)
            {
                Console.WriteLine("Shape Z-order successfully set to the default layer (0).");
            }
            else
            {
                Console.WriteLine($"Z-order verification failed. Current Z-order: {shape.ZOrderPosition}");
            }

            // Save the workbook
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
