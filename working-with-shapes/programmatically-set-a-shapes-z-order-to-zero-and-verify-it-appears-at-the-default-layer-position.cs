// Title: Set a shape's Z-order to the default layer (zero) and verify it with Aspose.Cells for .NET (C#)
// Common Searches: how to set a shape ZOrderPosition to 0 using Aspose.Cells C# | Aspose.Cells C# verify shape Z-order after adding rectangle | reset shape layering to default in an Excel worksheet with Aspose.Cells | C# code to check shape ZOrderPosition value in a workbook | Aspose.Cells set shape order to bottom layer programmatically
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# set shape Z-order in Excel workbook | default layer positioning for Excel shapes | verify shape ordering with Aspose.Cells API | add rectangle shape and adjust Z-order .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to 0 to place it on the default layer, confirms the property value, and saves the workbook as an .xlsx file.
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
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 40);

            // Set the Z-order of the shape to zero (default layer position)
            shape.ZOrderPosition = 0;

            // Verify that the Z-order is set to zero
            if (shape.ZOrderPosition == 0)
            {
                Console.WriteLine("Shape Z-order successfully set to zero (default layer).");
            }
            else
            {
                Console.WriteLine($"Unexpected Z-order value: {shape.ZOrderPosition}");
            }

            // Define output file path
            string outputPath = "ShapeZOrderDemo.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
