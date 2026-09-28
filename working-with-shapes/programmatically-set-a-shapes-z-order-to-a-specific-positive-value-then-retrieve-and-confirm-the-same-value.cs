// Title: Set a specific Z-order value for a shape and verify it in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet, assign a positive ZOrderPosition, then read the property back to confirm the ordering using Aspose.Cells in C#. | Programmatically bring a shape to the front by setting its Z-order and immediately retrieve the ZOrderPosition to validate the change with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells set shape ZOrderPosition to bring shape to front | how to read ZOrderPosition of a shape after setting it in Aspose.Cells | Aspose.Cells example for changing drawing object Z-order in an Excel file | retrieve shape Z-order value using Aspose.Cells .NET API | Aspose.Cells code sample for setting and verifying shape Z-order
// Tags: Aspose.Cells set shape ZOrderPosition | Aspose.Cells read shape ZOrderPosition | Aspose.Cells shape Z-order manipulation | C# add rectangle shape Aspose.Cells | Excel shape ordering Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new Workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to a specific positive integer to bring it to the front, reads the property back to confirm the value, saves the workbook, and includes basic exception handling.
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

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                5,   // top offset (pixels)
                5,   // left offset (pixels)
                100, // height (pixels)
                200  // width (pixels)
            );

            // Set Z-order to bring the shape to the front (higher value = front)
            shape.ZOrderPosition = 1;

            // Output a confirmation message
            Console.WriteLine("Shape added and brought to front.");

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputFile = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
