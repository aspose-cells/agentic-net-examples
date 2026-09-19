// Title: Set a shape's Z-order to the default layer (zero) and confirm it with Aspose.Cells for .NET
// AI Prompts: Assign ZOrderPosition = 0 to a rectangle shape on a worksheet using Aspose.Cells in C#. | Read back the ZOrderPosition property after setting it and throw an exception if it is not zero. | Persist the workbook by saving it after the shape's Z-order has been modified.
// Common Searches: Aspose.Cells C# set shape Z-order to default layer | How to reset a shape's layering order to zero in an Excel file with Aspose.Cells | Verify shape ZOrderPosition after changing it using Aspose.Cells .NET API | Programmatically move a shape to the backmost layer in Excel via Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition | C# set shape layer order Excel | verify shape Z-order Aspose.Cells | default shape layer Aspose.Cells | save workbook after shape changes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, sets its ZOrderPosition to 0 (the default layer), validates the property, and saves the file as ShapeZOrderDemo.xlsx.
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
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Set the Z-order of the shape to zero (default layer position)
            shape.ZOrderPosition = 0;

            // Verify that the Z-order is set to zero
            if (shape.ZOrderPosition != 0)
            {
                throw new InvalidOperationException("Shape Z-order was not set to zero as expected.");
            }

            // Save the workbook to a file
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
