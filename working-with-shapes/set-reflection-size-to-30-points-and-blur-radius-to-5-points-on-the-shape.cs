// Title: How to set a shape's reflection size to 30 points and blur radius to 5 points in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet and apply a reflection effect with a size of 30 points and a blur of 5 points using Aspose.Cells in C#. | Configure the Reflection.Size and Reflection.Blur properties of a shape in an Aspose.Cells workbook to customize its visual appearance.
// Common Searches: Aspose.Cells C# set shape reflection size to 30 points | How to add blur to shape reflection in Aspose.Cells workbook | C# example for configuring shape reflection properties in Excel with Aspose.Cells | Set reflection blur amount on rectangle shape using Aspose.Cells .NET | Adjust shape visual effects (reflection size, blur) in Aspose.Cells
// Tags: Aspose.Cells shape reflection size property | Aspose.Cells shape reflection blur property | C# Aspose.Cells add rectangle shape | Aspose.Cells visual effects on Excel shapes | Excel shape reflection customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets the shape's reflection size to 30 points and blur radius to 5 points, and saves the workbook as output.xlsx.
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
                2,   // upper left row
                2,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                100, // height (in points)
                200  // width (in points)
            );

            // Set reflection size to 30 points
            shape.Reflection.Size = 30;

            // Set blur amount (the correct property name) to 5 points
            shape.Reflection.Blur = 5;

            // Determine output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
