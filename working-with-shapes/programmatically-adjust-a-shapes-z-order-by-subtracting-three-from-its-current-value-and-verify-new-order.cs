// Title: Decrease a shape's Z-order by three positions and save the Excel workbook with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells reduce shape ZOrderPosition by three | how to move an Excel shape backward three layers using Aspose.Cells | ensure shape Z-order does not become negative in Aspose.Cells C# | retrieve and modify Z-order of a rectangle shape in Aspose.Cells | save workbook after adjusting shape Z-order with Aspose.Cells .NET
// Tags: adjust shape ZOrderPosition Aspose.Cells | decrease shape Z-order C# | prevent negative Z-order Aspose.Cells | save workbook after shape order change Aspose.Cells | rectangle shape Z-order manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, sets its initial ZOrderPosition to 5, computes a new position by subtracting three while ensuring it does not drop below zero, updates the shape, outputs the previous and adjusted Z-order values, and saves the workbook as ZOrderAdjusted.xlsx.
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
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                100, // height (in points)
                100  // width (in points)
            );

            // Set an initial Z-order position for demonstration (e.g., 5)
            shape.ZOrderPosition = 5;

            // Store the current Z-order
            int currentZOrder = shape.ZOrderPosition;

            // Subtract three from the current Z-order, ensuring it does not go below zero
            int newZOrder = Math.Max(0, currentZOrder - 3);
            shape.ZOrderPosition = newZOrder;

            // Output the Z-order values
            Console.WriteLine($"Previous Z-order: {currentZOrder}");
            Console.WriteLine($"Adjusted Z-order: {shape.ZOrderPosition}");

            // Define output file path
            string outputPath = "ZOrderAdjusted.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (the shape persists in the file)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
