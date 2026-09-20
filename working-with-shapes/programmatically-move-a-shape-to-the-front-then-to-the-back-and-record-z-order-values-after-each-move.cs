// Title: How to programmatically change a rectangle shape's Z-order (front and back) and read its ZOrderPosition using Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet, increase its ZOrderPosition to move it to the front, output the Z-order value, then decrease the ZOrderPosition to move it to the back and output the new value. | Using Aspose.Cells for .NET, adjust a shape's ZOrderPosition property to reorder the shape, printing the Z-order before and after each adjustment, and save the workbook.
// Common Searches: Aspose.Cells C# how to bring a shape to the front of an Excel sheet | C# set ZOrderPosition for a shape in Aspose.Cells workbook | retrieve Z-order value of a rectangle shape using Aspose.Cells .NET | move shape to back programmatically with Aspose.Cells drawing API | example of changing shape Z-order and saving workbook in C#
// Tags: Aspose.Cells shape Z-order manipulation | set shape ZOrderPosition .NET | bring shape to front Excel Aspose.Cells | send shape to back Aspose.Cells | read shape Z-order value C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, adds a rectangle shape, displays its initial ZOrderPosition, increments the position to bring the shape forward and prints the value, then decrements it to send the shape back and prints the final value, and finally saves the workbook as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Display the initial Z-order position
            Console.WriteLine("Initial Z-order: " + shape.ZOrderPosition);

            // Change Z-order manually (higher value brings the shape to front)
            shape.ZOrderPosition = shape.ZOrderPosition + 1;
            Console.WriteLine("After manual Z-order increment: " + shape.ZOrderPosition);

            // Change Z-order back (lower value sends the shape to back)
            shape.ZOrderPosition = shape.ZOrderPosition - 1;
            Console.WriteLine("After manual Z-order decrement: " + shape.ZOrderPosition);

            // Save the workbook to a file
            string outputPath = "ShapeZOrderDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
