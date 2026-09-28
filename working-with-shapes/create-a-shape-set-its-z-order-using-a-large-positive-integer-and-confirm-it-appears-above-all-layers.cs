// Title: Add a rectangle shape in Aspose.Cells, assign a very high ZOrderPosition, and verify it is the topmost shape
// AI Prompts: Create a rectangle shape on a worksheet and set its ZOrderPosition to 1,000,000 so it appears above all other shapes. | Iterate over the worksheet's Shapes collection to determine the maximum ZOrderPosition and compare it with the rectangle's value. | Save the workbook and print a message indicating whether the rectangle shape is the topmost layer.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to a large value | How to bring a shape to the front in an Excel file using Aspose.Cells .NET | Determine which shape has the highest Z-order in an Aspose.Cells worksheet | Example of verifying topmost shape after adding multiple shapes with Aspose.Cells
// Tags: Aspose.Cells set shape ZOrderPosition | Aspose.Cells frontmost shape handling | Aspose.Cells iterate shape collection Z-order | Aspose.Cells save workbook with modified shapes | Aspose.Cells rectangle shape addition

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape with ZOrderPosition set to 1,000,000, adds a second shape with default ordering, scans all shapes to find the highest Z-order, confirms the rectangle is on top, and saves the file as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (returns a Shape object)
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1, 1,                     // upper left row, column
                0, 0,                     // row offset, column offset
                100, 100);                // height, width

            // Set a very large Z-order value so the shape appears above all others
            rectangle.ZOrderPosition = 1_000_000;

            // Add another rectangle shape with default Z-order for comparison
            Shape otherShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5, 5,
                0, 0,
                80, 80);

            // Verify that the first rectangle has the highest Z-order among all shapes
            int maxZOrder = int.MinValue;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.ZOrderPosition > maxZOrder)
                    maxZOrder = shape.ZOrderPosition;
            }

            if (rectangle.ZOrderPosition == maxZOrder)
                Console.WriteLine("The rectangle shape is on top of all layers.");
            else
                Console.WriteLine("The rectangle shape is NOT on top.");

            // Save the workbook to a file
            string outputPath = "ShapeZOrderDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
