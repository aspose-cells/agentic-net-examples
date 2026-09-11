// Title: Move a worksheet shape to the front, capture its ZOrderPosition, then send it to the back with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that adds a rectangle shape to a worksheet, reads its ZOrderPosition, sets the shape to the highest Z-order, then resets it to the lowest Z-order and prints each value. | Show how to programmatically bring a shape to the front and send it to the back in an Excel file using Aspose.Cells, including verification of Z-order changes. | Create a sample that saves an Excel workbook after modifying a shape's ZOrderPosition, demonstrating both BringToFront and SendToBack operations.
// Common Searches: aspnet aspose.cells set shape ZOrderPosition to front | c# retrieve shape Z-order index in Excel workbook using Aspose.Cells | programmatic method for moving an Excel shape behind other objects with Aspose.Cells | step-by-step guide to adjust shape layering order in Aspose.Cells .NET | Aspose.Cells shape ordering Z-order manipulation tutorial
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# bring shape to front Aspose.Cells | C# send shape to back Aspose.Cells | Aspose.Cells worksheet shape ordering | Excel shape Z-order handling with Aspose.Cells | Aspose.Cells rectangle shape example

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, adds a rectangle shape, records its initial ZOrderPosition, moves the shape to the front by assigning the highest Z-order index, then moves it to the back by assigning zero, prints the original, front, and back Z-order values, validates the ordering logic, and saves the workbook as ShapeZOrderDemo.xlsx.
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
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 150, 80);

            // Capture the original Z-order index
            int originalZOrder = shape.ZOrderPosition;

            // Bring the shape to the front by setting its Z-order to the highest index
            shape.ZOrderPosition = sheet.Shapes.Count - 1;
            int frontZOrder = shape.ZOrderPosition;

            // Send the shape to the back by setting its Z-order to the lowest index
            shape.ZOrderPosition = 0;
            int backZOrder = shape.ZOrderPosition;

            // Output the Z-order values
            Console.WriteLine($"Original Z-order: {originalZOrder}");
            Console.WriteLine($"After BringToFront Z-order: {frontZOrder}");
            Console.WriteLine($"After SendToBack Z-order: {backZOrder}");

            // Verify expected behavior
            if (frontZOrder > originalZOrder && backZOrder < frontZOrder)
            {
                Console.WriteLine("Z-order changes as expected.");
            }
            else
            {
                Console.WriteLine("Unexpected Z-order behavior.");
            }

            // Ensure the output directory exists
            string outputPath = "ShapeZOrderDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
