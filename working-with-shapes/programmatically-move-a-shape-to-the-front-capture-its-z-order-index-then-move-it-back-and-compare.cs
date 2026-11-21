// Title: Move an Excel shape to the front or back and read its Z-order index using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape to front of worksheet | C# retrieve ZOrderPosition of a shape in an Excel file with Aspose.Cells | How to send a shape to the back in an Aspose.Cells workbook | Programmatically change shape layering order using Aspose.Cells .NET API | Example of moving Excel shape Z-order with Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition manipulation | bring shape to front Aspose.Cells | send shape to back Aspose.Cells | read shape Z-order index .NET | Excel shape layering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderDemo
{
    // This example creates a new workbook, adds a rectangle shape, captures its initial ZOrderPosition, moves the shape to the front by assigning the highest Z-order index, then sends it to the back by setting the index to zero, prints each Z-order value, validates the ordering changes, and saves the file as ShapeZOrderDemo.xlsx.
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
                // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

                // Capture the initial Z-order index
                int initialZ = shape.ZOrderPosition;

                // Bring the shape to the front by setting its Z-order to the highest value
                shape.ZOrderPosition = sheet.Shapes.Count; // front
                int frontZ = shape.ZOrderPosition;

                // Send the shape to the back by setting its Z-order to the lowest value (0)
                shape.ZOrderPosition = 0; // back
                int backZ = shape.ZOrderPosition;

                // Output the captured Z-order values
                Console.WriteLine($"Initial Z-order: {initialZ}");
                Console.WriteLine($"After BringToFront Z-order: {frontZ}");
                Console.WriteLine($"After SendToBack Z-order: {backZ}");

                // Verify expected behavior
                if (frontZ > initialZ && backZ < frontZ)
                    Console.WriteLine("Z-order changes as expected.");
                else
                    Console.WriteLine("Unexpected Z-order behavior.");

                // Determine output file path and ensure directory exists
                string outputPath = "ShapeZOrderDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
}
