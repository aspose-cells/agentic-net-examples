// Title: Increase a shape's Z-order by five positions in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition offset | how to move an Excel shape forward by 5 layers with Aspose.Cells | programmatically increase Z-order of a shape in .NET Aspose.Cells | adjust shape layering order in Excel using Aspose.Cells API | C# Aspose.Cells change shape Z-order relative to current value
// Tags: shape ZOrderPosition manipulation Aspose.Cells | increase shape Z-order offset .NET | Excel shape layer ordering Aspose.Cells | C# adjust shape stacking order | Aspose.Cells example moving shape forward

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The code loads or creates an Excel workbook, ensures a rectangle shape exists on the first worksheet, retrieves the first shape, reads its current ZOrderPosition, adds five to this value, updates the property, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created a new workbook at '{inputPath}'.");
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape; add a rectangle if none exist.
            if (sheet.Shapes.Count == 0)
            {
                Shape demoShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                demoShape.Name = "DemoShape";
            }

            // Retrieve the first shape on the sheet.
            Shape targetShape = sheet.Shapes[0];

            try
            {
                // Get current Z-order and increase it using ZOrderPosition property.
                int currentZOrder = targetShape.ZOrderPosition;
                targetShape.ZOrderPosition = currentZOrder + 5;

                Console.WriteLine($"Shape '{targetShape.Name}' Z-order changed from {currentZOrder} to {targetShape.ZOrderPosition}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adjusting Z-order: {ex.Message}");
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
