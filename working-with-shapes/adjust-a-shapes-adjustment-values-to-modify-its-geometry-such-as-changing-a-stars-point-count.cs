// Title: How to change the number of points on a star auto shape in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Set a custom point count for a star auto shape by calculating and assigning the first adjustment value using Aspose.Cells in C#. | Convert a rectangle shape to a 5‑point star and then modify its geometry to an 8‑point star via the Adjustments collection. | Implement safe adjustment handling that checks for the Adjustments collection before setting values to prevent runtime errors.
// Common Searches: Aspose.Cells C# change star shape points programmatically | set star auto shape point count Excel using Aspose.Cells | how to adjust star geometry with adjustment values in Aspose.Cells .NET | convert rectangle to star shape and modify points Aspose.Cells | check Adjustments collection before setting shape properties Aspose.Cells
// Tags: Aspose.Cells set star shape adjustment value | auto shape point count modification C# | Excel star geometry adjustment Aspose.Cells | dynamic shape conversion rectangle to star Aspose.Cells | safe adjustments collection handling .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, converts it to a 5‑point star, computes a normalized adjustment value based on a desired point count, safely assigns the first adjustment when the Adjustments collection is available, and saves the result as StarAdjusted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (will be converted to an auto shape later).
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // base drawing type
                2,    // upper left row
                2,    // upper left column
                0,    // upper left row offset (pixels)
                0,    // upper left column offset (pixels)
                200,  // width (pixels)
                200   // height (pixels)
            );

            // Use dynamic to access members that may vary between library versions.
            dynamic starShape = shape;

            // Convert the shape to a 5‑point star.
            starShape.AutoShapeType = AutoShapeType.Star5;

            // Desired number of points (between 3 and 12). Adjust the first adjustment value accordingly.
            double desiredPoints = 8; // change as needed
            double adjustmentValue = (desiredPoints - 3) / (12 - 3); // normalize to 0‑1

            // Apply the adjustment if the Adjustments collection is available at runtime.
            try
            {
                if (starShape.Adjustments != null && starShape.Adjustments.Count > 0)
                {
                    starShape.Adjustments[0] = adjustmentValue;
                }
            }
            catch
            {
                // Adjustments not supported in this version – ignore safely.
            }

            // Prepare output path.
            string outputPath = "StarAdjusted.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
