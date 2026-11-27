// Title: Set a star auto shape's inner radius ratio by adjusting its guide value with Aspose.Cells for .NET (C#)
// Common Searches: how to change the inner radius of a star auto shape using Aspose.Cells in C# | Aspose.Cells programmatically modify shape adjustment guide values | add an adjustment guide to an auto shape when none exist with Aspose.Cells .NET | C# example for updating star shape geometry in an Excel workbook via Aspose.Cells | adjust star shape points count by setting adjustment guide in Aspose.Cells
// Tags: Aspose.Cells set shape adjustment guide C# | modify star auto shape geometry Aspose.Cells | add shape guide to auto shape Aspose.Cells | adjust inner radius ratio of star shape .NET | Aspose.Cells shape geometry manipulation | C# update shape adjustment values Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeAdjustDemo
{
    // The example creates a new workbook, adds a 5‑point star auto shape, accesses its geometry adjustment guides, updates (or adds) the first guide to change the inner‑radius ratio, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to output workbook
            string outputPath = "StarShapeAdjustDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a 5‑point star auto shape
                // Parameters: AutoShapeType, upper left row, upper left column, offset X, offset Y, width, height
                Shape starShape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.Star5,   // 5‑point star
                    2,                     // row index (zero‑based)
                    2,                     // column index
                    0,                     // offset X in pixels
                    0,                     // offset Y in pixels
                    200,                   // width in pixels
                    200);                  // height in pixels

                // Access the geometry of the shape
                Geometry geometry = starShape.Geometry;

                // Display current adjustment values (if any)
                Console.WriteLine("Initial adjustment values:");
                if (geometry.ShapeAdjustValues.Count == 0)
                {
                    Console.WriteLine("  No adjustment guides present.");
                }
                else
                {
                    for (int i = 0; i < geometry.ShapeAdjustValues.Count; i++)
                    {
                        ShapeGuide guide = geometry.ShapeAdjustValues[i];
                        // ShapeGuide does not expose a Name property; only the value is available
                        Console.WriteLine($"  Guide {i}: Value = {guide.Value}");
                    }
                }

                // The star shape typically has one adjustment guide that controls the inner radius ratio.
                // Modify the first guide (if it exists) to change the star's appearance.
                if (geometry.ShapeAdjustValues.Count > 0)
                {
                    // Set a new ratio (value must be between 0.0 and 1.0)
                    double newRatio = 0.2; // smaller inner radius makes points longer
                    geometry.ShapeAdjustValues[0].Value = newRatio;
                    Console.WriteLine($"\nModified first adjustment guide to value {newRatio}");
                }
                else
                {
                    // If the collection is empty, add a new guide manually.
                    // The name "Adj1" is the conventional name for the first adjustment.
                    double newRatio = 0.2;
                    geometry.ShapeAdjustValues.Add("Adj1", newRatio);
                    Console.WriteLine($"\nAdded new adjustment guide 'Adj1' with value {newRatio}");
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
