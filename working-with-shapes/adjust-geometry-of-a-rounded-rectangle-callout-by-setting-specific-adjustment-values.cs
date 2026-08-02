// Title: Adjust Rounded Rectangle Callout Geometry with Shape Guides in Aspose.Cells for .NET
// Description: Creates a workbook, inserts a RoundedRectangle auto‑shape, accesses its Geometry.ShapeAdjustValues collection, adds an "adj1" guide to set corner roundness (e.g., 0.25), optionally updates the guide, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells shape adjustment | rounded rectangle callout geometry | ShapeGuideCollection .NET | adjust corner roundness Aspose.Cells | auto shape guide values C# | Excel shape geometry Aspose | custom shape guides Aspose.Cells
// Common Searches: how to set shape adjustment values in Aspose.Cells | change corner roundness of rounded rectangle in C# | modify auto shape geometry guides Aspose.Cells | Aspose.Cells add adj1 guide to shape | adjust rounded rectangle callout Aspose.Cells .NET
// Developer Intent: Customize the corner roundness of a RoundedRectangle callout by setting specific shape‑guide values.
// Use Cases: Create a rounded rectangle callout with 25% corner roundness and later increase it to 30% before saving. | Iterate through existing shape guides to apply business‑rule‑based modifications. | Apply different adjustment guides to multiple auto shapes for varied visual styling in a worksheet.
// AI Prompts: Show code to add multiple adjustment guides (adj1, adj2) to a rounded rectangle using Aspose.Cells for .NET. | Provide a snippet that reads all current ShapeGuide values of a shape and logs them to the console. | Explain how to reset shape geometry adjustments to their default settings in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGeometryDemo
{
    // Creates a workbook, inserts a RoundedRectangle auto‑shape, accesses its Geometry.ShapeAdjustValues collection, adds an "adj1" guide to set corner roundness (e.g., 0.25), optionally updates the guide, and saves the file as an Excel workbook.
    public class RoundedRectangleCalloutAdjustment
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rounded rectangle shape (AutoShapeType.RoundedRectangle)
                // Parameters: upper left column, upper left row, upper left offset X, upper left offset Y, width, height
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.RoundedRectangle, // shape type
                    2,    // upper left column
                    2,    // upper left row
                    0,    // offset X
                    0,    // offset Y
                    200,  // width
                    100   // height
                );

                // Access the geometry adjustment collection
                ShapeGuideCollection adjusts = shape.Geometry.ShapeAdjustValues;

                // Set specific adjustment values.
                // For a rounded rectangle the typical guide is:
                // "adj1" – roundness of the corners (0.0 – 1.0)
                adjusts.Add("adj1", 0.25); // 25% corner roundness

                // Optionally modify an existing guide if needed
                if (adjusts.Count > 0)
                {
                    // Change the first guide value to a new roundness
                    adjusts[0].Value = 0.30;
                }

                // Save the workbook with the modified shape
                string outputPath = "RoundedRectangleCalloutAdjusted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            RoundedRectangleCalloutAdjustment.Run();
        }
    }
}
