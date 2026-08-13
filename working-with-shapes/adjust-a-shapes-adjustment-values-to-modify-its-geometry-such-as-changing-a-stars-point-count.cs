// Title: Aspose.Cells .NET – Change a Star AutoShape’s Point Count by Adjusting Its Geometry
// Description: This C# example creates a new workbook, inserts a Star10 auto‑shape, accesses the shape’s Geometry object, and sets the first ShapeAdjustValues entry (or adds one) to 0.5 – the value that controls the star’s number of points. The workbook is then saved as AdjustedStarShape.xlsx.
// Keywords: Aspose.Cells shape adjustment | star auto shape point count | modify geometry values .NET | ShapeAdjustValues C# | adjust star shape Aspose.Cells
// Common Searches: how to change star shape points Aspose.Cells | adjust auto shape geometry .NET | add shape adjustment guide Excel C# | modify star AutoShape point count programmatically | Aspose.Cells change shape geometry values
// Developer Intent: Set or add an adjustment value for a Star AutoShape to alter its geometry—specifically the number of points—and save the updated workbook.
// Use Cases: Generate a custom star diagram with a precise point count for a financial dashboard. | Allow end‑users to select a star‑shape style, then programmatically apply the chosen point count before exporting to Excel. | Add missing adjustment guides to any auto‑shape to enable further geometric tweaks in automated report generation.
// AI Prompts: Show C# code that reads and updates multiple ShapeAdjustValues of a star shape using Aspose.Cells. | Provide an example that changes a star shape’s point count based on a variable and saves the workbook. | Explain how to add custom adjustment guides to other AutoShape types in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AdjustStarShapeDemoApp
{
    // This C# example creates a new workbook, inserts a Star10 auto‑shape, accesses the shape’s Geometry object, and sets the first ShapeAdjustValues entry (or adds one) to 0.5 – the value that controls the star’s number of points. The workbook is then saved as AdjustedStarShape.xlsx.
    class AdjustStarShapeDemo
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a star auto shape (Star10) to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left pixel offset X, Y, width, height
            Shape starShape = worksheet.Shapes.AddAutoShape(AutoShapeType.Star10, 2, 2, 0, 0, 200, 200);

            // Access the geometry of the shape
            Geometry geometry = starShape.Geometry;

            // Modify the first adjustment value if it exists; otherwise add a new guide
            if (geometry.ShapeAdjustValues.Count > 0)
            {
                // For star shapes the first adjust value controls the number of points (0‑1 range)
                geometry.ShapeAdjustValues[0].Value = 0.5; // Example value to change point count
            }
            else
            {
                // Add a new adjustment guide named "adj" with the desired value
                geometry.ShapeAdjustValues.Add("adj", 0.5);
            }

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdjustedStarShape.xlsx");

            // Save the workbook with the modified shape
            workbook.Save(outputPath);
        }
    }
}
