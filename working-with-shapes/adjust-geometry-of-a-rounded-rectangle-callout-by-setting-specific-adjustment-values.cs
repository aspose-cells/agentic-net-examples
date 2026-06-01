using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class RoundedRectangleCalloutGeometryDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rounded rectangle shape (callout not directly supported)
            Shape shape = worksheet.Shapes.AddAutoShape(
                AutoShapeType.RoundedRectangle, // shape type
                2,    // upper left column
                0,    // upper left row
                2,    // upper left offset X (pixels)
                0,    // upper left offset Y (pixels)
                200,  // width (pixels)
                150   // height (pixels)
            );

            // Access the geometry of the shape
            Geometry geometry = shape.Geometry;

            // Set specific adjustment values (example keys)
            geometry.ShapeAdjustValues.Add("adj1", 0.2);
            geometry.ShapeAdjustValues.Add("adj2", 0.3);
            geometry.ShapeAdjustValues.Add("adj3", 0.4);
            geometry.ShapeAdjustValues.Add("adj4", 0.5);

            // Optionally modify the first adjustment value
            if (geometry.ShapeAdjustValues.Count > 0)
            {
                geometry.ShapeAdjustValues[0].Value = 0.25;
            }

            // Save the workbook with the modified shape
            string outputPath = "RoundedRectangleCalloutGeometryDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}