// Title: Adjust Star Shape Geometry Using ShapeAdjustValues in Aspose.Cells for .NET
// Description: Demonstrates how to add a 10‑point star auto shape, access its Geometry object, modify or create an adjustment guide (e.g., inner radius), and save the workbook with the updated shape using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape adjustment | ShapeAdjustValues .NET | modify star auto shape | geometry adjustment guide | change inner radius Aspose.Cells | auto shape geometry C# | Aspose.Cells shape customization
// Common Searches: Aspose.Cells change star shape size | how to set adjustment guide for auto shape .NET | modify star inner radius programmatically | add ShapeAdjustValues to shape Aspose.Cells | adjust geometry of auto shapes in Excel
// Developer Intent: Update or add an adjustment guide to an auto shape to alter its geometry, such as changing a star's inner radius, before saving the workbook.
// Use Cases: Reduce the inner radius of a star to elongate its points for a custom chart. | Add a missing adjustment guide to a newly inserted shape so future geometry tweaks are possible. | Apply a uniform adjustment factor to multiple shapes in a worksheet based on user input.
// AI Prompts: Write C# code that iterates over all auto shapes in a worksheet and sets each shape's first ShapeAdjustValues entry to a value supplied by the user. | Show how to create a polygon auto shape and control the number of sides using adjustment guides with Aspose.Cells for .NET. | Explain how to retrieve, list, and modify all adjustment guide names and values for a given shape in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAdjustmentDemo
{
    // Demonstrates how to add a 10‑point star auto shape, access its Geometry object, modify or create an adjustment guide (e.g., inner radius), and save the workbook with the updated shape using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a star auto shape (10‑point star) to the worksheet
            // Parameters: AutoShapeType, upper left row, upper left column, upper left offset X, offset Y, width, height
            Shape starShape = worksheet.Shapes.AddAutoShape(AutoShapeType.Star10, 2, 2, 0, 0, 150, 150);

            // Access the geometry of the shape
            Geometry geometry = starShape.Geometry;

            // Ensure the shape has at least one adjustment guide
            if (geometry.ShapeAdjustValues.Count > 0)
            {
                // Modify the first adjustment value (e.g., inner radius of the star)
                // Setting it to 0.5 makes the star points longer; adjust as needed
                geometry.ShapeAdjustValues[0].Value = 0.5;
                Console.WriteLine("Adjusted first guide value to 0.5");
            }
            else
            {
                // If no guides exist, add a new one named "adj1"
                int index = geometry.ShapeAdjustValues.Add("adj1", 0.5);
                Console.WriteLine($"Added adjustment guide at index {index} with value 0.5");
            }

            // Save the workbook with the modified shape
            workbook.Save("StarShapeAdjusted.xlsx");
        }
    }
}
