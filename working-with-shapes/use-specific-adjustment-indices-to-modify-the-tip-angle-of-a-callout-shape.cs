// Title: Aspose.Cells .NET – Change Callout Tip Angle Using ShapeAdjustValues Index
// Description: Creates a workbook, inserts a RightArrowCallout auto‑shape, accesses its Geometry, sets the first adjustment guide (index 0) to 0.4 to control the tip angle, and saves the file.
// Keywords: Aspose.Cells callout adjustment | ShapeAdjustValues tip angle | C# modify callout geometry | adjust callout shape Aspose.Cells | Excel shape adjustment index
// Common Searches: how to set callout tip angle Aspose.Cells | ShapeAdjustValues index 0 callout .NET | change callout arrow angle programmatically | adjust callout geometry Aspose.Cells C#
// Developer Intent: Programmatically set the tip angle of a callout shape by updating its first ShapeAdjustValues entry.
// Use Cases: Generate a worksheet with a right‑arrow callout whose tip is narrowed to 40 % of the maximum angle. | Loop through multiple callout shapes and standardize tip angles for consistent visual styling. | Create dynamic reports where callout arrows highlight key cells with customized tip angles.
// AI Prompts: Show code to adjust the tip angle of CloudCallout and RoundedRectangleCallout shapes using Aspose.Cells geometry indexes. | Write a reusable method that receives a Shape and a proportion (0‑1) and safely sets the callout tip angle, handling shapes without adjustment guides. | Explain how to query the valid range of ShapeAdjustValues for a given callout type and apply it to set the tip angle.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts a RightArrowCallout auto‑shape, accesses its Geometry, sets the first adjustment guide (index 0) to 0.4 to control the tip angle, and saves the file.
    public class CalloutTipAngleAdjustment
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a right-arrow callout shape (you can also use other callout types)
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape callout = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout, // callout shape
                    2,    // upper left row
                    2,    // upper left column
                    0,    // top offset (pixels)
                    0,    // left offset (pixels)
                    150,  // height (pixels)
                    250); // width (pixels)

                // Access the geometry of the shape to work with adjustment guides
                Geometry geometry = callout.Geometry;

                // Ensure that the shape has adjustment guides (most callouts have at least one)
                if (geometry.ShapeAdjustValues.Count > 0)
                {
                    // The first adjustment guide (index 0) controls the tip angle of many callout shapes.
                    // Set the value to a desired proportion (0.0 – 1.0). Here we set it to 0.4 (40% of the max angle).
                    geometry.ShapeAdjustValues[0].Value = 0.4;

                    // Optionally, you can modify additional guides if needed, e.g., second guide at index 1.
                    // geometry.ShapeAdjustValues[1].Value = 0.2;
                }

                // Save the workbook with the modified shape
                workbook.Save("CalloutTipAngleAdjustment.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CalloutTipAngleAdjustment.Run();
        }
    }
}
