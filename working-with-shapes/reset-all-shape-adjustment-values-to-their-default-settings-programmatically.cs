// Title: Reset Shape Adjustment Values to Default with Aspose.Cells for .NET (C#)
// Description: Shows how to programmatically set all ShapeGuide adjustment values of an auto shape (e.g., Chevron) to their default (0) using Aspose.Cells for .NET and then save the workbook.
// Keywords: Aspose.Cells | C# | reset shape adjustment | ShapeGuide default | auto shape geometry | worksheet shapes | chevron shape | reset shape guides | Aspose.Cells example | default adjustment values
// Common Searches: reset shapeadjustvalues aspose.cells | default shape guide values c# | clear auto shape adjustments aspose | Aspose.Cells reset shape geometry | how to set shape adjustment to zero
// Developer Intent: Programmatically set every ShapeGuide.Value of a shape’s Geometry to its default (0) in a workbook.
// Use Cases: Standardize auto‑shape appearance before exporting or sharing a workbook. | Remove custom geometry tweaks when reusing a template shape across multiple sheets. | Prepare shapes for visual comparison by resetting all guides to their original values.
// AI Prompts: Write C# code that iterates through all shapes in an Aspose.Cells workbook and resets each ShapeGuide to its default value. | Create a reusable method that accepts a Shape object and sets all Geometry.ShapeAdjustValues to 0 using Aspose.Cells. | Explain how to verify that shape adjustment values have been reset after modification with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to programmatically set all ShapeGuide adjustment values of an auto shape (e.g., Chevron) to their default (0) using Aspose.Cells for .NET and then save the workbook.
    public class ResetShapeAdjustValuesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a sample auto shape that has adjustment values (e.g., a chevron)
                Shape shape = worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 5, 0, 0, 200, 100);

                // Access the geometry of the shape
                Geometry geometry = shape.Geometry;

                // Display initial adjustment values
                Console.WriteLine("Initial ShapeAdjustValues:");
                foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                {
                    Console.WriteLine($"Value: {guide.Value}");
                }

                // Reset each adjustment value to its default (0.0)
                foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                {
                    guide.Value = 0.0;
                }

                // Verify that values have been reset
                Console.WriteLine("\nAfter resetting to default:");
                foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                {
                    Console.WriteLine($"Value: {guide.Value}");
                }

                // Save the workbook with the modified shape
                string outputPath = "ResetShapeAdjustValuesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResetShapeAdjustValuesDemo.Run();
        }
    }
}
