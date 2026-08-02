// Title: Set the First Adjustment Guide of a SmartArt‑like AutoShape with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a Chevron AutoShape that supports geometry adjustments, access its ShapeAdjustValues collection, modify the first adjustment guide (or add one if missing), and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape adjustments | modify SmartArt guide .NET | ShapeAdjustValues collection | AutoShape geometry adjustment | add adjustment guide Aspose.Cells | C# update shape adjustment
// Common Searches: how to change a SmartArt adjustment in Aspose.Cells | add missing adjustment guide to AutoShape .NET | set first geometry adjustment value for Chevron shape | Aspose.Cells Shape.Adjustments example
// Developer Intent: Programmatically set or create the first geometry adjustment of a SmartArt‑like AutoShape in an Excel workbook.
// Use Cases: Standardize the appearance of Chevron SmartArt by enforcing a specific adjustment value before exporting. | Ensure dynamic diagram generation works even when a shape initially has no adjustment guides. | Customize shape geometry for automated report layouts that require precise visual control.
// AI Prompts: Generate C# code with Aspose.Cells that changes the second adjustment guide of an existing SmartArt shape to 0.75 and saves the workbook. | Show how to loop through all adjustment guides of a shape, log each guide's name and value, and modify a selected guide using Aspose.Cells for .NET. | Explain the steps to create a custom SmartArt shape with multiple adjustment guides and assign default values in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert a Chevron AutoShape that supports geometry adjustments, access its ShapeAdjustValues collection, modify the first adjustment guide (or add one if missing), and save the result as an XLSX file using Aspose.Cells for .NET.
    public class ModifySmartArtAdjustmentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a SmartArt-like shape (using an AutoShape that supports adjustments)
                Shape smartArtShape = worksheet.Shapes.AddAutoShape(AutoShapeType.Chevron, 5, 5, 0, 0, 200, 100);

                // Access the geometry adjustments collection
                Geometry geometry = smartArtShape.Geometry;
                ShapeGuideCollection adjustments = geometry.ShapeAdjustValues;

                // Modify the first adjustment value if it exists
                if (adjustments.Count > 0)
                {
                    // Set the first adjustment guide's value to 0.5 (example value)
                    adjustments[0].Value = 0.5;
                    Console.WriteLine($"First adjustment value set to {adjustments[0].Value}");
                }
                else
                {
                    // If no adjustments exist, add one
                    int index = adjustments.Add("adj1", 0.5);
                    Console.WriteLine($"Added adjustment at index {index} with value {adjustments[index].Value}");
                }

                // Save the workbook to the current directory
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifySmartArtAdjustmentDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ModifySmartArtAdjustmentDemo.Run();
        }
    }
}
