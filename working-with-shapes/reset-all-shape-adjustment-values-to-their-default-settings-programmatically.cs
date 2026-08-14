// Title: Reset Shape Adjustment Guides to Default Values with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, iterates through all worksheets and shapes, accesses each shape's Geometry, sets every ShapeGuide.Value to 0.0 (default), and saves the modified file.
// Keywords: Aspose.Cells | C# | reset shape adjustment | shape guide default | shape geometry | Excel workbook | iterate shapes | clear shape adjustments
// Common Searches: reset shape adjustment guides Aspose.Cells C# | clear shape geometry values in Excel using Aspose.Cells | set all shape guides to default with .NET | programmatically reset custom shape adjustments Aspose.Cells
// Developer Intent: Programmatically set every shape's adjustment guide to its default (0.0) across all worksheets in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Normalize imported template shapes before generating automated reports. | Ensure consistent rendering when exporting workbooks to PDF or image formats. | Prepare a shared workbook for multiple users by removing custom shape tweaks.
// AI Prompts: Write C# code with Aspose.Cells that resets all shape adjustment guides to 0.0 and saves the workbook. | Show how to log each shape's name and type before clearing its adjustment values using Aspose.Cells. | Explain how to modify the sample to apply a custom default value (e.g., 0.5) to every shape guide.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, iterates through all worksheets and shapes, accesses each shape's Geometry, sets every ShapeGuide.Value to 0.0 (default), and saves the modified file.
    public class ResetShapeAdjustValues
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        // Access the geometry of the shape
                        Geometry geometry = shape.Geometry;
                        if (geometry == null) continue;

                        // Reset each adjustment guide to its default (assumed 0.0)
                        foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                        {
                            guide.Value = 0.0;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
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
            ResetShapeAdjustValues.Run();
        }
    }
}
