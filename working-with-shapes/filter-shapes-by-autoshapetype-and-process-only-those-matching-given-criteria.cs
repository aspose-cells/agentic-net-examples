// Title: Filter AutoShape shapes by type and modify them with Aspose.Cells for .NET (C#)
// Description: Loads or creates an Excel workbook, adds a sample shape when needed, then scans every worksheet to find shapes whose AutoShapeType matches a given value (e.g., Rectangle). Matching shapes receive a new fill color and a renamed identifier before the file is saved.
// Keywords: Aspose.Cells | C# shape filtering | AutoShapeType | Excel shape fill color | rename Excel shapes | Aspose.Cells Shapes API | filter rectangle shapes | modify shape properties | Aspose.Cells .NET example | Excel automation
// Common Searches: Aspose.Cells filter shapes by AutoShapeType C# | Change fill color of rectangle shapes in Excel using Aspose.Cells | Rename Excel auto shapes programmatically with Aspose.Cells | Iterate through worksheet shapes Aspose.Cells .NET | How to process only specific shape types in an Excel file
// Developer Intent: Find shapes of a specified AutoShapeType and apply custom formatting or naming.
// Use Cases: Apply a uniform fill color to all rectangle auto shapes across multiple worksheets. | Prefix the names of all oval shapes with "Processed_" before saving. | Automatically add a placeholder shape when a workbook is missing, then process only shapes matching a target type. | Generate reports that highlight specific shape types by changing their appearance. | Batch‑update shape properties in a template workbook for branding purposes.
// AI Prompts: Write C# code using Aspose.Cells to iterate over all worksheets, filter shapes where AutoShapeType equals Rectangle, set their fill to LightCoral, and prepend "Processed_" to their names. | Show how to add a default shape if an Excel file does not exist, then process only shapes of type Triangle with Aspose.Cells. | Provide a reusable method that accepts an AutoShapeType parameter and updates matching shapes' color and name in a workbook. | Explain how to retrieve and modify shape properties (fill, line, name) for specific AutoShapeTypes using Aspose.Cells for .NET. | Create a console application example that demonstrates shape filtering, color change, and renaming with Aspose.Cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeFilterDemo
{
    // Loads or creates an Excel workbook, adds a sample shape when needed, then scans every worksheet to find shapes whose AutoShapeType matches a given value (e.g., Rectangle). Matching shapes receive a new fill color and a renamed identifier before the file is saved.
    public class ShapeFilterProcessor
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with a sample shape
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                // Add a sample rectangle shape for demonstration purposes
                Shape sampleShape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 1, 0, 1, 0, 100, 50);
                sampleShape.Name = "SampleRectangle";
            }

            // Define the AutoShapeType to filter (e.g., Rectangle)
            AutoShapeType targetType = AutoShapeType.Rectangle;

            // Iterate through worksheets and filter shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.AutoShapeType == targetType)
                    {
                        // Change fill color and rename the shape
                        shape.Fill.SolidFill.Color = Color.LightCoral;
                        shape.Name = "Processed_" + shape.Name;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
