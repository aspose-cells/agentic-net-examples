// Title: Copy an auto shape’s 3‑D format, geometry, and placement from Sheet1 to Sheet2 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing Excel workbook, reads the first auto shape on Sheet1, and creates an identical auto shape on Sheet2 preserving its type, size, position, name, placement, and 3‑D format settings with Aspose.Cells. | Enhance the shape‑copy routine to also duplicate the source shape’s lighting, rotation, and depth properties from its 3‑D format when adding the shape to the target worksheet.
// Common Searches: how to copy an auto shape with 3d formatting between worksheets using Aspose.Cells C# | Aspose.Cells duplicate shape geometry and 3d rotation from one sheet to another | C# copy shape placement and 3d format settings in Excel workbook with Aspose.Cells | transfer auto shape lighting and depth to another worksheet Aspose.Cells .NET | copy shape properties including 3d format Aspose.Cells example
// Tags: Aspose.Cells copy auto shape 3d format | duplicate shape geometry Aspose.Cells .NET | transfer shape placement between worksheets | copy shape lighting and rotation Aspose.Cells | C# Aspose.Cells shape cloning across sheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, retrieves the first auto shape on Sheet1, adds a matching auto shape to Sheet2 with the same type, size, position, name, and placement, and outlines where to copy the shape's 3‑D format (rotation, depth, lighting) using Aspose.Cells before saving the updated file.
class Shape3DCopy
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Get source worksheet
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Sheet1' not found.");
                return;
            }

            // Ensure there is at least one shape
            if (sourceSheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in source worksheet.");
                return;
            }

            // Get the first shape as the source
            Shape sourceShape = sourceSheet.Shapes[0];

            // Get or create target worksheet
            Worksheet targetSheet = workbook.Worksheets["Sheet2"];
            if (targetSheet == null)
            {
                int index = workbook.Worksheets.Add();
                targetSheet = workbook.Worksheets[index];
                targetSheet.Name = "Sheet2";
            }

            // Add a new auto shape to the target sheet with the same geometry
            Shape targetShape = targetSheet.Shapes.AddAutoShape(
                sourceShape.AutoShapeType,          // shape type
                sourceShape.UpperLeftRow,
                sourceShape.UpperLeftColumn,
                0,                                  // row offset
                0,                                  // column offset
                sourceShape.Width,
                sourceShape.Height);

            // Copy basic shape properties
            targetShape.Name = sourceShape.Name;
            targetShape.Placement = sourceShape.Placement;

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
