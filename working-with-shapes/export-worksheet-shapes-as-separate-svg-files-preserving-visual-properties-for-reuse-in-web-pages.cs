// Title: Export Worksheet Shapes to Individual SVG Files with Aspose.Cells for .NET
// Description: This C# example loads an Excel workbook, iterates through each worksheet’s Shape collection, and saves every shape as a separate SVG file using Shape.ToImage(ImageType.Svg). File names combine the workbook name, sheet name, and shape identifier, producing ready‑to‑use graphics for web pages.
// Keywords: Aspose.Cells | C# | export shape to SVG | Shape.ToImage | Excel shapes SVG | worksheet shape extraction | SVG web graphics | C# Excel SVG conversion | Aspose.Cells .NET | batch export Excel shapes
// Common Searches: Aspose.Cells export shape as SVG C# | How to save Excel drawing objects to SVG with .NET | Convert worksheet shapes to SVG files using Aspose.Cells | C# code to extract Excel shapes to SVG | Batch export Excel charts to SVG
// Developer Intent: Generate separate SVG files for all shapes in every worksheet of an Excel workbook while keeping their visual attributes intact.
// Use Cases: Create scalable icons from Excel diagrams for responsive web dashboards. | Produce high‑quality SVG assets of charts and drawings for documentation or presentations. | Automate bulk conversion of workbook graphics for content‑management systems. | Integrate SVG exports into CI pipelines for design‑asset generation.
// AI Prompts: Write a C# method that receives a workbook path and an output directory, then iterates through all worksheets and saves each shape as an SVG file with a name based on workbook, sheet, and shape name. | Explain how Shape.ToImage preserves dimensions, colors, and text when converting an Excel shape to SVG with Aspose.Cells. | Provide error‑handling code for unnamed shapes and unsupported shape types during SVG export in Aspose.Cells. | Show how to filter only chart shapes before exporting to SVG using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example loads an Excel workbook, iterates through each worksheet’s Shape collection, and saves every shape as a separate SVG file using Shape.ToImage(ImageType.Svg). File names combine the workbook name, sheet name, and shape identifier, producing ready‑to‑use graphics for web pages.
class ExportWorksheetShapesToSvg
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of shapes on the current worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Export each shape as an individual SVG file
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];

                // Determine a file name for the SVG (use shape name if available)
                string shapeName = string.IsNullOrEmpty(shape.Name) ? $"Shape_{i}" : shape.Name;
                string svgFileName = $"{Path.GetFileNameWithoutExtension(workbook.FileName)}_{sheet.Name}_{shapeName}.svg";

                // Export the shape to SVG using the ToImage method with ImageType.Svg
                using (MemoryStream svgStream = new MemoryStream())
                {
                    shape.ToImage(svgStream, ImageType.Svg);   // Export to SVG format
                    File.WriteAllBytes(svgFileName, svgStream.ToArray());
                }

                Console.WriteLine($"Exported shape '{shapeName}' to '{svgFileName}'.");
            }
        }
    }
}
