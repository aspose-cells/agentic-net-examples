// Title: How to apply an 8‑point blue glow effect to every shape in an Aspose.Cells worksheet using C#
// Common Searches: C# Aspose.Cells set blue glow with 8 point radius on all worksheet shapes | apply glow effect to multiple shapes in Aspose.Cells workbook programmatically | iterate shapes and configure Glow property in Aspose.Cells C# example | Aspose.Cells shape formatting add blue outer glow to rectangles and ovals | set shape glow size and color in Aspose.Cells using .NET
// Tags: Aspose.Cells shape glow API C# | apply outer glow to Excel shapes | set glow radius and color programmatically | enumerate worksheet shapes Aspose.Cells | blue glow effect example Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a rectangle, an oval, and a line shape, then loops through all shapes on the first worksheet and sets each shape's Glow property to a blue color with an 8‑point radius and full opacity, finally saving the file as GlowedShapes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape
                Shape rect = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 80);
                rect.Name = "SampleRectangle";

                // Add an oval (ellipse) shape using the correct enum value
                Shape ellipse = sheet.Shapes.AddShape(MsoDrawingType.Oval, 2, 0, 2, 0, 150, 80);
                ellipse.Name = "SampleEllipse";

                // Add a line shape
                Shape line = sheet.Shapes.AddShape(MsoDrawingType.Line, 3, 0, 3, 0, 200, 0);
                line.Name = "SampleLine";

                // Apply a blue glow effect (size = 8 points) to all shapes
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        GlowEffect glow = shape.Glow; // read‑only property, members are settable

                        CellsColor glowColor = workbook.CreateCellsColor();
                        glowColor.Color = Color.Blue; // System.Drawing.Color

                        glow.Color = glowColor;
                        glow.Size = 8;               // 8‑point radius
                        glow.Transparency = 0.0;     // fully opaque

                        Console.WriteLine($"Applied blue glow (size 8pt) to shape \"{shape.Name}\" of type {shape.Type}.");
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Failed to apply glow to shape \"{shape.Name}\": {exShape.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputPath = "GlowedShapes.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
