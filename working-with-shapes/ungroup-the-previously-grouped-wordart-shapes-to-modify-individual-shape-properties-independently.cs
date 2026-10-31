// Title: Ungroup WordArt GroupShape and edit fill color, line weight, and rotation of individual shapes with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# ungroup WordArt shapes and change each shape's fill color | set line weight for individual Excel shapes after ungrouping using Aspose.Cells | rotate WordArt shapes programmatically with Aspose.Cells .NET | iterate over worksheet shapes and modify properties with Aspose.Cells in C# | C# code to ungroup grouped shapes in an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells ungroup GroupShape | modify shape fill color Aspose.Cells | set shape line weight Aspose.Cells | rotate Excel shape Aspose.Cells | process WordArt shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through all shapes on the first worksheet, ungroups any GroupShape objects, collects the resulting individual shapes, and then sets each shape's fill color to LightBlue, line weight to 1.5, and rotation angle to 15 degrees before saving the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ungroup any grouped shapes first
            for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
            {
                Shape shape = sheet.Shapes[i];
                if (shape is GroupShape groupShape)
                {
                    try
                    {
                        // Ungroup adds inner shapes to the worksheet and removes the group shape
                        groupShape.Ungroup();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to ungroup shape at index {i}: {ex.Message}");
                    }
                }
            }

            // Collect all individual shapes for processing
            List<Shape> individualShapes = new List<Shape>();
            foreach (Shape shape in sheet.Shapes)
            {
                individualShapes.Add(shape);
            }

            // Modify each shape independently
            foreach (Shape shp in individualShapes)
            {
                // Set fill color if supported
                try
                {
                    if (shp.FillFormat != null)
                    {
                        shp.FillFormat.ForeColor = Color.LightBlue;
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to set fill color: {ex.Message}");
                }

                // Set line weight
                try
                {
                    if (shp.Line != null)
                    {
                        shp.Line.Weight = 1.5f;
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to set line weight: {ex.Message}");
                }

                // Rotate each shape by 15 degrees
                try
                {
                    shp.RotationAngle = 15;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to set rotation angle: {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
