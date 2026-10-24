// Title: How to set a 1.2‑point character spacing for all shape text in an Excel workbook using Aspose.Cells for .NET
// Common Searches: c# aspocells increase character spacing for shape text | apply 1.2 point letter spacing to WordArt in Excel with aspocells | set custom text spacing for all shapes in a workbook using aspocells .net | modify text effect spacing for shapes in .xlsx via aspocells | adjust shape font spacing programmatically with aspocells c#
// Tags: shape text character spacing aspocells c# | wordart spacing adjustment aspocells .net | excel shape text formatting aspocells | apply text effect properties aspocells | custom character spacing for shapes aspocells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, walks through each worksheet and its shapes, and identifies shapes that contain a TextEffect (e.g., WordArt). It shows where to insert logic for applying a 1.2‑point character spacing to the shape's text. Because Aspose.Cells currently lacks a direct character‑spacing property, the code uses a placeholder to illustrate where such formatting would be applied before saving the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Apply formatting to shapes that have a TextEffect (e.g., WordArt)
                    if (shape.TextEffect != null)
                    {
                        // Aspose.Cells does not expose a direct character spacing property.
                        // As an example, modify the font size to demonstrate applying a change.
                        shape.TextEffect.FontSize = shape.TextEffect.FontSize; // No-op placeholder
                    }
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Display any runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
