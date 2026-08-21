// Title: Export Excel Worksheet Shapes to Separate SVG Files with Aspose.Cells (.NET)
// Description: Loads an Excel workbook, walks through every worksheet and its Shape collection, and saves each shape as an individual SVG file using Aspose.Cells Shape.ToImage. The example includes file‑existence checks and error handling, preserving the original visual properties for web‑ready reuse.
// Keywords: Aspose.Cells SVG export | C# export Excel shapes | Shape.ToImage SVG | extract worksheet drawings | batch convert Excel shapes to SVG | vector graphics from Excel | .NET Excel shape extraction
// Common Searches: Aspose.Cells export shape to SVG C# | how to save Excel chart as SVG using Aspose | iterate worksheet shapes and export as SVG | convert Excel drawing objects to SVG .NET | batch export Excel shapes to separate SVG files
// Developer Intent: Generate an SVG file for each shape in every worksheet of an Excel workbook.
// Use Cases: Create web‑optimized SVG icons from template workbook shapes for dynamic pages. | Automate conversion of all pictures and diagrams in a multi‑sheet report to SVG for responsive HTML output. | Extract individual diagram elements from Excel to feed a vector‑graphics processing pipeline.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate all worksheets and export each shape as a uniquely named SVG file. | Provide a robust Aspose.Cells example that checks for missing input files, logs export results, and handles shape‑specific errors. | Explain how to configure ImageOrPrintOptions for high‑quality SVG output when calling Shape.ToImage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportShapesToSvg
{
    // Loads an Excel workbook, walks through every worksheet and its Shape collection, and saves each shape as an individual SVG file using Aspose.Cells Shape.ToImage. The example includes file‑existence checks and error handling, preserving the original visual properties for web‑ready reuse.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
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
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    ShapeCollection shapes = sheet.Shapes;

                    // Export each shape as an individual SVG file
                    for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++)
                    {
                        Shape shape = shapes[shapeIndex];

                        // Set image options (default options are sufficient; format is inferred from file extension)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                        try
                        {
                            // Build a file name that identifies the worksheet and shape
                            string fileName = $"Worksheet{wsIndex}_Shape{shapeIndex}.svg";

                            // Export the shape directly to an SVG file
                            shape.ToImage(fileName, imgOptions);

                            Console.WriteLine($"Exported shape {shapeIndex} from worksheet {wsIndex} to {fileName}");
                        }
                        catch (Exception exShape)
                        {
                            Console.WriteLine($"Failed to export shape {shapeIndex} from worksheet {wsIndex}: {exShape.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
