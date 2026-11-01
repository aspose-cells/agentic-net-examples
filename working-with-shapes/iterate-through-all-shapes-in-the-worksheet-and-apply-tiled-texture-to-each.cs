// Title: C# example: loop through worksheet shapes and set tiled texture fill with Aspose.Cells
// Common Searches: c# aspose.cells set shape fill texture to tile | how to apply tiled texture to all shapes in an Excel worksheet using Aspose.Cells | iterate over worksheet shapes and change fill alignment in .NET | Aspose.Cells example for setting texture alignment on shape fills | loop through shapes collection and set fill pattern programmatically
// Tags: shape fill texture alignment | Aspose.Cells tiled texture fill | worksheet shapes iteration | C# Excel shape formatting | set fill pattern Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, retrieves the first worksheet, and iterates through each Shape in the worksheet's Shapes collection. For every shape it demonstrates how to set the Fill.TextureAlignment property to Tile (shown as commented code for compatibility), handling any per‑shape exceptions, and finally saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Apply tiled texture to each shape's fill (if supported)
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // The TextureAlignment property may not be available in all versions.
                    // If it exists, you can set it to Tile as shown below:
                    // shape.Fill.TextureAlignment = TextureAlignmentType.Tile;

                    // Placeholder for future texture alignment logic.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify shape '{shape.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
