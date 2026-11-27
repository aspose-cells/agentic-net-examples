// Title: Apply a tiled texture fill to every shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# loop through worksheet shapes and set tiled texture fill with Aspose.Cells | Aspose.Cells set FillType.Texture and enable tiling for Excel shape objects | Apply built‑in BlueTissuePaper texture to all shapes in a workbook using .NET | How to check shape fill support before applying texture in Aspose.Cells | Save workbook after modifying shape fill properties in Aspose.Cells C#
// Tags: tiled texture fill implementation Aspose.Cells | FillType.Texture usage in C# | shape enumeration in Excel worksheets | texture fill capability validation Aspose.Cells | persist workbook after shape fill changes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ApplyTiledTextureToShapes
{
    // The example loads an existing workbook (or creates a new one), adds a sample rectangle if needed, iterates through each worksheet and each shape, sets the shape's Fill.FillType to Texture, enables tiling, assigns a built‑in BlueTissuePaper texture, and saves the workbook to the specified output file.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths
            string inputPath = args.Length > 0 ? args[0] : string.Empty;
            string outputPath = args.Length > 1 ? args[1] : "Workbook_With_TiledTexture.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if a valid path is provided; otherwise create a new one
                if (!string.IsNullOrWhiteSpace(inputPath) && File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "DemoSheet";

                    // Add a sample rectangle shape for the demo
                    ws.Shapes.AddRectangle(1, 0, 0, 100, 100, 80);
                    Console.WriteLine("Created new workbook with a sample rectangle shape.");
                }

                // Iterate through all worksheets and apply tiled texture to each shape
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing worksheet \"{worksheet.Name}\" with {worksheet.Shapes.Count} shape(s).");

                    for (int i = 0; i < worksheet.Shapes.Count; i++)
                    {
                        Shape shape = worksheet.Shapes[i];

                        // Ensure the shape supports fill (most shapes do)
                        if (shape.Fill == null)
                        {
                            Console.WriteLine($"Shape {i} ({shape.Name}) does not support fill. Skipping.");
                            continue;
                        }

                        // Set fill type to texture and configure tiled texture
                        shape.Fill.FillType = FillType.Texture;
                        TextureFill textureFill = shape.Fill.TextureFill;
                        textureFill.IsTiling = true;
                        textureFill.Type = TextureType.BlueTissuePaper; // any built‑in texture

                        Console.WriteLine($"Applied tiled texture to shape {i} ({shape.Name}).");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while applying tiled texture to shapes:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
