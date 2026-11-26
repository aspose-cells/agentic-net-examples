// Title: Apply a tiled blue tissue‑paper texture to a named shape across multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to add a tiled texture fill to a shape in each Excel file using Aspose.Cells C# | batch process .xlsx files to set texture fill on a specific shape with Aspose.Cells | C# Aspose.Cells replace missing shape with rectangle and apply texture fill | set TilePicOption scaling for texture fill on shapes in multiple workbooks Aspose.Cells | automate adding texture fill to named shape in a folder of Excel workbooks using Aspose.Cells
// Tags: tiled texture fill on shapes Aspose.Cells | process multiple .xlsx workbooks Aspose.Cells C# | add placeholder rectangle shape Aspose.Cells | TilePicOption configuration Aspose.Cells | texture fill transparency setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchTextureTiling
{
    // The C# program scans a folder for .xlsx files, loads each workbook with Aspose.Cells, locates (or creates) a shape named "TargetShape" on the first worksheet, applies a blue tissue‑paper texture fill with tiling enabled, configures tile scaling and transparency via TilePicOption, and saves the updated workbooks to an output directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Resolve input and output directories
            string inputFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
            string outputFolder = args.Length > 1 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "Processed");

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Process each .xlsx file in the input folder (non‑recursive)
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (workbookFiles.Length == 0)
            {
                Console.WriteLine($"No .xlsx files found in {inputFolder}");
                return;
            }

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    Console.WriteLine($"Processing: {Path.GetFileName(filePath)}");

                    // Ensure the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Work with the first worksheet (adjust as needed)
                    Worksheet worksheet = workbook.Worksheets[0];

                    // ------------------------------------------------------------
                    // Locate the target shape.
                    // ------------------------------------------------------------
                    Shape targetShape;
                    const string targetShapeName = "TargetShape";

                    // Search by name (case‑insensitive)
                    targetShape = null!;
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (string.Equals(shape.Name, targetShapeName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetShape = shape;
                            break;
                        }
                    }

                    // If the shape does not exist, add a rectangle shape.
                    if (targetShape == null)
                    {
                        Console.WriteLine($"Shape \"{targetShapeName}\" not found. Adding a rectangle shape as a placeholder.");
                        // AddRectangle now returns a RectangleShape instance (newer API)
                        RectangleShape rectShape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 300);
                        targetShape = rectShape;
                        targetShape.Name = targetShapeName;
                    }

                    // ------------------------------------------------------------
                    // Apply tiled texture fill to the shape.
                    // ------------------------------------------------------------
                    targetShape.Fill.FillType = FillType.Texture;
                    TextureFill textureFill = targetShape.Fill.TextureFill;

                    // Choose a built‑in texture type.
                    textureFill.Type = TextureType.BlueTissuePaper;

                    // Enable tiling.
                    textureFill.IsTiling = true;

                    // Optional: configure tile options (scale, offset, etc.).
                    TilePicOption tileOptions = new TilePicOption
                    {
                        ScaleX = 0.5,   // 50% width scaling
                        ScaleY = 0.5,   // 50% height scaling
                        OffsetX = 0,    // No horizontal offset
                        OffsetY = 0     // No vertical offset
                    };
                    textureFill.TilePicOption = tileOptions;

                    // Optional: adjust overall transparency (0.0 = opaque, 1.0 = fully transparent)
                    textureFill.Transparency = 0.0;

                    // Save the modified workbook to the output folder.
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Saved processed workbook to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file \"{Path.GetFileName(filePath)}\": {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
