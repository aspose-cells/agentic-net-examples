// Title: Export each worksheet shape to a separate SVG file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export individual shape from Excel to SVG | How to save worksheet shapes as separate SVG images with Aspose.Cells | Aspose.Cells ImageOrPrintOptions SVG CssPrefix example in C# | Batch export Excel drawing objects to SVG files using Aspose.Cells | Extract Excel shapes to SVG files with .NET API
// Tags: shape.ToImage SVG export with Aspose.Cells | ImageOrPrintOptions SvgCssPrefix configuration | export worksheet shapes to individual SVG files | batch extract Excel shapes as SVG using .NET | sanitize worksheet name for SVG filenames

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportWorksheetShapesToSvg
{
    // The sample loads an Excel workbook, creates an output folder, iterates through each worksheet's ShapeCollection, and exports every shape as an individual SVG file. It builds a safe filename using the worksheet name, configures ImageOrPrintOptions for SVG output with a CSS prefix, and calls shape.ToImage to generate the SVG files while handling missing files and export errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path (first argument) or default file name.
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            // Prepare output folder.
            string outputFolder = Path.Combine(Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(), "output_shapes");
            Directory.CreateDirectory(outputFolder);
            Console.WriteLine($"SVG files will be saved to: {outputFolder}");

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet.
                for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[wsIndex];
                    ShapeCollection shapes = sheet.Shapes;

                    if (shapes == null || shapes.Count == 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains no shapes.");
                        continue;
                    }

                    Console.WriteLine($"Exporting {shapes.Count} shape(s) from worksheet \"{sheet.Name}\"...");

                    // Export each shape as an individual SVG file.
                    for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++)
                    {
                        Shape shape = shapes[shapeIndex];

                        // Build a unique file name for the shape.
                        string safeSheetName = MakeFileSystemSafe(sheet.Name);
                        string fileName = $"{safeSheetName}_Shape_{shapeIndex + 1}.svg";
                        string fullPath = Path.Combine(outputFolder, fileName);

                        // Configure image options for SVG output.
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            ImageType = ImageType.Svg,
                            OnePagePerSheet = true,
                            // Use the (now deprecated) property to set a CSS prefix for the SVG.
                            SvgCssPrefix = $"shape{shapeIndex + 1}_"
                        };

                        try
                        {
                            // Export the shape.
                            shape.ToImage(fullPath, options);
                            Console.WriteLine($"  Saved shape #{shapeIndex + 1} to \"{fileName}\"");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Failed to export shape #{shapeIndex + 1}: {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Replaces invalid filename characters with underscores.
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
