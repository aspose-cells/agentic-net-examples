// Title: How to batch add a diagonal semi‑transparent WordArt watermark to every worksheet in multiple .xlsx files using Aspose.Cells for .NET
// Common Searches: C# code to add a diagonal WordArt watermark to all worksheets in a folder of Excel files with Aspose.Cells | batch process .xlsx files to insert semi‑transparent watermark using Aspose.Cells .NET | automate adding the same watermark text to multiple Excel workbooks in C# | apply WordArt shape as watermark to each sheet in a directory with Aspose.Cells | Aspose.Cells example for bulk watermarking of Excel workbooks
// Tags: batch watermarking with Aspose.Cells | WordArt shape watermark in Excel using C# | diagonal semi‑transparent watermark for worksheets | automated processing of multiple .xlsx workbooks | Aspose.Cells bulk worksheet watermark example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchWatermark
{
    // The program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, inserts a WordArt shape containing the supplied text as an 80% transparent diagonal watermark on every worksheet, and saves the updated files to an output directory.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 3)
                {
                    Console.WriteLine("Usage: BatchWatermark <inputDir> <outputDir> <watermarkText>");
                    return;
                }

                string inputDir = args[0];
                string outputDir = args[1];
                string watermarkText = args[2];

                if (!Directory.Exists(inputDir))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDir}");
                    return;
                }

                if (!Directory.Exists(outputDir))
                {
                    try
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to create output directory: {outputDir}");
                        Console.WriteLine(ex.Message);
                        return;
                    }
                }

                string[] excelFiles = Directory.GetFiles(inputDir, "*.xlsx", SearchOption.TopDirectoryOnly);
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine($"No .xlsx files found in {inputDir}");
                    return;
                }

                foreach (string filePath in excelFiles)
                {
                    try
                    {
                        // Verify source file exists
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load workbook
                        Workbook workbook = new Workbook(filePath);

                        // Apply watermark to each worksheet
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Add WordArt as watermark
                            Shape watermarkShape = sheet.Shapes.AddWordArt(
                                PresetWordArtStyle.WordArtStyle1,
                                watermarkText,
                                5,   // topRow (approximate)
                                0,   // top offset (pixels)
                                5,   // leftColumn (approximate)
                                0,   // left offset (pixels)
                                200, // height (pixels)
                                600  // width (pixels)
                            );

                            // Make semi‑transparent
                            watermarkShape.Fill.Transparency = 0.8; // 80% transparent

                            // Send behind other objects
                            watermarkShape.ZOrderPosition = -1;

                            // Rotate for diagonal appearance
                            watermarkShape.RotationAngle = -45;
                        }

                        // Determine output path
                        string fileName = Path.GetFileName(filePath);
                        string outputPath = Path.Combine(outputDir, fileName);

                        // Save modified workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Processed and saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch watermarking completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
