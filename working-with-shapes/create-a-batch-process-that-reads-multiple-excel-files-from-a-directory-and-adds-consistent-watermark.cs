// Title: How to batch add a diagonal semi‑transparent text watermark to every worksheet in multiple .xlsx files using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory for .xlsx files, opens each workbook with Aspose.Cells, and inserts a semi‑transparent diagonal text‑effect shape labeled "CONFIDENTIAL" on every worksheet. | Adjust the watermark shape so it has 50 % transparency, -45° rotation, moves with cells, is sent to the back, and then save each modified workbook to a separate output folder.
// Common Searches: asp.net core batch watermark multiple excel files using aspose.cells | c# program to add diagonal text watermark to all sheets in a folder of xlsx workbooks | how to apply a semi transparent watermark to every worksheet with Aspose.Cells | automate adding CONFIDENTIAL watermark to many Excel workbooks in C#
// Tags: batch processing Excel workbooks with Aspose.Cells | insert text effect shape as watermark in .xlsx | configure shape transparency and rotation Aspose.Cells | move watermark with cells placement type | export watermarked workbooks to separate directory

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# console application iterates over all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, adds a semi‑transparent diagonal text‑effect shape labeled "CONFIDENTIAL" to every worksheet, and saves the watermarked workbooks to an output directory.
class BatchWatermarkProcessor
{
    static void Main(string[] args)
    {
        // Directory containing the Excel files to process
        string inputDirectory = @"C:\InputExcelFiles";
        // Directory where the processed files will be saved
        string outputDirectory = @"C:\OutputExcelFiles";

        // Verify input directory exists
        if (!Directory.Exists(inputDirectory))
        {
            Console.WriteLine($"Input directory does not exist: {inputDirectory}");
            return;
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        // Define the watermark text (consistent for all files)
        const string watermarkText = "CONFIDENTIAL";

        // Process each .xlsx file in the input directory
        foreach (string filePath in Directory.GetFiles(inputDirectory, "*.xlsx"))
        {
            // Guard against missing files (should not happen with GetFiles, but added for safety)
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found, skipping: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply watermark to every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Add a text effect shape that acts as a watermark
                        Shape watermarkShape = sheet.Shapes.AddTextEffect(
                            MsoPresetTextEffect.TextEffect1,
                            watermarkText,
                            "Arial",
                            72,
                            true,
                            false,
                            0,
                            0,
                            500,
                            200,
                            0,
                            0);

                        // Set shape formatting
                        // Note: ForeColor and IsVisible properties are not available in some versions;
                        // they are omitted to maintain compatibility.
                        watermarkShape.Fill.Transparency = 0.5; // 50% transparent
                        watermarkShape.RotationAngle = -45;    // Diagonal orientation
                        watermarkShape.Placement = PlacementType.Move; // Move with cells
                        watermarkShape.ZOrderPosition = 0;    // Send to back
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Failed to add watermark to sheet '{sheet.Name}': {shapeEx.Message}");
                    }
                }

                // Build the output file path (same name, different folder)
                string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));

                // Save the modified workbook (overwrites if file exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch watermarking completed.");
    }
}
