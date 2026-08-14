// Title: Batch apply a tiled texture fill to a shape in multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Description: C# program that scans a folder for .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, finds a shape named "TargetShape" (or the first shape), sets its Fill.Texture to BlueTissuePaper with tiling enabled, and saves the updated files to a separate output directory while logging progress and errors.
// Keywords: Aspose.Cells | C# | .NET | batch process Excel workbooks | shape texture fill | tiled fill Aspose.Cells | TextureType.BlueTissuePaper | automate shape styling | Excel folder processing | shape fill API | bulk workbook modification
// Common Searches: How to batch apply a tiled texture to a shape in Excel files using Aspose.Cells | Aspose.Cells C# example for setting shape fill texture in multiple workbooks | Automate shape fill changes across a folder of spreadsheets | Apply texture fill to a specific shape in every Excel workbook | Save processed Excel files to a different directory with Aspose.Cells
// Developer Intent: Iterate through all Excel files in a directory and programmatically set a tiled texture fill on a designated shape within each workbook.
// Use Cases: Standardize the visual theme of a report suite by adding a tiled background to a logo shape in every file. | Prepare a batch of template workbooks where a placeholder shape receives a consistent tissue‑paper texture. | Migrate legacy spreadsheets to a new branding style by applying a tiled texture to the first shape in each workbook.
// AI Prompts: Generate C# code with Aspose.Cells that applies a tiled TextureType.Wood fill to a shape called "HeaderShape" in all .xlsx files of a given folder and saves them to an output path. | Show how to add robust error handling and logging when the target shape is missing during batch processing of Excel workbooks with Aspose.Cells. | Create a command‑line option to choose the texture type and toggle tiling for the batch shape‑fill operation using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchTextureProcessor
{
    // C# program that scans a folder for .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, finds a shape named "TargetShape" (or the first shape), sets its Fill.Texture to BlueTissuePaper with tiling enabled, and saves the updated files to a separate output directory while logging progress and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where modified workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all Excel files in the input folder (xlsx, xls, xlsm)
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in workbookFiles)
            {
                // Process only supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                    continue;

                // Skip if the file does not exist (safety check)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Assume the shape is on the first worksheet; adjust as needed
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Find the shape by name; replace "TargetShape" with the actual shape name
                    Shape targetShape = null;
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        if (shape.Name == "TargetShape")
                        {
                            targetShape = shape;
                            break;
                        }
                    }

                    // If the shape name is unknown, fallback to the first shape
                    if (targetShape == null && worksheet.Shapes.Count > 0)
                    {
                        targetShape = worksheet.Shapes[0];
                    }

                    if (targetShape != null)
                    {
                        // Apply a tiled texture using the new Fill API
                        targetShape.Fill.Texture = TextureType.BlueTissuePaper;
                        targetShape.Fill.TextureFill.IsTiling = true;
                    }
                    else
                    {
                        Console.WriteLine($"No shape found in workbook: {Path.GetFileName(filePath)}");
                    }

                    // Save the modified workbook to the output folder (preserving original name)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                    workbook.Dispose();

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
