// Title: Batch convert SmartArt to GroupShape in Excel files using Aspose.Cells for .NET (C#)
// Description: Scans a folder for .xlsx workbooks, loads each with Aspose.Cells, walks every worksheet and shape, detects SmartArt objects, replaces them with GroupShape instances via GetResultOfSmartArt, and saves the updated files to an output directory with UpdateSmartArt enabled.
// Keywords: Aspose.Cells SmartArt conversion | C# batch Excel processing | GroupShape from SmartArt | convert SmartArt in bulk | Excel shape iteration Aspose | save workbook UpdateSmartArt
// Common Searches: how to replace SmartArt with GroupShape in C# | batch convert SmartArt in multiple Excel files | Aspose.Cells convert SmartArt to editable shapes | process all worksheets to change SmartArt | save Excel workbook after SmartArt conversion
// Developer Intent: Automatically replace every SmartArt diagram in all worksheets of each Excel file within a directory with a GroupShape and persist the changes.
// Use Cases: Standardize reports by turning SmartArt diagrams into editable GroupShapes before further styling. | Prepare workbooks for environments that lack SmartArt support, such as older Office versions or third‑party viewers. | Run a nightly job that cleans up a repository of financial models, ensuring all SmartArt is converted to static shapes for consistent rendering.
// AI Prompts: Write C# code that logs the original SmartArt name and the generated GroupShape name during batch conversion with Aspose.Cells. | Show how to add a semi‑transparent watermark to each GroupShape after converting SmartArt in a folder of workbooks. | Explain how to modify the save options to flatten SmartArt into raster images instead of GroupShapes when saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Scans a folder for .xlsx workbooks, loads each with Aspose.Cells, walks every worksheet and shape, detects SmartArt objects, replaces them with GroupShape instances via GetResultOfSmartArt, and saves the updated files to an output directory with UpdateSmartArt enabled.
class SmartArtBatchConverter
{
    static void Main()
    {
        // Input and output directories
        string inputFolder = @"C:\InputExcelFiles";
        string outputFolder = @"C:\ConvertedExcelFiles";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Verify the file still exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape is a SmartArt shape
                        if (shape.IsSmartArt)
                        {
                            // Convert SmartArt to a GroupShape
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            // Optional: manipulate the resulting groupShape here
                            // e.g., reposition it
                            // groupShape.Left += 10;
                            // groupShape.Top += 10;
                        }
                    }
                }

                // Prepare save options to persist the conversion
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true // ensures SmartArt conversion is saved
                };

                // Build output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Converted: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("SmartArt conversion completed for all files.");
    }
}
