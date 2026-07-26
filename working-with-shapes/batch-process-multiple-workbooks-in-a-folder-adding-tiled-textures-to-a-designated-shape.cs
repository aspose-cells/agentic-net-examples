// Title: Batch apply a tiled texture to a named shape in multiple Excel files with Aspose.Cells (C#)
// Description: Iterates through every .xlsx file in a source folder, loads each workbook with Aspose.Cells, finds the shape called "TargetShape" on every worksheet, sets its Fill.Texture to TextureType.BlueTissuePaper (default tiled alignment), and saves the updated workbook to a separate output directory.
// Keywords: Aspose.Cells | C# | batch process Excel workbooks | shape fill texture | tiled texture | BlueTissuePaper | multiple .xlsx files | folder automation | Excel shape styling | programmatic texture assignment
// Common Searches: apply tiled texture to shape in many Excel files using Aspose.Cells | C# batch update shape fill for all worksheets in a folder | Aspose.Cells set texture for a specific shape across workbooks | automate shape formatting in multiple .xlsx files | how to loop through Excel files and change shape fill with Aspose
// Developer Intent: Automatically add a tiled texture to the shape named "TargetShape" in every worksheet of each workbook located in a given folder and write the modified files to an output directory.
// Use Cases: Standardize the appearance of a logo shape across a suite of generated reports. | Prepare template workbooks with a decorative background before distribution to clients. | Ensure a specific shape has a tiled texture for consistent printing results in bulk Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to scan all .xlsx files in a directory, locate a shape named "TargetShape" on each sheet, set its Fill.Texture to TextureType.BlueTissuePaper, and save the workbooks to a separate folder. | Show how to modify the script to apply a tiled texture to every shape whose name starts with "Header_" in each workbook processed. | Explain how to change the texture type to TextureType.Wood while keeping the tiled alignment for the target shape in a batch operation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Iterates through every .xlsx file in a source folder, loads each workbook with Aspose.Cells, finds the shape called "TargetShape" on every worksheet, sets its Fill.Texture to TextureType.BlueTissuePaper (default tiled alignment), and saves the updated workbook to a separate output directory.
class Program
{
    static void Main()
    {
        // Folder containing source workbooks
        string sourceFolder = @"C:\InputWorkbooks";
        // Folder where modified workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Verify the source file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Source file not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Find the shape named "TargetShape"
                    Shape targetShape = null;
                    foreach (Shape shape in sheet.Shapes)
                    {
                        if (shape.Name == "TargetShape")
                        {
                            targetShape = shape;
                            break;
                        }
                    }

                    // If the shape exists, apply a tiled texture
                    if (targetShape != null)
                    {
                        // Apply texture (Fill property returns FillFormat)
                        targetShape.Fill.Texture = TextureType.BlueTissuePaper;

                        // Note: TextureAlignment property is not available in the current Aspose.Cells version.
                        // The default alignment is Tile, which satisfies the requirement.
                    }
                }

                // Build the output file path
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the modified workbook
                workbook.Save(outputPath);
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
