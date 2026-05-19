using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class BatchTextureProcessor
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Folder where the modified workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";
        Directory.CreateDirectory(outputFolder);

        // Ensure the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Verify the file still exists before loading
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook inside a using block for proper disposal
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Access the first worksheet (adjust if needed)
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Ensure the worksheet contains at least one shape
                    if (worksheet.Shapes.Count > 0)
                    {
                        // Get the first shape
                        Shape shape = worksheet.Shapes[0];

                        // Apply a texture fill (tiling is the default behavior)
                        shape.Fill.Texture = TextureType.BlueTissuePaper;
                    }
                    else
                    {
                        // Add a rectangle shape to demonstrate the texture
                        Shape newShape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 150, 200);
                        newShape.Fill.Texture = TextureType.BlueTissuePaper;
                    }

                    // Save the modified workbook to the output folder (overwrites if exists)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}