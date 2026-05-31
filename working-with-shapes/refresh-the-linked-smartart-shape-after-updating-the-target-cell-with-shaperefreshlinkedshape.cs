using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    class RefreshSmartArtDemo
    {
        static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Update the cell linked to the SmartArt
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["B2"].PutValue("Updated Value");

                // Iterate through shapes; no explicit refresh method is required
                // because UpdateSmartArt = true in save options will refresh linked data.
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        // Placeholder for any future SmartArt-specific handling
                    }
                }

                // Save with SmartArt update enabled
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    UpdateSmartArt = true
                };
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}