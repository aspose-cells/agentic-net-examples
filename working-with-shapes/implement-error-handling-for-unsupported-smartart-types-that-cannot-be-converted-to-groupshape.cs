using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

namespace AsposeCellsSmartArtHandling
{
    public class SmartArtConversionDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that may contain SmartArt shapes
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        try
                        {
                            // Convert the SmartArt to a GroupShape
                            GroupShape groupShape = shape.GetResultOfSmartArt();

                            // If conversion returns null, the SmartArt type is unsupported
                            if (groupShape == null)
                            {
                                Console.WriteLine($"SmartArt shape (Id={shape.Id}) cannot be converted to GroupShape.");
                                continue;
                            }

                            // Example processing: move the resulting group shape
                            groupShape.Left += 50;
                            groupShape.Top += 20;
                            Console.WriteLine($"SmartArt shape (Id={shape.Id}) converted and repositioned.");
                        }
                        catch (Exception ex)
                        {
                            // Handle any unexpected errors during conversion
                            Console.WriteLine($"Error converting SmartArt shape (Id={shape.Id}): {ex.Message}");
                        }
                    }
                }
            }

            // Save the workbook with SmartArt updates enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateSmartArt = true
            };
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}