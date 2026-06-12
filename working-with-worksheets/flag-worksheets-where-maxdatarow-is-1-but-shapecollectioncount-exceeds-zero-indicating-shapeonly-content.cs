using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook while preserving all shapes
            var loadOptions = new LoadOptions
            {
                // Keep all shapes for accurate detection
                IgnoreUselessShapes = false
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Examine each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow == -1 means no data cells
                int maxDataRow = sheet.Cells.MaxDataRow;
                int shapeCount = sheet.Shapes.Count;

                if (maxDataRow == -1 && shapeCount > 0)
                {
                    // Flag the sheet by adding a custom property (store bool as string)
                    sheet.CustomProperties.Add("ShapeOnlyContent", "true");
                    Console.WriteLine($"Flagged sheet: {sheet.Name} (Shapes={shapeCount})");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}